"""
Document parsing module for DOC to Excel Converter

This module handles the parsing of Word documents to extract
MCQ questions, options, answers, and explanations.

Author: Your Name
Version: 1.0.0
"""

import os
import re
from typing import List, Dict, Tuple, Optional, Union
from abc import ABC, abstractmethod

from docx import Document
from docx.text.run import Run

from .config import Config
from .utils import normalize_text


class BaseParser(ABC):
    """Abstract base class for document parsers."""
    
    def __init__(self, config: Config):
        self.config = config
        self._compile_patterns()
    
    def _compile_patterns(self) -> None:
        """Compile regex patterns for efficiency."""
        self.question_patterns = [re.compile(p, re.I) for p in self.config.question_patterns]
        self.option_patterns = [re.compile(p, re.S) for p in self.config.option_patterns]
        self.answer_patterns = [re.compile(p, re.I) for p in self.config.answer_patterns]
        self.explanation_patterns = [re.compile(p, re.I) for p in self.config.explanation_patterns]
        
        # Section stop markers
        self.section_stop_patterns = [
            re.compile(r"^\s*Analysis\s+of\s+Other\s+Options", re.I),
            re.compile(r"^\s*Analysis\s+of\s+Distractors", re.I),
            re.compile(r"^\s*Distractors", re.I),
            re.compile(r"^\s*Key\s+Insights?", re.I),
        ]
    
    @abstractmethod
    def parse_document(self, document: Document, file_path: str) -> List[Dict[str, str]]:
        """Parse a Word document and extract questions."""
        pass


class QuestionParser(BaseParser):
    """
    Parser for extracting MCQ questions from Word documents.
    
    Handles various question formats and extracts:
    - Question numbers and topics
    - Question text/stem
    - Multiple choice options (A-E)
    - Correct answers
    - Explanations (highlighted vs normal text)
    """
    
    # Answer letter to number mapping
    LETTER_TO_NUM = {"A": "1", "B": "2", "C": "3", "D": "4", "E": "5"}
    NUM_TO_LETTER = {v: k for k, v in LETTER_TO_NUM.items()}
    
    def parse_document(self, document: Document, file_path: str) -> List[Dict[str, str]]:
        """
        Parse a Word document and extract all questions.
        
        Args:
            document: python-docx Document object
            file_path: Path to the source file
            
        Returns:
            List of question dictionaries
        """
        paragraphs = document.paragraphs
        question_blocks = self._find_question_blocks(paragraphs)
        
        if not question_blocks:
            return []
        
        questions = []
        for start_idx, end_idx, question_num, topic in question_blocks:
            try:
                question_data = self._parse_question_block(
                    paragraphs, start_idx, end_idx, question_num, topic, file_path
                )
                
                if self._is_valid_question(question_data):
                    questions.append(question_data)
                    
            except Exception as e:
                print(f"Error parsing question {question_num} in {file_path}: {e}")
                continue
        
        return questions
    
    def _find_question_blocks(self, paragraphs) -> List[Tuple[int, int, str, str]]:
        """
        Find question blocks in the document.
        
        Args:
            paragraphs: List of document paragraphs
            
        Returns:
            List of tuples (start_idx, end_idx, question_num, topic)
        """
        blocks = []
        
        for i, para in enumerate(paragraphs):
            text = normalize_text(para.text)
            
            for pattern in self.question_patterns:
                match = pattern.match(text)
                if match:
                    question_num = match.group(1)
                    topic = match.group(2) if len(match.groups()) > 1 and match.group(2) else ""
                    
                    # Find end of this question block
                    end_idx = len(paragraphs)
                    for j in range(i + 1, len(paragraphs)):
                        next_text = normalize_text(paragraphs[j].text)
                        for next_pattern in self.question_patterns:
                            if next_pattern.match(next_text):
                                end_idx = j
                                break
                        if end_idx != len(paragraphs):
                            break
                    
                    blocks.append((i, end_idx, question_num, topic))
                    break
        
        return blocks
    
    def _parse_question_block(self, paragraphs, start_idx: int, end_idx: int, 
                            question_num: str, topic: str, source_file: str) -> Dict[str, str]:
        """
        Parse a single question block.
        
        Args:
            paragraphs: Document paragraphs
            start_idx: Start index of question block
            end_idx: End index of question block  
            question_num: Question number
            topic: Question topic
            source_file: Source file path
            
        Returns:
            Dictionary containing parsed question data
        """
        # Skip the header line
        content_start = start_idx + 1
        
        # Find where options start
        options_start_idx = self._find_options_start(paragraphs, content_start, end_idx)
        
        # Parse question stem
        question_stem = self._parse_question_stem(paragraphs, content_start, options_start_idx)
        
        # Parse options
        options, options_end_idx = self._parse_options_section(paragraphs, options_start_idx, end_idx)
        
        # Find correct answer
        correct_answer, answer_end_idx = self._find_correct_answer(paragraphs, options_end_idx, end_idx)
        
        # Parse explanation
        explanation_start_idx = max(options_end_idx, answer_end_idx)
        correct_explanation, incorrect_explanation = self._parse_explanation_section(
            paragraphs, explanation_start_idx, end_idx
        )
        
        # Build result
        result = {
            "MCQ_NO": question_num,
            "Question": question_stem,
            "Answer1": options.get("A", ""),
            "Answer2": options.get("B", ""),
            "Answer3": options.get("C", ""),
            "Answer4": options.get("D", ""),
            "Answer5": options.get("E", ""),
            "CorAns": self.LETTER_TO_NUM.get(correct_answer, "") if correct_answer else "",
            "CorrectExplanation": correct_explanation,
            "IncorrectExplanation": incorrect_explanation or correct_explanation,
            "Topic": topic,
            "SourceFile": os.path.basename(source_file),
            "Group": self.config.group,
            "Type": self.config.question_type,
        }
        
        return result
    
    def _find_options_start(self, paragraphs, start_idx: int, end_idx: int) -> int:
        """Find the index where options start."""
        for i in range(start_idx, end_idx):
            text = normalize_text(paragraphs[i].text)
            for pattern in self.option_patterns:
                if pattern.match(text):
                    return i
        return start_idx
    
    def _parse_question_stem(self, paragraphs, start_idx: int, end_idx: int) -> str:
        """Parse the question stem/text."""
        stem_parts = []
        
        for i in range(start_idx, end_idx):
            text = normalize_text(paragraphs[i].text)
            if text:
                stem_parts.append(text)
        
        return " ".join(stem_parts)
    
    def _parse_options_section(self, paragraphs, start_idx: int, end_idx: int) -> Tuple[Dict[str, str], int]:
        """
        Parse the options section.
        
        Returns:
            Tuple of (options_dict, options_end_index)
        """
        options = {}
        current_option = None
        current_text_parts = []
        options_end_idx = end_idx
        
        def save_current_option():
            if current_option and current_text_parts:
                options[current_option] = normalize_text(" ".join(current_text_parts))
        
        for i in range(start_idx, end_idx):
            text = normalize_text(paragraphs[i].text)
            if not text:
                continue
            
            # Check for section boundaries
            if self._is_section_boundary(text):
                options_end_idx = i
                break
            
            # Check if this is an option line
            option_match = None
            for pattern in self.option_patterns:
                option_match = pattern.match(text)
                if option_match:
                    break
            
            if option_match:
                save_current_option()
                current_option = option_match.group(1).upper()
                option_text = option_match.group(2).strip() if len(option_match.groups()) > 1 else ""
                current_text_parts = [option_text] if option_text else []
            else:
                # Continuation text for current option
                if current_option:
                    current_text_parts.append(text)
        
        save_current_option()
        return options, options_end_idx
    
    def _find_correct_answer(self, paragraphs, start_idx: int, end_idx: int) -> Tuple[Optional[str], int]:
        """Find the correct answer line."""
        for i in range(start_idx, end_idx):
            text = normalize_text(paragraphs[i].text)
            
            for pattern in self.answer_patterns:
                match = pattern.match(text)
                if match:
                    answer = match.group(1).upper()
                    # Convert number to letter if needed
                    if answer.isdigit() and answer in self.NUM_TO_LETTER:
                        answer = self.NUM_TO_LETTER[answer]
                    return answer, i + 1
        
        return None, start_idx
    
    def _parse_explanation_section(self, paragraphs, start_idx: int, end_idx: int) -> Tuple[str, str]:
        """
        Parse explanation section, separating highlighted from normal text.
        
        Returns:
            Tuple of (correct_explanation, incorrect_explanation)
        """
        # Find explanation header
        explanation_start = start_idx
        for i in range(start_idx, end_idx):
            text = normalize_text(paragraphs[i].text)
            for pattern in self.explanation_patterns:
                if pattern.match(text):
                    explanation_start = i + 1
                    break
        
        # Find end of explanation section
        explanation_end = end_idx
        for i in range(explanation_start, end_idx):
            text = normalize_text(paragraphs[i].text)
            if self._is_section_boundary(text):
                explanation_end = i
                break
        
        # Extract highlighted vs normal text
        highlighted_parts = []
        normal_parts = []
        
        for i in range(explanation_start, explanation_end):
            highlighted, normal = self._extract_highlighted_text(paragraphs[i])
            if highlighted:
                highlighted_parts.append(highlighted)
            if normal:
                normal_parts.append(normal)
        
        highlighted_text = " ".join(highlighted_parts).strip()
        normal_text = " ".join(normal_parts).strip()
        
        # If no highlighting detected, treat all as correct explanation
        if not highlighted_text and normal_text:
            highlighted_text = normal_text
            normal_text = ""
        
        return highlighted_text, normal_text
    
    def _extract_highlighted_text(self, paragraph) -> Tuple[str, str]:
        """
        Extract highlighted and normal text from a paragraph.
        
        Returns:
            Tuple of (highlighted_text, normal_text)
        """
        highlighted_parts = []
        normal_parts = []
        
        for run in paragraph.runs:
            if not run.text:
                continue
            
            if self._is_highlighted(run):
                highlighted_parts.append(run.text)
            else:
                normal_parts.append(run.text)
        
        highlighted_text = normalize_text(" ".join(highlighted_parts))
        normal_text = normalize_text(" ".join(normal_parts))
        
        # If no highlighting detected, return all text as normal
        if not highlighted_text:
            return "", normalize_text(paragraph.text)
        
        return highlighted_text, normal_text
    
    def _is_highlighted(self, run) -> bool:
        """Check if a text run is highlighted."""
        try:
            # Check for highlight color
            if hasattr(run.font, 'highlight_color') and run.font.highlight_color is not None:
                return True
            
            # Check for XML highlight elements
            if hasattr(run, '_element'):
                highlight = run._element.xpath('.//w:highlight')
                if highlight:
                    return True
                    
        except Exception:
            pass
        
        return False
    
    def _is_section_boundary(self, text: str) -> bool:
        """Check if text represents a section boundary."""
        # Check for answer patterns
        for pattern in self.answer_patterns:
            if pattern.match(text):
                return True
        
        # Check for explanation headers
        for pattern in self.explanation_patterns:
            if pattern.match(text):
                return True
        
        # Check for section stop markers
        for pattern in self.section_stop_patterns:
            if pattern.match(text):
                return True
        
        return False
    
    def _is_valid_question(self, question_data: Dict[str, str]) -> bool:
        """Check if parsed question data is valid."""
        # Must have question text
        if not question_data.get("Question", "").strip():
            return False
        
        # Must have at least one option
        options = [question_data.get(f"Answer{i}", "") for i in range(1, 6)]
        if not any(opt.strip() for opt in options):
            return False
        
        # Should have correct answer if validation is enabled
        if self.config.validate_answers and not question_data.get("CorAns", "").strip():
            return False
        
        return True


class EnhancedMedicalParser(QuestionParser):
    """
    Enhanced parser specifically designed for medical question formats.
    
    Includes additional patterns and handling for medical-specific content.
    """
    
    def __init__(self, config: Config):
        super().__init__(config)
        
        # Additional medical-specific patterns
        self.medical_question_patterns = [
            re.compile(r"^\s*Case\s+(\d+)(?:\s*[-–—]\s*(.+))?\s*$", re.I),
            re.compile(r"^\s*Scenario\s+(\d+)(?:\s*[-–—]\s*(.+))?\s*$", re.I),
        ]
        
        self.question_patterns.extend(self.medical_question_patterns)
    
    def _parse_question_stem(self, paragraphs, start_idx: int, end_idx: int) -> str:
        """Enhanced question stem parsing for medical content."""
        stem_parts = []
        
        for i in range(start_idx, end_idx):
            text = normalize_text(paragraphs[i].text)
            if text:
                # Clean medical-specific formatting
                text = self._clean_medical_text(text)
                stem_parts.append(text)
        
        return " ".join(stem_parts)
    
    def _clean_medical_text(self, text: str) -> str:
        """Clean medical-specific text formatting."""
        # Remove common medical formatting artifacts
        text = re.sub(r'\b(mg|mcg|mL|kg|mmHg)\b', lambda m: m.group().replace(' ', ''), text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()