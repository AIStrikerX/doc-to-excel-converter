"""
Validation module for DOC to Excel Converter

This module handles validation of parsed questions and data integrity checks.

Author: Your Name
Version: 1.0.0
"""

import re
from typing import List, Dict, Tuple, Optional, Set
from dataclasses import dataclass

from .config import Config
from .utils import is_empty_or_whitespace


@dataclass
class ValidationError:
    """Represents a validation error."""
    question_id: str
    error_type: str
    message: str
    severity: str = "error"  # error, warning, info


class QuestionValidator:
    """
    Validator for MCQ questions extracted from documents.
    
    Performs various validation checks to ensure data quality and completeness.
    """
    
    def __init__(self, config: Config):
        self.config = config
        self.valid_answer_values = {"1", "2", "3", "4", "5"}
        self.valid_answer_letters = {"A", "B", "C", "D", "E"}
    
    def validate_questions(self, questions: List[Dict[str, str]]) -> Tuple[List[Dict[str, str]], List[str]]:
        """
        Validate a list of questions and return valid ones with error messages.
        
        Args:
            questions: List of question dictionaries to validate
            
        Returns:
            Tuple of (valid_questions, error_messages)
        """
        valid_questions = []
        errors = []
        
        for question in questions:
            validation_errors = self.validate_single_question(question)
            
            if validation_errors:
                # Check if any errors are critical (not just warnings)
                critical_errors = [e for e in validation_errors if e.severity == "error"]
                
                if not critical_errors:
                    # Only warnings, include the question
                    valid_questions.append(question)
                
                # Collect all error messages
                for error in validation_errors:
                    error_msg = f"Question {question.get('MCQ_NO', 'Unknown')}: {error.message}"
                    errors.append(error_msg)
            else:
                valid_questions.append(question)
        
        return valid_questions, errors
    
    def validate_single_question(self, question: Dict[str, str]) -> List[ValidationError]:
        """
        Validate a single question and return list of validation errors.
        
        Args:
            question: Question dictionary to validate
            
        Returns:
            List of validation errors
        """
        errors = []
        question_id = question.get('MCQ_NO', 'Unknown')
        
        # Required field validation
        errors.extend(self._validate_required_fields(question, question_id))
        
        # Question text validation
        errors.extend(self._validate_question_text(question, question_id))
        
        # Options validation
        errors.extend(self._validate_options(question, question_id))
        
        # Correct answer validation
        errors.extend(self._validate_correct_answer(question, question_id))
        
        # Explanation validation
        errors.extend(self._validate_explanations(question, question_id))
        
        # Cross-field validation
        errors.extend(self._validate_consistency(question, question_id))
        
        return errors
    
    def _validate_required_fields(self, question: Dict[str, str], question_id: str) -> List[ValidationError]:
        """Validate that required fields are present and not empty."""
        errors = []
        
        required_fields = ["MCQ_NO", "Question", "CorAns"]
        
        for field in required_fields:
            if field not in question:
                errors.append(ValidationError(
                    question_id, "missing_field", 
                    f"Required field '{field}' is missing"
                ))
            elif is_empty_or_whitespace(question[field]):
                errors.append(ValidationError(
                    question_id, "empty_field",
                    f"Required field '{field}' is empty"
                ))
        
        return errors
    
    def _validate_question_text(self, question: Dict[str, str], question_id: str) -> List[ValidationError]:
        """Validate question text content."""
        errors = []
        
        question_text = question.get("Question", "")
        
        if question_text:
            # Check minimum length
            if len(question_text.strip()) < 10:
                errors.append(ValidationError(
                    question_id, "question_too_short",
                    "Question text is too short (less than 10 characters)",
                    severity="warning"
                ))
            
            # Check for common formatting issues
            if question_text.strip().endswith("?") == False and "?" not in question_text:
                errors.append(ValidationError(
                    question_id, "missing_question_mark",
                    "Question text doesn't seem to be a question (no question mark found)",
                    severity="warning"
                ))
            
            # Check for placeholder text
            placeholder_patterns = [
                r"\[.*?\]",  # [placeholder]
                r"xxx+",     # xxx or xxxx
                r"___+",     # blank spaces
            ]
            
            for pattern in placeholder_patterns:
                if re.search(pattern, question_text, re.I):
                    errors.append(ValidationError(
                        question_id, "placeholder_text",
                        "Question text appears to contain placeholder content",
                        severity="warning"
                    ))
                    break
        
        return errors
    
    def _validate_options(self, question: Dict[str, str], question_id: str) -> List[ValidationError]:
        """Validate multiple choice options."""
        errors = []
        
        # Get all options
        options = {}
        for i in range(1, 6):
            option_key = f"Answer{i}"
            option_text = question.get(option_key, "").strip()
            if option_text:
                options[chr(64 + i)] = option_text  # Convert 1->A, 2->B, etc.
        
        # Must have at least 2 options
        if len(options) < 2:
            errors.append(ValidationError(
                question_id, "insufficient_options",
                f"Question has only {len(options)} options (minimum 2 required)"
            ))
        
        # Check for duplicate options
        option_texts = list(options.values())
        unique_texts = set(option_texts)
        
        if len(option_texts) != len(unique_texts):
            errors.append(ValidationError(
                question_id, "duplicate_options",
                "Question has duplicate options",
                severity="warning"
            ))
        
        # Check option length and quality
        for letter, text in options.items():
            if len(text) < 2:
                errors.append(ValidationError(
                    question_id, "option_too_short",
                    f"Option {letter} is too short",
                    severity="warning"
                ))
            
            # Check for placeholder content in options
            if re.search(r"(option|answer|choice)\s*[a-e]", text, re.I):
                errors.append(ValidationError(
                    question_id, "placeholder_option",
                    f"Option {letter} appears to be placeholder text",
                    severity="warning"
                ))
        
        return errors
    
    def _validate_correct_answer(self, question: Dict[str, str], question_id: str) -> List[ValidationError]:
        """Validate the correct answer field."""
        errors = []
        
        correct_answer = question.get("CorAns", "").strip()
        
        if correct_answer:
            # Must be valid answer value (1-5)
            if correct_answer not in self.valid_answer_values:
                errors.append(ValidationError(
                    question_id, "invalid_answer_format",
                    f"Correct answer '{correct_answer}' is not valid (must be 1-5)"
                ))
            else:
                # Check if corresponding option exists
                option_key = f"Answer{correct_answer}"
                if is_empty_or_whitespace(question.get(option_key, "")):
                    errors.append(ValidationError(
                        question_id, "missing_correct_option",
                        f"Correct answer points to option {correct_answer} but that option is empty"
                    ))
        
        return errors
    
    def _validate_explanations(self, question: Dict[str, str], question_id: str) -> List[ValidationError]:
        """Validate explanation fields."""
        errors = []
        
        correct_explanation = question.get("CorrectExplanation", "").strip()
        incorrect_explanation = question.get("IncorrectExplanation", "").strip()
        
        # Check if explanations exist
        if not correct_explanation and not incorrect_explanation:
            errors.append(ValidationError(
                question_id, "missing_explanations",
                "No explanations provided for this question",
                severity="warning"
            ))
        
        # Check explanation length and quality
        for exp_type, exp_text in [("CorrectExplanation", correct_explanation), 
                                  ("IncorrectExplanation", incorrect_explanation)]:
            if exp_text:
                if len(exp_text) < 10:
                    errors.append(ValidationError(
                        question_id, "explanation_too_short",
                        f"{exp_type} is very short (less than 10 characters)",
                        severity="warning"
                    ))
                
                # Check for placeholder content
                if re.search(r"(explanation|rationale|because)", exp_text, re.I) and len(exp_text) < 20:
                    errors.append(ValidationError(
                        question_id, "placeholder_explanation",
                        f"{exp_type} appears to be placeholder text",
                        severity="warning"
                    ))
        
        return errors
    
    def _validate_consistency(self, question: Dict[str, str], question_id: str) -> List[ValidationError]:
        """Validate cross-field consistency."""
        errors = []
        
        # Check if question number matches expected format
        mcq_no = question.get("MCQ_NO", "").strip()
        if mcq_no and not re.match(r"^\d+$", mcq_no):
            errors.append(ValidationError(
                question_id, "invalid_question_number",
                f"Question number '{mcq_no}' is not a valid integer",
                severity="warning"
            ))
        
        # Check source file field
        source_file = question.get("SourceFile", "").strip()
        if source_file and not source_file.lower().endswith('.docx'):
            errors.append(ValidationError(
                question_id, "invalid_source_file",
                f"Source file '{source_file}' doesn't appear to be a .docx file",
                severity="warning"
            ))
        
        return errors
    
    def validate_batch_consistency(self, questions: List[Dict[str, str]]) -> List[str]:
        """
        Validate consistency across multiple questions in a batch.
        
        Args:
            questions: List of questions to validate
            
        Returns:
            List of validation error messages
        """
        errors = []
        
        if not questions:
            return errors
        
        # Check for duplicate question numbers
        question_numbers = [q.get("MCQ_NO", "") for q in questions]
        seen_numbers = set()
        
        for num in question_numbers:
            if num and num in seen_numbers:
                errors.append(f"Duplicate question number found: {num}")
            elif num:
                seen_numbers.add(num)
        
        # Check for sequential numbering (warning only)
        valid_numbers = [int(num) for num in question_numbers if num.isdigit()]
        if valid_numbers:
            valid_numbers.sort()
            expected_sequence = list(range(1, len(valid_numbers) + 1))
            
            if valid_numbers != expected_sequence:
                errors.append("Question numbers are not sequential (this may be intentional)")
        
        # Check for consistent metadata
        groups = set(q.get("Group", "") for q in questions)
        types = set(q.get("Type", "") for q in questions)
        
        if len(groups) > 1:
            errors.append(f"Multiple groups found in batch: {', '.join(groups)}")
        
        if len(types) > 1:
            errors.append(f"Multiple question types found in batch: {', '.join(types)}")
        
        return errors
    
    def get_validation_summary(self, questions: List[Dict[str, str]]) -> Dict[str, int]:
        """
        Get validation summary statistics.
        
        Args:
            questions: List of questions to analyze
            
        Returns:
            Dictionary with validation statistics
        """
        summary = {
            "total_questions": len(questions),
            "valid_questions": 0,
            "questions_with_warnings": 0,
            "questions_with_errors": 0,
            "missing_explanations": 0,
            "incomplete_options": 0
        }
        
        for question in questions:
            validation_errors = self.validate_single_question(question)
            
            has_errors = any(e.severity == "error" for e in validation_errors)
            has_warnings = any(e.severity == "warning" for e in validation_errors)
            
            if not has_errors:
                summary["valid_questions"] += 1
            else:
                summary["questions_with_errors"] += 1
            
            if has_warnings:
                summary["questions_with_warnings"] += 1
            
            # Check specific issues
            if is_empty_or_whitespace(question.get("CorrectExplanation", "")):
                summary["missing_explanations"] += 1
            
            option_count = sum(1 for i in range(1, 6) 
                             if not is_empty_or_whitespace(question.get(f"Answer{i}", "")))
            if option_count < 3:
                summary["incomplete_options"] += 1
        
        return summary