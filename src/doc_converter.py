"""
DOC to Excel Converter - Core Module

This module contains the main DocToExcelConverter class that handles
the conversion of Word documents containing MCQs to Excel format.

Author: Your Name
Version: 1.0.0
"""

import os
import re
from typing import List, Dict, Tuple, Optional, Union
from pathlib import Path

import pandas as pd
from docx import Document
from docx.text.run import Run

from .config import Config
from .parsers import QuestionParser
from .validators import QuestionValidator
from .utils import normalize_text, setup_logging


class DocToExcelConverter:
    """
    Main converter class for processing Word documents containing MCQs
    and converting them to structured Excel files.
    """
    
    def __init__(self, config: Optional[Config] = None, parser: Optional[QuestionParser] = None):
        """
        Initialize the converter with optional configuration and parser.
        
        Args:
            config: Configuration object with converter settings
            parser: Custom parser for question extraction
        """
        self.config = config or Config()
        self.parser = parser or QuestionParser(self.config)
        self.validator = QuestionValidator(self.config)
        self.logger = setup_logging(self.config.verbose, self.config.quiet)
        
        # Statistics
        self.stats = {
            'files_processed': 0,
            'questions_extracted': 0,
            'questions_valid': 0,
            'errors': []
        }
    
    def convert_file(self, input_path: str, output_path: str, **kwargs) -> pd.DataFrame:
        """
        Convert a single .docx file to Excel format.
        
        Args:
            input_path: Path to the input .docx file
            output_path: Path for the output Excel file
            **kwargs: Additional parameters to override config
            
        Returns:
            DataFrame containing the processed questions
            
        Raises:
            FileNotFoundError: If input file doesn't exist
            ValueError: If file format is not supported
        """
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        if not input_path.lower().endswith('.docx'):
            raise ValueError(f"Unsupported file format. Expected .docx, got: {input_path}")
        
        self.logger.info(f"Converting file: {input_path}")
        
        # Parse the document
        questions = self.parse_docx(input_path)
        
        if not questions:
            self.logger.warning(f"No questions found in {input_path}")
            return pd.DataFrame()
        
        # Update metadata with kwargs
        group = kwargs.get('group', self.config.group)
        question_type = kwargs.get('type', self.config.question_type)
        
        for question in questions:
            question['Group'] = group
            question['Type'] = question_type
        
        # Create DataFrame and save to Excel
        df = self._create_dataframe(questions)
        self._save_to_excel(df, output_path)
        
        # Update statistics
        self.stats['files_processed'] = 1
        self.stats['questions_extracted'] = len(questions)
        self.stats['questions_valid'] = len(df)
        
        self.logger.info(f"Successfully converted {len(df)} questions to {output_path}")
        return df
    
    def convert_directory(self, input_path: str, output_path: str, **kwargs) -> pd.DataFrame:
        """
        Convert all .docx files in a directory to a single Excel file.
        
        Args:
            input_path: Path to the directory containing .docx files
            output_path: Path for the output Excel file
            **kwargs: Additional parameters to override config
            
        Returns:
            DataFrame containing all processed questions
            
        Raises:
            NotADirectoryError: If input path is not a directory
            FileNotFoundError: If no .docx files found in directory
        """
        if not os.path.isdir(input_path):
            raise NotADirectoryError(f"Input path is not a directory: {input_path}")
        
        # Find all .docx files
        pattern = kwargs.get('pattern', '*.docx')
        docx_files = list(Path(input_path).glob(pattern))
        
        # Filter out temporary files
        docx_files = [f for f in docx_files if not f.name.startswith('~')]
        
        if not docx_files:
            raise FileNotFoundError(f"No .docx files found in {input_path}")
        
        max_files = kwargs.get('max_files')
        if max_files:
            docx_files = docx_files[:max_files]
        
        self.logger.info(f"Processing {len(docx_files)} files from {input_path}")
        
        all_questions = []
        processed_files = 0
        
        # Process each file
        for file_path in docx_files:
            try:
                self.logger.info(f"Processing: {file_path}")
                questions = self.parse_docx(str(file_path))
                
                if questions:
                    # Update metadata
                    group = kwargs.get('group', self.config.group)
                    question_type = kwargs.get('type', self.config.question_type)
                    
                    for question in questions:
                        question['Group'] = group
                        question['Type'] = question_type
                    
                    all_questions.extend(questions)
                    processed_files += 1
                    
                    self.logger.info(f"Extracted {len(questions)} questions from {file_path}")
                else:
                    self.logger.warning(f"No questions found in {file_path}")
                    
            except Exception as e:
                error_msg = f"Error processing {file_path}: {str(e)}"
                self.logger.error(error_msg)
                self.stats['errors'].append(error_msg)
                continue
        
        if not all_questions:
            self.logger.warning("No questions extracted from any files")
            return pd.DataFrame()
        
        # Create DataFrame and save to Excel
        df = self._create_dataframe(all_questions)
        self._save_to_excel(df, output_path)
        
        # Update statistics
        self.stats['files_processed'] = processed_files
        self.stats['questions_extracted'] = len(all_questions)
        self.stats['questions_valid'] = len(df)
        
        self.logger.info(f"Successfully processed {processed_files} files with {len(df)} questions")
        return df
    
    def parse_docx(self, file_path: str) -> List[Dict[str, str]]:
        """
        Parse a single .docx file and extract questions.
        
        Args:
            file_path: Path to the .docx file
            
        Returns:
            List of question dictionaries
        """
        try:
            doc = Document(file_path)
            questions = self.parser.parse_document(doc, file_path)
            
            # Validate questions if enabled
            if self.config.validate_answers:
                valid_questions, validation_errors = self.validator.validate_questions(questions)
                
                if validation_errors:
                    for error in validation_errors:
                        self.logger.warning(f"Validation error in {file_path}: {error}")
                        self.stats['errors'].append(error)
                
                return valid_questions
            
            return questions
            
        except Exception as e:
            error_msg = f"Error parsing {file_path}: {str(e)}"
            self.logger.error(error_msg)
            self.stats['errors'].append(error_msg)
            return []
    
    def _create_dataframe(self, questions: List[Dict[str, str]]) -> pd.DataFrame:
        """
        Create a pandas DataFrame from the list of questions.
        
        Args:
            questions: List of question dictionaries
            
        Returns:
            DataFrame with standardized columns and clean data
        """
        if not questions:
            return pd.DataFrame()
        
        # Define standard column order
        columns = [
            "MCQ_NO",
            "Question", 
            "Answer1",
            "CorAns",
            "Answer2",
            "Answer3", 
            "Answer4",
            "Answer5",
            "CorrectExplanation",
            "IncorrectExplanation",
            "Topic",
            "Group",
            "Type",
            "SourceFile"
        ]
        
        # Create DataFrame
        df = pd.DataFrame(questions)
        
        # Ensure all required columns exist
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        
        # Reorder columns
        df = df.reindex(columns=columns)
        
        # Clean text data
        text_columns = [
            "Question", "Answer1", "Answer2", "Answer3", "Answer4", "Answer5",
            "CorrectExplanation", "IncorrectExplanation", "Topic"
        ]
        
        for col in text_columns:
            if col in df.columns:
                df[col] = df[col].astype(str).apply(normalize_text)
        
        # Fill empty IncorrectExplanation with CorrectExplanation if configured
        if "CorrectExplanation" in df.columns and "IncorrectExplanation" in df.columns:
            mask_empty = df["IncorrectExplanation"].astype(str).str.strip().eq("")
            df.loc[mask_empty, "IncorrectExplanation"] = df.loc[mask_empty, "CorrectExplanation"]
        
        return df
    
    def _save_to_excel(self, df: pd.DataFrame, output_path: str) -> None:
        """
        Save DataFrame to Excel file with optional summary sheet.
        
        Args:
            df: DataFrame to save
            output_path: Path for the output Excel file
        """
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Write main questions sheet
                df.to_excel(writer, sheet_name='Questions', index=False)
                
                # Create summary sheet if enabled
                if self.config.create_summary:
                    self._create_summary_sheet(writer, df)
                
                self.logger.info(f"Excel file saved: {output_path}")
                
        except Exception as e:
            error_msg = f"Error saving Excel file {output_path}: {str(e)}"
            self.logger.error(error_msg)
            self.stats['errors'].append(error_msg)
            raise
    
    def _create_summary_sheet(self, writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
        """
        Create a summary sheet with processing statistics.
        
        Args:
            writer: ExcelWriter object
            df: DataFrame with processed questions
        """
        import datetime
        
        summary_data = [
            ["Processing Summary", ""],
            ["Total Questions", len(df)],
            ["Files Processed", self.stats['files_processed']],
            ["Questions per File", f"{len(df) / max(1, self.stats['files_processed']):.1f}"],
            ["Processing Date", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            ["", ""],
            ["Question Distribution", ""],
        ]
        
        # Add group distribution
        if 'Group' in df.columns:
            group_counts = df['Group'].value_counts()
            for group, count in group_counts.items():
                summary_data.append([f"  {group}", count])
        
        summary_data.append(["", ""])
        
        # Add type distribution  
        if 'Type' in df.columns:
            type_counts = df['Type'].value_counts()
            summary_data.append(["Question Types", ""])
            for qtype, count in type_counts.items():
                summary_data.append([f"  {qtype}", count])
        
        # Add errors if any
        if self.stats['errors']:
            summary_data.extend([
                ["", ""],
                ["Errors Encountered", len(self.stats['errors'])],
            ])
            for i, error in enumerate(self.stats['errors'][:10], 1):  # Show first 10 errors
                summary_data.append([f"  Error {i}", error])
        
        summary_df = pd.DataFrame(summary_data, columns=["Metric", "Value"])
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    def get_stats(self) -> Dict[str, Union[int, List[str]]]:
        """
        Get processing statistics.
        
        Returns:
            Dictionary with processing statistics
        """
        return self.stats.copy()
    
    def reset_stats(self) -> None:
        """Reset processing statistics."""
        self.stats = {
            'files_processed': 0,
            'questions_extracted': 0, 
            'questions_valid': 0,
            'errors': []
        }