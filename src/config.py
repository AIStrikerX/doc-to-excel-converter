"""
Configuration module for DOC to Excel Converter

This module handles all configuration settings and provides
a centralized way to manage converter behavior.

Author: Your Name
Version: 1.0.0
"""

from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional
import json
import os


@dataclass
class Config:
    """
    Configuration class for the DOC to Excel converter.
    
    Attributes:
        group: Default group/category for questions
        question_type: Default type for questions (MCQ, True/False, etc.)
        validate_answers: Whether to validate parsed answers
        include_topics: Whether to extract and include question topics
        create_summary: Whether to create summary sheet in Excel output
        verbose: Enable verbose logging
        quiet: Suppress all output except errors
        max_questions_per_file: Maximum questions to process per file (0 = unlimited)
        encoding: Text encoding for processing
    """
    
    # Basic settings
    group: str = "General"
    question_type: str = "MCQ"
    
    # Processing options
    validate_answers: bool = True
    include_topics: bool = True
    create_summary: bool = True
    
    # Output settings
    verbose: bool = False
    quiet: bool = False
    
    # Advanced settings
    max_questions_per_file: int = 0  # 0 = unlimited
    encoding: str = "utf-8"
    
    # Parsing patterns
    question_patterns: List[str] = field(default_factory=lambda: [
        r"^\s*Question\s+(\d+)(?:\s*[-–—]\s*(.+))?\s*$",
        r"^\s*Q[\.\s]*(\d+)(?:\s*[-–—]\s*(.+))?\s*$"
    ])
    
    option_patterns: List[str] = field(default_factory=lambda: [
        r"^([A-E])[\.\)]\s*(.*)",
        r"^([1-5])[\.\)]\s*(.*)"
    ])
    
    answer_patterns: List[str] = field(default_factory=lambda: [
        r"^\s*Correct\s*Answer\s*[:\-]?\s*([A-Ea-e1-5])\s*$",
        r"^\s*Answer\s*[:\-]?\s*([A-Ea-e1-5])\s*$"
    ])
    
    explanation_patterns: List[str] = field(default_factory=lambda: [
        r"^\s*Explanation(\s+(of\s+the\s+)?Correct\s+Answer)?\s*[:\-]?\s*$",
        r"^\s*Rationale\s*[:\-]?\s*$"
    ])
    
    # Column mapping
    column_mapping: Dict[str, str] = field(default_factory=lambda: {
        "question_number": "MCQ_NO",
        "question_text": "Question",
        "option_a": "Answer1", 
        "option_b": "Answer2",
        "option_c": "Answer3",
        "option_d": "Answer4",
        "option_e": "Answer5",
        "correct_answer": "CorAns",
        "correct_explanation": "CorrectExplanation",
        "incorrect_explanation": "IncorrectExplanation",
        "topic": "Topic",
        "group": "Group",
        "type": "Type",
        "source_file": "SourceFile"
    })
    
    @classmethod
    def from_dict(cls, config_dict: Dict[str, Any]) -> 'Config':
        """
        Create Config instance from dictionary.
        
        Args:
            config_dict: Dictionary containing configuration values
            
        Returns:
            Config instance
        """
        # Extract known fields
        known_fields = {
            field.name for field in cls.__dataclass_fields__.values()
        }
        
        filtered_dict = {
            key: value for key, value in config_dict.items() 
            if key in known_fields
        }
        
        return cls(**filtered_dict)
    
    @classmethod
    def from_file(cls, config_path: str) -> 'Config':
        """
        Load configuration from JSON file.
        
        Args:
            config_path: Path to the configuration file
            
        Returns:
            Config instance
            
        Raises:
            FileNotFoundError: If config file doesn't exist
            json.JSONDecodeError: If config file is not valid JSON
        """
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"Configuration file not found: {config_path}")
        
        with open(config_path, 'r', encoding='utf-8') as f:
            config_dict = json.load(f)
        
        return cls.from_dict(config_dict)
    
    def to_dict(self) -> Dict[str, Any]:
        """
        Convert Config instance to dictionary.
        
        Returns:
            Dictionary representation of configuration
        """
        import dataclasses
        return dataclasses.asdict(self)
    
    def save_to_file(self, config_path: str) -> None:
        """
        Save configuration to JSON file.
        
        Args:
            config_path: Path where to save the configuration
        """
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
    
    def update(self, **kwargs) -> 'Config':
        """
        Create new Config instance with updated values.
        
        Args:
            **kwargs: Configuration values to update
            
        Returns:
            New Config instance with updated values
        """
        import dataclasses
        return dataclasses.replace(self, **kwargs)
    
    def validate(self) -> List[str]:
        """
        Validate configuration settings.
        
        Returns:
            List of validation errors (empty if valid)
        """
        errors = []
        
        if not self.group:
            errors.append("Group cannot be empty")
        
        if not self.question_type:
            errors.append("Question type cannot be empty")
        
        if self.max_questions_per_file < 0:
            errors.append("Max questions per file must be non-negative")
        
        if not self.encoding:
            errors.append("Encoding cannot be empty")
        
        # Validate patterns
        import re
        for i, pattern in enumerate(self.question_patterns):
            try:
                re.compile(pattern)
            except re.error as e:
                errors.append(f"Invalid question pattern {i}: {e}")
        
        for i, pattern in enumerate(self.option_patterns):
            try:
                re.compile(pattern)
            except re.error as e:
                errors.append(f"Invalid option pattern {i}: {e}")
        
        return errors


# Predefined configurations
MEDICAL_CONFIG = Config(
    group="Medical Questions",
    question_type="Medical Board Review",
    validate_answers=True,
    include_topics=True,
    create_summary=True
)

EDUCATIONAL_CONFIG = Config(
    group="Educational Content",
    question_type="Academic Quiz",
    validate_answers=True,
    include_topics=True,
    create_summary=True
)

SIMPLE_CONFIG = Config(
    group="General",
    question_type="MCQ",
    validate_answers=False,
    include_topics=False,
    create_summary=False,
    verbose=False
)


def get_default_config() -> Config:
    """Get the default configuration."""
    return Config()


def load_config_from_env() -> Config:
    """
    Load configuration from environment variables.
    
    Returns:
        Config instance with values from environment variables
    """
    return Config(
        group=os.getenv('DOC_CONVERTER_GROUP', 'General'),
        question_type=os.getenv('DOC_CONVERTER_TYPE', 'MCQ'),
        validate_answers=os.getenv('DOC_CONVERTER_VALIDATE', 'true').lower() == 'true',
        include_topics=os.getenv('DOC_CONVERTER_TOPICS', 'true').lower() == 'true',
        create_summary=os.getenv('DOC_CONVERTER_SUMMARY', 'true').lower() == 'true',
        verbose=os.getenv('DOC_CONVERTER_VERBOSE', 'false').lower() == 'true',
        quiet=os.getenv('DOC_CONVERTER_QUIET', 'false').lower() == 'true'
    )