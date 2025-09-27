"""
DOC to Excel Converter Package

A powerful Python package for converting Word documents containing
Multiple Choice Questions (MCQs) to structured Excel files.

Author: Your Name
Version: 1.0.0
"""

__version__ = "1.0.0"
__author__ = "Your Name"
__email__ = "your.email@example.com"
__description__ = "Convert Word documents containing MCQs to Excel format"

# Import main classes for easy access
from .doc_converter import DocToExcelConverter
from .config import Config, MEDICAL_CONFIG, EDUCATIONAL_CONFIG, SIMPLE_CONFIG
from .parsers import QuestionParser, EnhancedMedicalParser
from .validators import QuestionValidator, ValidationError
from .utils import (
    normalize_text,
    validate_paths,
    setup_logging,
    ProgressTracker
)

__all__ = [
    # Main converter
    'DocToExcelConverter',
    
    # Configuration
    'Config',
    'MEDICAL_CONFIG',
    'EDUCATIONAL_CONFIG', 
    'SIMPLE_CONFIG',
    
    # Parsers
    'QuestionParser',
    'EnhancedMedicalParser',
    
    # Validators
    'QuestionValidator',
    'ValidationError',
    
    # Utilities
    'normalize_text',
    'validate_paths',
    'setup_logging',
    'ProgressTracker'
]