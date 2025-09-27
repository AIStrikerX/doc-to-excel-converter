"""
Utility functions for DOC to Excel Converter

This module contains helper functions used throughout the converter.

Author: Your Name
Version: 1.0.0
"""

import os
import re
import logging
from typing import Tuple, Optional, Any
from pathlib import Path


def normalize_text(text: str) -> str:
    """
    Normalize and clean text by removing extra whitespace and special characters.
    
    Args:
        text: Input text to normalize
        
    Returns:
        Cleaned and normalized text
    """
    if not text:
        return ""
    
    # Replace multiple whitespaces with single space
    text = re.sub(r'\s+', ' ', text)
    
    # Replace special Unicode characters
    text = re.sub(r'\u00a0', ' ', text)              # Non-breaking space
    text = re.sub(r'[\u2018\u2019]', "'", text)      # Smart single quotes
    text = re.sub(r'[\u201c\u201d]', '"', text)      # Smart double quotes
    text = re.sub(r'[\u2013\u2014]', '-', text)      # En/em dashes to hyphen
    text = re.sub(r'[\u2026]', '...', text)          # Ellipsis
    
    return text.strip()


def validate_paths(input_path: Optional[str], output_path: Optional[str]) -> Tuple[bool, bool, str]:
    """
    Validate input and output paths.
    
    Args:
        input_path: Path to input file or directory
        output_path: Path to output file
        
    Returns:
        Tuple of (input_valid, output_valid, error_message)
    """
    error_msg = ""
    
    # Validate input path
    input_valid = True
    if not input_path:
        input_valid = False
        error_msg = "Input path is required"
    elif not os.path.exists(input_path):
        input_valid = False
        error_msg = f"Input path does not exist: {input_path}"
    elif os.path.isfile(input_path) and not input_path.lower().endswith('.docx'):
        input_valid = False
        error_msg = f"Input file must be a .docx file: {input_path}"
    
    # Validate output path
    output_valid = True
    if output_path:
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
            except OSError:
                output_valid = False
                error_msg = f"Cannot create output directory: {output_dir}"
        
        if not output_path.lower().endswith('.xlsx'):
            output_path += '.xlsx'
    
    return input_valid, output_valid, error_msg


def setup_logging(verbose: bool = False, quiet: bool = False) -> logging.Logger:
    """
    Set up logging configuration.
    
    Args:
        verbose: Enable verbose logging
        quiet: Suppress all output except errors
        
    Returns:
        Configured logger instance
    """
    logger = logging.getLogger('doc_converter')
    
    # Clear existing handlers
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # Set log level
    if quiet:
        level = logging.ERROR
    elif verbose:
        level = logging.DEBUG
    else:
        level = logging.INFO
    
    logger.setLevel(level)
    
    # Create console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)
    
    # Create formatter
    if verbose:
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
    else:
        formatter = logging.Formatter('%(levelname)s: %(message)s')
    
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger


def get_file_size(file_path: str) -> str:
    """
    Get human-readable file size.
    
    Args:
        file_path: Path to the file
        
    Returns:
        File size as formatted string
    """
    try:
        size = os.path.getsize(file_path)
        
        # Convert to human-readable format
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        
        return f"{size:.1f} TB"
    
    except OSError:
        return "Unknown"


def count_docx_files(directory: str, pattern: str = "*.docx") -> int:
    """
    Count .docx files in a directory.
    
    Args:
        directory: Directory path to search
        pattern: File pattern to match
        
    Returns:
        Number of matching files
    """
    try:
        path = Path(directory)
        files = list(path.glob(pattern))
        # Filter out temporary files
        files = [f for f in files if not f.name.startswith('~')]
        return len(files)
    except Exception:
        return 0


def create_backup(file_path: str) -> Optional[str]:
    """
    Create a backup copy of a file.
    
    Args:
        file_path: Path to the file to backup
        
    Returns:
        Path to the backup file, or None if backup failed
    """
    try:
        if not os.path.exists(file_path):
            return None
        
        backup_path = f"{file_path}.backup"
        counter = 1
        
        # Find available backup filename
        while os.path.exists(backup_path):
            backup_path = f"{file_path}.backup.{counter}"
            counter += 1
        
        # Copy file
        import shutil
        shutil.copy2(file_path, backup_path)
        return backup_path
    
    except Exception:
        return None


def replace_option_letters_with_numbers(text: str) -> str:
    """
    Convert A/B/C/D/E references to numbered format in explanations.
    
    Args:
        text: Text containing option references
        
    Returns:
        Text with options converted to numbers
    """
    if not text:
        return text
    
    # Replace standalone A/B/C/D/E with number-
    replacements = {
        r'\bA[\)\.]': '1-',
        r'\bB[\)\.]': '2-',
        r'\bC[\)\.]': '3-',
        r'\bD[\)\.]': '4-',
        r'\bE[\)\.]': '5-'
    }
    
    for pattern, replacement in replacements.items():
        text = re.sub(pattern, replacement, text)
    
    return text


def sanitize_filename(filename: str) -> str:
    """
    Sanitize filename by removing/replacing invalid characters.
    
    Args:
        filename: Original filename
        
    Returns:
        Sanitized filename safe for filesystem use
    """
    # Remove invalid characters
    invalid_chars = r'<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    
    # Remove leading/trailing dots and spaces
    filename = filename.strip('. ')
    
    # Limit length
    if len(filename) > 200:
        name, ext = os.path.splitext(filename)
        filename = name[:200-len(ext)] + ext
    
    return filename


def format_duration(seconds: float) -> str:
    """
    Format duration in seconds to human-readable format.
    
    Args:
        seconds: Duration in seconds
        
    Returns:
        Formatted duration string
    """
    if seconds < 1:
        return f"{seconds*1000:.0f}ms"
    elif seconds < 60:
        return f"{seconds:.1f}s"
    elif seconds < 3600:
        minutes = int(seconds // 60)
        secs = int(seconds % 60)
        return f"{minutes}m {secs}s"
    else:
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        return f"{hours}h {minutes}m"


def merge_dicts(*dicts) -> dict:
    """
    Merge multiple dictionaries, with later ones taking precedence.
    
    Args:
        *dicts: Variable number of dictionaries to merge
        
    Returns:
        Merged dictionary
    """
    result = {}
    for d in dicts:
        if d:
            result.update(d)
    return result


def safe_get(dictionary: dict, key: str, default: Any = "") -> Any:
    """
    Safely get value from dictionary with default.
    
    Args:
        dictionary: Dictionary to get value from
        key: Key to look up
        default: Default value if key not found
        
    Returns:
        Value from dictionary or default
    """
    try:
        return dictionary.get(key, default)
    except (AttributeError, TypeError):
        return default


def is_empty_or_whitespace(text: str) -> bool:
    """
    Check if text is empty or contains only whitespace.
    
    Args:
        text: Text to check
        
    Returns:
        True if text is empty or whitespace only
    """
    return not text or not text.strip()


def truncate_text(text: str, max_length: int = 100, suffix: str = "...") -> str:
    """
    Truncate text to specified length with optional suffix.
    
    Args:
        text: Text to truncate
        max_length: Maximum length of result
        suffix: Suffix to add if text is truncated
        
    Returns:
        Truncated text
    """
    if not text or len(text) <= max_length:
        return text
    
    truncated = text[:max_length - len(suffix)]
    return truncated + suffix


class ProgressTracker:
    """Simple progress tracker for console output."""
    
    def __init__(self, total: int, description: str = "Processing"):
        self.total = total
        self.current = 0
        self.description = description
    
    def update(self, increment: int = 1) -> None:
        """Update progress by increment."""
        self.current += increment
        self._print_progress()
    
    def set_progress(self, current: int) -> None:
        """Set current progress value."""
        self.current = current
        self._print_progress()
    
    def _print_progress(self) -> None:
        """Print current progress to console."""
        if self.total > 0:
            percentage = (self.current / self.total) * 100
            print(f"\r{self.description}: {self.current}/{self.total} ({percentage:.1f}%)", end="", flush=True)
        else:
            print(f"\r{self.description}: {self.current}", end="", flush=True)
    
    def finish(self) -> None:
        """Mark progress as complete."""
        print()  # New line after progress