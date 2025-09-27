#!/usr/bin/env python3
"""
Basic usage example for DOC to Excel Converter

This example shows the simplest way to use the converter
to process Word documents containing MCQs.

Author: Your Name
"""

import sys
import os

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from doc_converter import DocToExcelConverter
from config import Config


def basic_single_file_conversion():
    """Convert a single .docx file to Excel."""
    print("=== Basic Single File Conversion ===")
    
    # Create converter with default settings
    converter = DocToExcelConverter()
    
    # Convert single file
    input_file = "../sample_data/sample1.docx"  # Update this path
    output_file = "output_basic.xlsx"
    
    try:
        if os.path.exists(input_file):
            result_df = converter.convert_file(input_file, output_file)
            print(f"✅ Successfully converted {len(result_df)} questions!")
            print(f"📊 Output saved to: {output_file}")
        else:
            print(f"❌ Input file not found: {input_file}")
            print("Please update the input_file path in this script")
    
    except Exception as e:
        print(f"❌ Error during conversion: {e}")


def basic_directory_conversion():
    """Convert all .docx files in a directory to Excel."""
    print("\n=== Basic Directory Conversion ===")
    
    # Create converter with default settings
    converter = DocToExcelConverter()
    
    # Convert all files in directory
    input_dir = "../sample_data/"  # Update this path
    output_file = "output_directory.xlsx"
    
    try:
        if os.path.exists(input_dir):
            result_df = converter.convert_directory(input_dir, output_file)
            print(f"✅ Successfully converted {len(result_df)} questions!")
            print(f"📊 Output saved to: {output_file}")
            
            # Show some statistics
            stats = converter.get_stats()
            print(f"📈 Files processed: {stats['files_processed']}")
            print(f"📋 Questions extracted: {stats['questions_extracted']}")
            
        else:
            print(f"❌ Input directory not found: {input_dir}")
            print("Please update the input_dir path in this script")
    
    except Exception as e:
        print(f"❌ Error during conversion: {e}")


def conversion_with_custom_settings():
    """Convert with custom settings."""
    print("\n=== Conversion with Custom Settings ===")
    
    # Create custom configuration
    config = Config(
        group="Sample Questions",
        question_type="Example MCQ",
        validate_answers=True,
        include_topics=True,
        verbose=True
    )
    
    # Create converter with custom config
    converter = DocToExcelConverter(config)
    
    # Convert with custom metadata
    input_dir = "../sample_data/"  # Update this path
    output_file = "output_custom.xlsx"
    
    try:
        if os.path.exists(input_dir):
            result_df = converter.convert_directory(
                input_dir, 
                output_file,
                group="Custom Group",  # Override config
                type="Custom Type"     # Override config
            )
            print(f"✅ Successfully converted {len(result_df)} questions!")
            print(f"📊 Output saved to: {output_file}")
        else:
            print(f"❌ Input directory not found: {input_dir}")
    
    except Exception as e:
        print(f"❌ Error during conversion: {e}")


def main():
    """Run all basic examples."""
    print("DOC to Excel Converter - Basic Usage Examples")
    print("=" * 50)
    
    # Run examples
    basic_single_file_conversion()
    basic_directory_conversion()
    conversion_with_custom_settings()
    
    print("\n" + "=" * 50)
    print("Examples completed! Check the output files for results.")
    print("\nTips:")
    print("- Update the file paths in this script to match your data")
    print("- Check the generated Excel files for the converted questions")
    print("- Look at the Summary sheet for processing statistics")


if __name__ == "__main__":
    main()