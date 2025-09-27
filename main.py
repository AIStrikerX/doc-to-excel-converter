#!/usr/bin/env python3
"""
DOC to Excel Converter - Command Line Interface

A powerful tool for bulk processing Word documents containing MCQs
and converting them to structured Excel files.

Usage:
    python main.py --input <path> --output <path> [options]
    python main.py --config <config_file>
    python main.py --help

Author: Your Name
Version: 1.0.0
"""

import argparse
import sys
import os
import json
from pathlib import Path
from typing import Optional, Dict, Any

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from doc_converter import DocToExcelConverter
from config import Config
from utils import setup_logging, validate_paths


def parse_arguments() -> argparse.Namespace:
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Convert Word documents containing MCQs to Excel format',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Convert single file
    python main.py --input "questions.docx" --output "result.xlsx"
    
    # Convert directory with custom settings
    python main.py --input "docs/" --output "questions.xlsx" \\
                  --group "Medicine" --type "Board Review"
    
    # Use configuration file
    python main.py --config "config.json"
    
    # Batch process with validation
    python main.py --input "medical_docs/" --output "output.xlsx" \\
                  --validate --verbose
        """
    )
    
    # Input/Output arguments
    parser.add_argument(
        '--input', '-i',
        type=str,
        help='Input file (.docx) or directory containing .docx files'
    )
    
    parser.add_argument(
        '--output', '-o',
        type=str,
        help='Output Excel file path (.xlsx)'
    )
    
    # Configuration
    parser.add_argument(
        '--config', '-c',
        type=str,
        help='Configuration file path (JSON format)'
    )
    
    # Question metadata
    parser.add_argument(
        '--group', '-g',
        type=str,
        default='General',
        help='Question group/category (default: General)'
    )
    
    parser.add_argument(
        '--type', '-t',
        type=str,
        default='MCQ',
        help='Question type (default: MCQ)'
    )
    
    # Processing options
    parser.add_argument(
        '--validate',
        action='store_true',
        help='Validate parsed questions for completeness'
    )
    
    parser.add_argument(
        '--include-topics',
        action='store_true',
        help='Include question topics in output'
    )
    
    parser.add_argument(
        '--create-summary',
        action='store_true',
        default=True,
        help='Create summary sheet in Excel output (default: True)'
    )
    
    # Output options
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose output'
    )
    
    parser.add_argument(
        '--quiet', '-q',
        action='store_true',
        help='Suppress all output except errors'
    )
    
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Parse files but don\'t create output (useful for validation)'
    )
    
    # Advanced options
    parser.add_argument(
        '--max-files',
        type=int,
        help='Maximum number of files to process (for testing)'
    )
    
    parser.add_argument(
        '--pattern',
        type=str,
        default='*.docx',
        help='File pattern to match (default: *.docx)'
    )
    
    parser.add_argument(
        '--encoding',
        type=str,
        default='utf-8',
        help='Text encoding for output (default: utf-8)'
    )
    
    # Version
    parser.add_argument(
        '--version',
        action='version',
        version='%(prog)s 1.0.0'
    )
    
    return parser.parse_args()


def load_config_file(config_path: str) -> Dict[str, Any]:
    """Load configuration from JSON file."""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        print(f"✅ Loaded configuration from {config_path}")
        return config
    except FileNotFoundError:
        print(f"❌ Configuration file not found: {config_path}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"❌ Invalid JSON in configuration file: {e}")
        sys.exit(1)


def merge_config_and_args(config_data: Dict[str, Any], args: argparse.Namespace) -> Config:
    """Merge configuration file data with command line arguments."""
    # Command line arguments override config file
    merged_config = Config(
        group=args.group if args.group != 'General' else config_data.get('group', 'General'),
        question_type=args.type if args.type != 'MCQ' else config_data.get('type', 'MCQ'),
        validate_answers=args.validate if args.validate else config_data.get('settings', {}).get('validate_answers', False),
        include_topics=args.include_topics if args.include_topics else config_data.get('settings', {}).get('include_topics', True),
        create_summary=args.create_summary,
        verbose=args.verbose,
        quiet=args.quiet
    )
    
    return merged_config


def main():
    """Main entry point for the CLI application."""
    args = parse_arguments()
    
    # Setup logging based on verbosity
    logger = setup_logging(verbose=args.verbose, quiet=args.quiet)
    
    try:
        # Load configuration if specified
        config_data = {}
        if args.config:
            config_data = load_config_file(args.config)
            
            # Use config file paths if not overridden by command line
            if not args.input and 'input_path' in config_data:
                args.input = config_data['input_path']
            if not args.output and 'output_path' in config_data:
                args.output = config_data['output_path']
        
        # Validate required arguments
        if not args.input:
            print("❌ Error: Input path is required. Use --input or specify in config file.")
            sys.exit(1)
            
        if not args.output and not args.dry_run:
            print("❌ Error: Output path is required. Use --output or specify in config file.")
            sys.exit(1)
        
        # Validate paths
        input_valid, output_valid, error_msg = validate_paths(args.input, args.output)
        if not input_valid:
            print(f"❌ Input path error: {error_msg}")
            sys.exit(1)
        
        if not args.dry_run and not output_valid:
            print(f"❌ Output path error: {error_msg}")
            sys.exit(1)
        
        # Create configuration
        config = merge_config_and_args(config_data, args) if config_data else Config(
            group=args.group,
            question_type=args.type,
            validate_answers=args.validate,
            include_topics=args.include_topics,
            create_summary=args.create_summary,
            verbose=args.verbose,
            quiet=args.quiet
        )
        
        if not args.quiet:
            print("🚀 Starting DOC to Excel conversion...")
            print(f"📂 Input: {args.input}")
            if not args.dry_run:
                print(f"📊 Output: {args.output}")
            print(f"🏷️  Group: {config.group}")
            print(f"📝 Type: {config.question_type}")
        
        # Initialize converter
        converter = DocToExcelConverter(config)
        
        # Process files
        if os.path.isfile(args.input):
            # Single file processing
            if not args.quiet:
                print(f"\n📄 Processing single file: {args.input}")
            
            if args.dry_run:
                questions = converter.parse_docx(args.input)
                print(f"✅ Dry run complete. Would process {len(questions)} questions.")
            else:
                result_df = converter.convert_file(args.input, args.output)
                if not args.quiet:
                    print(f"✅ Successfully processed {len(result_df)} questions!")
        
        elif os.path.isdir(args.input):
            # Directory processing
            if not args.quiet:
                print(f"\n📁 Processing directory: {args.input}")
            
            # Get file count for progress
            docx_files = list(Path(args.input).glob(args.pattern))
            if args.max_files:
                docx_files = docx_files[:args.max_files]
            
            if not docx_files:
                print(f"❌ No .docx files found in {args.input}")
                sys.exit(1)
            
            if not args.quiet:
                print(f"📋 Found {len(docx_files)} files to process")
            
            if args.dry_run:
                total_questions = 0
                for file_path in docx_files:
                    questions = converter.parse_docx(str(file_path))
                    total_questions += len(questions)
                    if args.verbose:
                        print(f"  📄 {file_path.name}: {len(questions)} questions")
                
                print(f"✅ Dry run complete. Would process {total_questions} questions from {len(docx_files)} files.")
            else:
                result_df = converter.convert_directory(args.input, args.output)
                if not args.quiet:
                    print(f"✅ Successfully processed {len(result_df)} questions from {len(docx_files)} files!")
        
        else:
            print(f"❌ Invalid input path: {args.input}")
            sys.exit(1)
        
        if not args.quiet and not args.dry_run:
            print(f"\n🎉 Conversion completed successfully!")
            print(f"📊 Output saved to: {args.output}")
    
    except KeyboardInterrupt:
        print("\n⚠️ Operation cancelled by user")
        sys.exit(1)
    
    except Exception as e:
        print(f"\n❌ Error during processing: {str(e)}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()