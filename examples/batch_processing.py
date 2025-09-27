#!/usr/bin/env python3
"""
Batch processing example for DOC to Excel Converter

This example shows how to process multiple directories
and handle errors gracefully in batch operations.

Author: Your Name
"""

import sys
import os
from pathlib import Path
import time

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from doc_converter import DocToExcelConverter
from config import Config, MEDICAL_CONFIG
from utils import ProgressTracker


class BatchProcessor:
    """Handles batch processing of multiple document collections."""
    
    def __init__(self, config=None):
        self.config = config or MEDICAL_CONFIG
        self.converter = DocToExcelConverter(self.config)
        self.results = []
    
    def process_multiple_directories(self, directory_paths, output_dir="batch_output"):
        """
        Process multiple directories and create separate Excel files.
        
        Args:
            directory_paths: List of directory paths to process
            output_dir: Directory to save output files
        """
        print(f"🚀 Starting batch processing of {len(directory_paths)} directories...")
        
        # Create output directory
        os.makedirs(output_dir, exist_ok=True)
        
        progress = ProgressTracker(len(directory_paths), "Processing directories")
        
        for i, dir_path in enumerate(directory_paths):
            dir_name = Path(dir_path).name
            output_file = os.path.join(output_dir, f"{dir_name}_questions.xlsx")
            
            try:
                print(f"\n📁 Processing directory: {dir_path}")
                start_time = time.time()
                
                result_df = self.converter.convert_directory(dir_path, output_file)
                
                processing_time = time.time() - start_time
                stats = self.converter.get_stats()
                
                self.results.append({
                    'directory': dir_path,
                    'output_file': output_file,
                    'questions': len(result_df),
                    'files_processed': stats['files_processed'],
                    'processing_time': processing_time,
                    'errors': len(stats['errors']),
                    'success': True
                })
                
                print(f"✅ Completed: {len(result_df)} questions in {processing_time:.1f}s")
                
            except Exception as e:
                print(f"❌ Error processing {dir_path}: {e}")
                self.results.append({
                    'directory': dir_path,
                    'output_file': output_file,
                    'questions': 0,
                    'files_processed': 0,
                    'processing_time': 0,
                    'errors': 1,
                    'success': False,
                    'error_message': str(e)
                })
            
            progress.update()
            self.converter.reset_stats()  # Reset for next directory
        
        progress.finish()
        self.print_batch_summary()
    
    def process_and_merge(self, directory_paths, output_file="merged_questions.xlsx"):
        """
        Process multiple directories and merge all questions into one file.
        
        Args:
            directory_paths: List of directory paths to process
            output_file: Output Excel file for merged results
        """
        print(f"🔄 Processing and merging {len(directory_paths)} directories...")
        
        all_questions = []
        progress = ProgressTracker(len(directory_paths), "Processing for merge")
        
        for dir_path in directory_paths:
            try:
                print(f"\n📁 Processing: {dir_path}")
                
                # Get list of docx files
                docx_files = list(Path(dir_path).glob("*.docx"))
                docx_files = [f for f in docx_files if not f.name.startswith('~')]
                
                for file_path in docx_files:
                    questions = self.converter.parse_docx(str(file_path))
                    
                    # Add directory info to questions
                    for question in questions:
                        question['SourceDirectory'] = Path(dir_path).name
                        question['Group'] = f"{self.config.group} - {Path(dir_path).name}"
                    
                    all_questions.extend(questions)
                
                print(f"✅ Extracted questions from {len(docx_files)} files")
                
            except Exception as e:
                print(f"❌ Error processing {dir_path}: {e}")
            
            progress.update()
        
        progress.finish()
        
        if all_questions:
            # Create merged DataFrame and save
            import pandas as pd
            
            # Add SourceDirectory to column mapping if not present
            columns = [
                "MCQ_NO", "Question", "Answer1", "CorAns", "Answer2", "Answer3",
                "Answer4", "Answer5", "CorrectExplanation", "IncorrectExplanation",
                "Topic", "Group", "Type", "SourceFile", "SourceDirectory"
            ]
            
            df = pd.DataFrame(all_questions)
            
            # Ensure all columns exist
            for col in columns:
                if col not in df.columns:
                    df[col] = ""
            
            df = df.reindex(columns=columns)
            
            # Save with summary
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Questions', index=False)
                
                # Create detailed summary
                summary_data = [
                    ["Merge Summary", ""],
                    ["Total Questions", len(df)],
                    ["Total Directories", len(directory_paths)],
                    ["", ""],
                    ["Questions by Directory", ""]
                ]
                
                if 'SourceDirectory' in df.columns:
                    dir_counts = df['SourceDirectory'].value_counts()
                    for directory, count in dir_counts.items():
                        summary_data.append([f"  {directory}", count])
                
                summary_df = pd.DataFrame(summary_data, columns=["Metric", "Value"])
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            print(f"\n✅ Merged {len(all_questions)} questions into {output_file}")
        else:
            print("\n❌ No questions found to merge")
    
    def print_batch_summary(self):
        """Print summary of batch processing results."""
        print("\n" + "="*60)
        print("BATCH PROCESSING SUMMARY")
        print("="*60)
        
        total_questions = sum(r['questions'] for r in self.results)
        total_files = sum(r['files_processed'] for r in self.results)
        successful_dirs = sum(1 for r in self.results if r['success'])
        failed_dirs = len(self.results) - successful_dirs
        
        print(f"📊 Total Questions Processed: {total_questions}")
        print(f"📁 Total Files Processed: {total_files}")
        print(f"✅ Successful Directories: {successful_dirs}")
        print(f"❌ Failed Directories: {failed_dirs}")
        
        if self.results:
            avg_questions = total_questions / len(self.results)
            print(f"📈 Average Questions per Directory: {avg_questions:.1f}")
        
        # Show individual results
        print("\nDetailed Results:")
        print("-" * 60)
        for result in self.results:
            status = "✅" if result['success'] else "❌"
            print(f"{status} {Path(result['directory']).name}: {result['questions']} questions")
            if not result['success'] and 'error_message' in result:
                print(f"    Error: {result['error_message']}")


def example_batch_processing():
    """Example of processing multiple directories."""
    
    # Example directory paths - update these to match your data
    directories = [
        "../sample_data/cardiology/",
        "../sample_data/pulmonology/", 
        "../sample_data/gastroenterology/",
        # Add more directories as needed
    ]
    
    # Filter to only existing directories
    existing_dirs = [d for d in directories if os.path.exists(d)]
    
    if not existing_dirs:
        print("❌ No input directories found. Please update the directory paths.")
        print("Example directories to create:")
        for d in directories:
            print(f"  - {d}")
        return
    
    print(f"Found {len(existing_dirs)} directories to process:")
    for d in existing_dirs:
        print(f"  - {d}")
    
    # Create batch processor
    processor = BatchProcessor(MEDICAL_CONFIG)
    
    # Process each directory separately
    print("\n" + "="*50)
    print("OPTION 1: Process directories separately")
    print("="*50)
    processor.process_multiple_directories(existing_dirs, "batch_output_separate")
    
    # Process and merge all directories
    print("\n" + "="*50)
    print("OPTION 2: Process and merge all directories")
    print("="*50)
    processor.process_and_merge(existing_dirs, "batch_output_merged.xlsx")


def example_custom_batch_config():
    """Example of batch processing with custom configuration."""
    
    # Custom config for specific requirements
    custom_config = Config(
        group="Custom Batch Processing",
        question_type="Batch MCQ",
        validate_answers=True,
        include_topics=True,
        verbose=True
    )
    
    processor = BatchProcessor(custom_config)
    
    # Add your processing logic here
    print("Custom batch processor created with:")
    print(f"  Group: {custom_config.group}")
    print(f"  Type: {custom_config.question_type}")
    print(f"  Validation: {'Enabled' if custom_config.validate_answers else 'Disabled'}")


def main():
    """Run batch processing examples."""
    print("DOC to Excel Converter - Batch Processing Examples")
    print("=" * 60)
    
    # Run examples
    example_batch_processing()
    example_custom_batch_config()
    
    print("\n" + "=" * 60)
    print("Batch processing examples completed!")
    print("\nTips for batch processing:")
    print("- Organize your .docx files in separate directories by topic/subject")
    print("- Use descriptive directory names (they become part of the output)")
    print("- Check the batch_output/ directory for individual Excel files")
    print("- Use the merged output for combined analysis across all topics")


if __name__ == "__main__":
    main()