"""
Test suite for DOC to Excel Converter

This module contains unit tests for the converter functionality.

Author: Your Name
"""

import unittest
import os
import tempfile
import sys
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from config import Config
from utils import normalize_text, validate_paths, is_empty_or_whitespace
from validators import QuestionValidator, ValidationError


class TestUtils(unittest.TestCase):
    """Test utility functions."""
    
    def test_normalize_text(self):
        """Test text normalization."""
        # Test basic whitespace normalization
        self.assertEqual(normalize_text("  hello   world  "), "hello world")
        
        # Test empty/None input
        self.assertEqual(normalize_text(""), "")
        self.assertEqual(normalize_text(None), "")
        
        # Test Unicode character replacement
        smart_quotes = ""Hello world""
        expected = '"Hello world"'
        self.assertEqual(normalize_text(smart_quotes), expected)
        
        # Test multiple spaces
        self.assertEqual(normalize_text("a    b\n\n  c"), "a b c")
    
    def test_validate_paths(self):
        """Test path validation."""
        # Test with non-existent paths
        input_valid, output_valid, error = validate_paths("nonexistent.docx", "output.xlsx")
        self.assertFalse(input_valid)
        self.assertTrue("does not exist" in error.lower())
        
        # Test with empty input
        input_valid, output_valid, error = validate_paths(None, "output.xlsx")
        self.assertFalse(input_valid)
        self.assertTrue("required" in error.lower())
    
    def test_is_empty_or_whitespace(self):
        """Test empty/whitespace checking."""
        self.assertTrue(is_empty_or_whitespace(""))
        self.assertTrue(is_empty_or_whitespace("   "))
        self.assertTrue(is_empty_or_whitespace("\n\t  "))
        self.assertFalse(is_empty_or_whitespace("hello"))
        self.assertFalse(is_empty_or_whitespace("  hello  "))


class TestConfig(unittest.TestCase):
    """Test configuration functionality."""
    
    def test_default_config(self):
        """Test default configuration creation."""
        config = Config()
        
        self.assertEqual(config.group, "General")
        self.assertEqual(config.question_type, "MCQ")
        self.assertTrue(config.validate_answers)
        self.assertTrue(config.include_topics)
        self.assertFalse(config.verbose)
    
    def test_config_from_dict(self):
        """Test creating config from dictionary."""
        config_dict = {
            "group": "Test Group",
            "question_type": "Test Type",
            "verbose": True
        }
        
        config = Config.from_dict(config_dict)
        
        self.assertEqual(config.group, "Test Group")
        self.assertEqual(config.question_type, "Test Type")
        self.assertTrue(config.verbose)
        # Default values should be preserved
        self.assertTrue(config.validate_answers)
    
    def test_config_validation(self):
        """Test configuration validation."""
        # Valid config
        config = Config()
        errors = config.validate()
        self.assertEqual(len(errors), 0)
        
        # Invalid config
        invalid_config = Config(group="", max_questions_per_file=-1)
        errors = invalid_config.validate()
        self.assertGreater(len(errors), 0)
        self.assertTrue(any("Group cannot be empty" in error for error in errors))


class TestValidators(unittest.TestCase):
    """Test question validation functionality."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.config = Config()
        self.validator = QuestionValidator(self.config)
        
        # Sample valid question
        self.valid_question = {
            "MCQ_NO": "1",
            "Question": "What is the capital of France?",
            "Answer1": "London",
            "Answer2": "Berlin", 
            "Answer3": "Paris",
            "Answer4": "Madrid",
            "Answer5": "",
            "CorAns": "3",
            "CorrectExplanation": "Paris is the capital and largest city of France.",
            "IncorrectExplanation": "Other cities are capitals of different countries.",
            "Topic": "Geography",
            "Group": "General Knowledge",
            "Type": "MCQ",
            "SourceFile": "test.docx"
        }
    
    def test_valid_question(self):
        """Test validation of a valid question."""
        errors = self.validator.validate_single_question(self.valid_question)
        self.assertEqual(len(errors), 0)
    
    def test_missing_required_fields(self):
        """Test validation with missing required fields."""
        # Remove required field
        invalid_question = self.valid_question.copy()
        del invalid_question["Question"]
        
        errors = self.validator.validate_single_question(invalid_question)
        self.assertGreater(len(errors), 0)
        self.assertTrue(any(error.error_type == "missing_field" for error in errors))
    
    def test_empty_required_fields(self):
        """Test validation with empty required fields."""
        invalid_question = self.valid_question.copy()
        invalid_question["Question"] = ""
        
        errors = self.validator.validate_single_question(invalid_question)
        self.assertGreater(len(errors), 0)
        self.assertTrue(any(error.error_type == "empty_field" for error in errors))
    
    def test_invalid_correct_answer(self):
        """Test validation with invalid correct answer."""
        invalid_question = self.valid_question.copy()
        invalid_question["CorAns"] = "6"  # Invalid answer number
        
        errors = self.validator.validate_single_question(invalid_question)
        self.assertGreater(len(errors), 0)
        self.assertTrue(any(error.error_type == "invalid_answer_format" for error in errors))
    
    def test_insufficient_options(self):
        """Test validation with insufficient options."""
        invalid_question = self.valid_question.copy()
        invalid_question["Answer1"] = "Only option"
        invalid_question["Answer2"] = ""
        invalid_question["Answer3"] = ""
        invalid_question["Answer4"] = ""
        
        errors = self.validator.validate_single_question(invalid_question)
        self.assertGreater(len(errors), 0)
        self.assertTrue(any(error.error_type == "insufficient_options" for error in errors))
    
    def test_batch_validation(self):
        """Test batch validation functionality."""
        questions = [self.valid_question.copy() for _ in range(3)]
        
        # Add duplicate question number
        questions[1]["MCQ_NO"] = "1"  # Same as questions[0]
        
        errors = self.validator.validate_batch_consistency(questions)
        self.assertGreater(len(errors), 0)
        self.assertTrue(any("Duplicate question number" in error for error in errors))
    
    def test_validation_summary(self):
        """Test validation summary generation."""
        questions = [self.valid_question.copy() for _ in range(2)]
        
        # Make one question invalid
        questions[1]["Question"] = ""
        
        summary = self.validator.get_validation_summary(questions)
        
        self.assertEqual(summary["total_questions"], 2)
        self.assertEqual(summary["valid_questions"], 1)
        self.assertEqual(summary["questions_with_errors"], 1)


class TestMockDocumentParsing(unittest.TestCase):
    """Test document parsing with mock data."""
    
    def setUp(self):
        """Set up mock question data."""
        self.mock_questions = [
            {
                "MCQ_NO": "1",
                "Question": "Which organ pumps blood throughout the body?",
                "Answer1": "Liver",
                "Answer2": "Heart",
                "Answer3": "Kidney",
                "Answer4": "Lung",
                "CorAns": "2",
                "CorrectExplanation": "The heart is a muscular organ that pumps blood.",
                "IncorrectExplanation": "Other organs have different functions.",
                "Topic": "Anatomy",
                "Group": "Medical",
                "Type": "MCQ",
                "SourceFile": "anatomy.docx"
            },
            {
                "MCQ_NO": "2", 
                "Question": "What is the normal body temperature in Celsius?",
                "Answer1": "35°C",
                "Answer2": "36°C",
                "Answer3": "37°C", 
                "Answer4": "38°C",
                "CorAns": "3",
                "CorrectExplanation": "Normal body temperature is around 37°C.",
                "IncorrectExplanation": "Other temperatures are abnormal.",
                "Topic": "Physiology",
                "Group": "Medical",
                "Type": "MCQ", 
                "SourceFile": "physiology.docx"
            }
        ]
    
    def test_mock_data_structure(self):
        """Test that mock data has correct structure."""
        required_fields = [
            "MCQ_NO", "Question", "Answer1", "Answer2", "Answer3", "Answer4",
            "CorAns", "CorrectExplanation", "Topic", "Group", "Type", "SourceFile"
        ]
        
        for question in self.mock_questions:
            for field in required_fields:
                self.assertIn(field, question)
                self.assertIsNotNone(question[field])
    
    def test_mock_data_validation(self):
        """Test validation of mock data."""
        config = Config()
        validator = QuestionValidator(config)
        
        for question in self.mock_questions:
            errors = validator.validate_single_question(question)
            self.assertEqual(len(errors), 0, f"Question {question['MCQ_NO']} should be valid")


class TestIntegration(unittest.TestCase):
    """Integration tests for the complete workflow."""
    
    def setUp(self):
        """Set up temporary directories for testing."""
        self.temp_dir = tempfile.mkdtemp()
        self.output_file = os.path.join(self.temp_dir, "test_output.xlsx")
    
    def tearDown(self):
        """Clean up temporary files."""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_output_file_creation(self):
        """Test that output files can be created."""
        # This test would require actual docx files to process
        # For now, we just test that the output directory can be created
        self.assertTrue(os.path.exists(self.temp_dir))
        
        # Test that we can write to the output file path
        test_content = "test"
        with open(self.output_file, 'w') as f:
            f.write(test_content)
        
        self.assertTrue(os.path.exists(self.output_file))


def run_tests():
    """Run all tests and return results."""
    # Create test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add test classes
    test_classes = [
        TestUtils,
        TestConfig,
        TestValidators,
        TestMockDocumentParsing,
        TestIntegration
    ]
    
    for test_class in test_classes:
        tests = loader.loadTestsFromTestCase(test_class)
        suite.addTests(tests)
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    return result


if __name__ == "__main__":
    print("DOC to Excel Converter - Test Suite")
    print("=" * 50)
    
    result = run_tests()
    
    print("\n" + "=" * 50)
    print("Test Results Summary:")
    print(f"Tests run: {result.testsRun}")
    print(f"Failures: {len(result.failures)}")
    print(f"Errors: {len(result.errors)}")
    
    if result.failures:
        print("\nFailures:")
        for test, traceback in result.failures:
            print(f"- {test}")
    
    if result.errors:
        print("\nErrors:")
        for test, traceback in result.errors:
            print(f"- {test}")
    
    success = len(result.failures) == 0 and len(result.errors) == 0
    print(f"\nOverall: {'PASSED' if success else 'FAILED'}")
    
    sys.exit(0 if success else 1)