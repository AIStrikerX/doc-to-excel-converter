# DOC to Excel Converter 📄➡️📊

A powerful Python tool for bulk processing Word documents (.docx) containing Multiple Choice Questions (MCQs) and converting them to structured Excel files. Perfect for educational institutions, exam preparation services, and content creators who need to digitize large volumes of question banks.

## 🌟 Features

- **Bulk Processing**: Process single files or entire directories of .docx files
- **Smart Parsing**: Automatically detects question blocks, options (A-E), correct answers, and explanations
- **Highlighting Detection**: Separates highlighted text (correct explanations) from normal text (incorrect explanations)
- **Flexible Output**: Generates Excel files with customizable sheets and metadata
- **Error Handling**: Robust parsing with detailed error reporting and validation
- **Template Support**: Includes templates for consistent question formatting
- **CLI Interface**: Easy-to-use command-line interface with configuration options

## 📋 Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage Examples](#usage-examples)
- [Document Format](#document-format)
- [Output Format](#output-format)
- [Configuration](#configuration)
- [API Reference](#api-reference)
- [Contributing](#contributing)
- [License](#license)

## 🚀 Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package installer)

### Install from Source
```bash
git clone https://github.com/yourusername/doc-to-excel-converter.git
cd doc-to-excel-converter
pip install -r requirements.txt
```

### Install Required Packages
```bash
pip install pandas python-docx openpyxl argparse
```

## ⚡ Quick Start

### Basic Usage
```python
from src.doc_converter import DocToExcelConverter

# Initialize converter
converter = DocToExcelConverter()

# Convert single file
converter.convert_file("input.docx", "output.xlsx")

# Convert entire directory
converter.convert_directory("input_folder/", "output.xlsx")
```

### Command Line Usage
```bash
# Convert single file
python main.py --input "document.docx" --output "questions.xlsx"

# Convert directory with custom settings
python main.py --input "docs/" --output "medical_questions.xlsx" --group "Medicine" --type "Board Review"

# Batch processing with configuration
python main.py --config config.json
```

## 📖 Usage Examples

### Example 1: Medical Question Bank
```python
from src.doc_converter import DocToExcelConverter

converter = DocToExcelConverter()
result = converter.convert_directory(
    input_path="medical_docs/",
    output_path="medical_questions.xlsx",
    group="Internal Medicine",
    question_type="USMLE"
)
print(f"Processed {len(result)} questions")
```

### Example 2: Custom Configuration
```python
from src.doc_converter import DocToExcelConverter
from src.config import Config

config = Config(
    group="Cardiology",
    question_type="Board Review",
    include_topics=True,
    validate_answers=True
)

converter = DocToExcelConverter(config)
converter.convert_directory("cardiology_questions/", "output.xlsx")
```

### Example 3: Batch Processing with Error Handling
```python
import os
from src.doc_converter import DocToExcelConverter

converter = DocToExcelConverter()
input_folders = ["folder1/", "folder2/", "folder3/"]

for i, folder in enumerate(input_folders):
    try:
        output_file = f"batch_output_{i+1}.xlsx"
        result = converter.convert_directory(folder, output_file)
        print(f"✅ Successfully processed {folder}: {len(result)} questions")
    except Exception as e:
        print(f"❌ Error processing {folder}: {e}")
```

## 📝 Document Format

Your Word documents should follow this structure:

```
Question 1-- [Optional Topic]
This is the question stem. It can span multiple lines and contain 
formatting, images, and complex text.

A. First option text
B. Second option text  
C. Third option text
D. Fourth option text
E. Fifth option text (optional)

Correct Answer: C

Explanation of the Correct Answer:
This explanation can contain highlighted text for correct explanations
and normal text for incorrect explanations. The parser will separate
these automatically based on text highlighting.
```

### Key Requirements:
- **Question Headers**: Must start with "Question" followed by a number
- **Options**: Use A., B., C., D., E. format
- **Answer Format**: "Correct Answer: [Letter]" or "Answer: [Letter]"
- **Explanations**: Section titled "Explanation" or "Explanation of the Correct Answer"

## 📊 Output Format

The generated Excel file contains:

### Questions Sheet
| Column | Description |
|--------|-------------|
| MCQ_NO | Question number |
| Question | Question stem text |
| Answer1-5 | Options A through E |
| CorAns | Correct answer (1-5) |
| CorrectExplanation | Highlighted explanation text |
| IncorrectExplanation | Non-highlighted explanation text |
| Topic | Question topic (if specified) |
| Group | Question group/category |
| Type | Question type |
| SourceFile | Original filename |

### Summary Sheet
- Total questions processed
- Files processed count  
- Questions per file average
- Processing timestamp
- Error summary

## ⚙️ Configuration

### Config File (config.json)
```json
{
    "input_path": "input_docs/",
    "output_path": "output.xlsx",
    "group": "Medical Questions",
    "type": "MCQ",
    "settings": {
        "include_topics": true,
        "validate_answers": true,
        "create_summary": true,
        "handle_errors": "skip",
        "output_format": "excel"
    },
    "parsing": {
        "question_patterns": ["Question \\d+", "Q\\d+"],
        "answer_patterns": ["Correct Answer:", "Answer:"],
        "option_patterns": ["[A-E]\\.", "[A-E]\\)"]
    }
}
```

### Environment Variables
```bash
export DOC_CONVERTER_INPUT_PATH="./input/"
export DOC_CONVERTER_OUTPUT_PATH="./output/"
export DOC_CONVERTER_GROUP="Default"
export DOC_CONVERTER_TYPE="MCQ"
```

## 🔧 API Reference

### DocToExcelConverter Class

#### Methods

**`__init__(config=None)`**
- Initialize converter with optional configuration

**`convert_file(input_path, output_path, **kwargs)`**
- Convert single .docx file to Excel
- Returns: DataFrame with processed questions

**`convert_directory(input_path, output_path, **kwargs)`**
- Convert all .docx files in directory
- Returns: DataFrame with all processed questions

**`parse_docx(file_path)`**
- Parse single .docx file and extract questions
- Returns: List of question dictionaries

**`validate_questions(questions)`**
- Validate parsed questions for completeness
- Returns: Tuple of (valid_questions, errors)

### Config Class

**`Config(group, type, **settings)`**
- Configuration management for converter settings

### Utilities

**`normalize_text(text)`**
- Clean and normalize text formatting

**`extract_highlighted_text(paragraph)`**
- Extract highlighted vs normal text from Word paragraphs

## 🛠️ Advanced Usage

### Custom Parsers
```python
from src.parsers import BaseParser

class CustomMedicalParser(BaseParser):
    def parse_question_block(self, paragraphs, start, end):
        # Custom parsing logic
        return super().parse_question_block(paragraphs, start, end)

converter = DocToExcelConverter(parser=CustomMedicalParser())
```

### Batch Processing Script
```python
#!/usr/bin/env python3
import sys
from pathlib import Path
from src.batch_processor import BatchProcessor

def main():
    processor = BatchProcessor()
    
    # Process multiple directories
    input_dirs = sys.argv[1:]
    for dir_path in input_dirs:
        processor.process_directory(dir_path)

if __name__ == "__main__":
    main()
```

## 🧪 Testing

Run the test suite:
```bash
# Run all tests
python -m pytest tests/

# Run specific test
python -m pytest tests/test_converter.py

# Run with coverage
python -m pytest --cov=src tests/
```

## 📁 Project Structure

```
doc-to-excel-converter/
│
├── src/                          # Source code
│   ├── __init__.py
│   ├── doc_converter.py          # Main converter class
│   ├── parsers.py                # Document parsing logic
│   ├── config.py                 # Configuration management
│   ├── utils.py                  # Utility functions
│   └── validators.py             # Data validation
│
├── examples/                     # Usage examples
│   ├── basic_usage.py
│   ├── batch_processing.py
│   └── custom_parser.py
│
├── tests/                        # Test files
│   ├── test_converter.py
│   ├── test_parsers.py
│   └── test_utils.py
│
├── templates/                    # Document templates
│   └── question_template.docx
│
├── sample_data/                  # Sample input files
│   ├── sample1.docx
│   └── sample2.docx
│
├── docs/                         # Documentation
│   ├── api.md
│   └── examples.md
│
├── requirements.txt              # Python dependencies
├── setup.py                     # Package setup
├── config.json                  # Default configuration
├── main.py                      # CLI entry point
└── README.md                    # This file
```

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guidelines](CONTRIBUTING.md) for details.

### Development Setup
```bash
git clone https://github.com/yourusername/doc-to-excel-converter.git
cd doc-to-excel-converter
pip install -e .
pip install -r requirements-dev.txt
```

### Submitting Changes
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙋‍♂️ Support

- **Documentation**: [docs/](docs/)
- **Issues**: [GitHub Issues](https://github.com/yourusername/doc-to-excel-converter/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/doc-to-excel-converter/discussions)

## 🎯 Roadmap

- [ ] Support for more document formats (DOC, RTF)
- [ ] GUI application with drag-and-drop interface
- [ ] Integration with Learning Management Systems
- [ ] Advanced question validation and statistics
- [ ] Cloud processing capabilities
- [ ] Question bank merging and deduplication

## 📈 Changelog

### v1.0.0 (2024-01-XX)
- Initial release
- Basic .docx to Excel conversion
- Command-line interface
- Batch processing support

---

**Made with ❤️ for educators and content creators**