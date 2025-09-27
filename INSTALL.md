# Installation and Setup Guide

## Quick Start

1. **Clone the Repository**
   ```bash
   git clone https://github.com/yourusername/doc-to-excel-converter.git
   cd doc-to-excel-converter
   ```

2. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Test the Installation**
   ```bash
   python enhanced_converter.py --input sample_data/sample_medical_questions.docx --output test_output.xlsx
   ```

## Detailed Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package installer)

### Step-by-Step Installation

1. **Download the Repository**
   - Option A: Clone with Git
     ```bash
     git clone https://github.com/yourusername/doc-to-excel-converter.git
     ```
   - Option B: Download ZIP from GitHub and extract

2. **Navigate to the Directory**
   ```bash
   cd doc-to-excel-converter
   ```

3. **Create Virtual Environment (Recommended)**
   ```bash
   python -m venv venv
   
   # On Windows:
   venv\Scripts\activate
   
   # On macOS/Linux:
   source venv/bin/activate
   ```

4. **Install Required Packages**
   ```bash
   pip install -r requirements.txt
   ```

5. **Install the Package (Optional)**
   ```bash
   pip install -e .
   ```

### Verify Installation

Run the test suite:
```bash
python tests/test_converter.py
```

Try the enhanced converter:
```bash
python enhanced_converter.py --help
```

## Usage Examples

### Command Line Usage

**Convert Single File:**
```bash
python enhanced_converter.py --input "questions.docx" --output "result.xlsx"
```

**Convert Directory:**
```bash
python enhanced_converter.py --input "docs/" --output "all_questions.xlsx" --group "Medicine" --type "Board Review"
```

**Using Main CLI:**
```bash
python main.py --input "medical_docs/" --output "medical_questions.xlsx" --validate --verbose
```

### Python API Usage

```python
from src.doc_converter import DocToExcelConverter
from src.config import MEDICAL_CONFIG

# Basic usage
converter = DocToExcelConverter()
df = converter.convert_file("questions.docx", "output.xlsx")

# With custom configuration
converter = DocToExcelConverter(MEDICAL_CONFIG)
df = converter.convert_directory("docs/", "output.xlsx")
```

## Configuration

### Using Configuration File

Create a `config.json` file:
```json
{
  "input_path": "./input_docs/",
  "output_path": "./output.xlsx",
  "question_settings": {
    "group": "Medical Questions",
    "type": "Board Review",
    "validate_answers": true
  }
}
```

Use with:
```bash
python main.py --config config.json
```

### Environment Variables

Set environment variables:
```bash
export DOC_CONVERTER_GROUP="Medical"
export DOC_CONVERTER_TYPE="MCQ"
export DOC_CONVERTER_VERBOSE="true"
```

## Troubleshooting

### Common Issues

1. **ImportError: No module named 'docx'**
   ```bash
   pip install python-docx
   ```

2. **ImportError: No module named 'pandas'**
   ```bash
   pip install pandas openpyxl
   ```

3. **Permission Error when writing output**
   - Check that the output directory exists and is writable
   - Close any Excel files that might have the output file open

4. **No questions found in document**
   - Verify the document format matches the expected structure
   - Check that questions start with "Question 1", "Question 2", etc.
   - Ensure options are formatted as "A. option text", "B. option text", etc.

### Getting Help

- Check the [API documentation](docs/api.md)
- Look at [example files](examples/)
- Run with `--verbose` flag for detailed output
- Create an issue on GitHub for bugs or questions

## Development Setup

For contributors:

1. **Fork and Clone**
   ```bash
   git clone https://github.com/yourusername/doc-to-excel-converter.git
   cd doc-to-excel-converter
   ```

2. **Install Development Dependencies**
   ```bash
   pip install -r requirements.txt
   pip install pytest pytest-cov black flake8 mypy
   ```

3. **Run Tests**
   ```bash
   python -m pytest tests/
   ```

4. **Format Code**
   ```bash
   black src/ tests/ examples/
   ```

## Performance Tips

- For large batches of files, use the directory processing mode
- Enable validation only when needed (it slows processing)
- Use the `--max-files` option to limit processing during testing
- Close other applications when processing very large documents

## Supported Formats

### Input Formats
- `.docx` (Word 2007+)

### Output Formats  
- `.xlsx` (Excel 2007+)

### Question Format Requirements
- Questions must start with "Question N" where N is a number
- Options must be formatted as A., B., C., D., E.
- Correct answer must be specified as "Correct Answer: X" or "Answer: X"
- Explanations should follow "Explanation of the Correct Answer" header