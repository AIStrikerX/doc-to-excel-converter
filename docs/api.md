# API Reference

## DocToExcelConverter Class

The main class for converting Word documents to Excel format.

### Constructor

```python
DocToExcelConverter(config=None, parser=None)
```

**Parameters:**
- `config` (Config, optional): Configuration object with converter settings
- `parser` (QuestionParser, optional): Custom parser for question extraction

### Methods

#### convert_file()

```python
convert_file(input_path, output_path, **kwargs) -> pd.DataFrame
```

Convert a single .docx file to Excel format.

**Parameters:**
- `input_path` (str): Path to the input .docx file
- `output_path` (str): Path for the output Excel file
- `**kwargs`: Additional parameters to override config

**Returns:**
- `pd.DataFrame`: DataFrame containing the processed questions

**Raises:**
- `FileNotFoundError`: If input file doesn't exist
- `ValueError`: If file format is not supported

#### convert_directory()

```python
convert_directory(input_path, output_path, **kwargs) -> pd.DataFrame
```

Convert all .docx files in a directory to a single Excel file.

**Parameters:**
- `input_path` (str): Path to the directory containing .docx files
- `output_path` (str): Path for the output Excel file
- `**kwargs`: Additional parameters to override config

**Returns:**
- `pd.DataFrame`: DataFrame containing all processed questions

#### parse_docx()

```python
parse_docx(file_path) -> List[Dict[str, str]]
```

Parse a single .docx file and extract questions.

**Parameters:**
- `file_path` (str): Path to the .docx file

**Returns:**
- `List[Dict[str, str]]`: List of question dictionaries

#### get_stats()

```python
get_stats() -> Dict[str, Union[int, List[str]]]
```

Get processing statistics.

**Returns:**
- `Dict`: Dictionary with processing statistics

## Config Class

Configuration management for converter settings.

### Constructor

```python
Config(group="General", question_type="MCQ", validate_answers=True, ...)
```

**Parameters:**
- `group` (str): Default group/category for questions
- `question_type` (str): Default type for questions
- `validate_answers` (bool): Whether to validate parsed answers
- `include_topics` (bool): Whether to extract question topics
- `create_summary` (bool): Whether to create summary sheet
- `verbose` (bool): Enable verbose logging
- `quiet` (bool): Suppress output except errors

### Methods

#### from_dict()

```python
@classmethod
from_dict(config_dict) -> Config
```

Create Config instance from dictionary.

#### from_file()

```python
@classmethod  
from_file(config_path) -> Config
```

Load configuration from JSON file.

#### validate()

```python
validate() -> List[str]
```

Validate configuration settings and return list of errors.

## QuestionParser Class

Parser for extracting MCQ questions from Word documents.

### Constructor

```python
QuestionParser(config)
```

**Parameters:**
- `config` (Config): Configuration object

### Methods

#### parse_document()

```python
parse_document(document, file_path) -> List[Dict[str, str]]
```

Parse a Word document and extract all questions.

## QuestionValidator Class

Validator for MCQ questions.

### Constructor

```python
QuestionValidator(config)
```

### Methods

#### validate_questions()

```python
validate_questions(questions) -> Tuple[List[Dict[str, str]], List[str]]
```

Validate a list of questions and return valid ones with error messages.

#### validate_single_question()

```python
validate_single_question(question) -> List[ValidationError]
```

Validate a single question and return list of validation errors.

## Utility Functions

### normalize_text()

```python
normalize_text(text) -> str
```

Normalize and clean text by removing extra whitespace and special characters.

### validate_paths()

```python
validate_paths(input_path, output_path) -> Tuple[bool, bool, str]
```

Validate input and output paths.

### setup_logging()

```python
setup_logging(verbose=False, quiet=False) -> logging.Logger
```

Set up logging configuration.

## Constants

### Answer Mapping

```python
LETTER_TO_NUM = {"A": "1", "B": "2", "C": "3", "D": "4", "E": "5"}
```

### Predefined Configurations

```python
MEDICAL_CONFIG = Config(
    group="Medical Questions",
    question_type="Medical Board Review",
    validate_answers=True,
    include_topics=True
)

EDUCATIONAL_CONFIG = Config(
    group="Educational Content",
    question_type="Academic Quiz", 
    validate_answers=True,
    include_topics=True
)

SIMPLE_CONFIG = Config(
    group="General",
    question_type="MCQ",
    validate_answers=False,
    include_topics=False
)
```

## Data Structures

### Question Dictionary Format

Each parsed question is represented as a dictionary with the following structure:

```python
{
    "MCQ_NO": "1",                    # Question number
    "Question": "Question text...",    # Question stem
    "Answer1": "Option A text",        # Option A
    "Answer2": "Option B text",        # Option B  
    "Answer3": "Option C text",        # Option C
    "Answer4": "Option D text",        # Option D
    "Answer5": "Option E text",        # Option E (optional)
    "CorAns": "3",                    # Correct answer (1-5)
    "CorrectExplanation": "...",      # Highlighted explanation
    "IncorrectExplanation": "...",    # Non-highlighted explanation
    "Topic": "Subject topic",         # Question topic
    "Group": "Question group",        # Question group/category
    "Type": "MCQ",                    # Question type
    "SourceFile": "filename.docx"     # Source filename
}
```

### Excel Output Format

The generated Excel file contains two sheets:

#### Questions Sheet

| Column | Description |
|--------|-------------|
| MCQ_NO | Question number |
| Question | Question stem text |
| Answer1 | Option A |
| CorAns | Correct answer (1-5) |
| Answer2 | Option B |
| Answer3 | Option C |
| Answer4 | Option D |
| Answer5 | Option E |
| CorrectExplanation | Highlighted explanation text |
| IncorrectExplanation | Non-highlighted explanation text |
| Topic | Question topic |
| Group | Question group |
| Type | Question type |
| SourceFile | Source filename |

#### Summary Sheet

| Metric | Value |
|--------|-------|
| Total Questions | Number of questions processed |
| Files Processed | Number of files processed |
| Questions per File | Average questions per file |
| Processing Date | When processing occurred |

## Error Handling

The converter uses a comprehensive error handling system:

### Exception Types

- `FileNotFoundError`: Input file or directory not found
- `ValueError`: Invalid file format or configuration
- `PermissionError`: Cannot write to output location

### Validation Errors

The `ValidationError` class represents validation issues:

```python
@dataclass
class ValidationError:
    question_id: str      # Question identifier
    error_type: str       # Type of error
    message: str          # Human-readable message
    severity: str         # "error", "warning", or "info"
```

### Error Recovery

The converter continues processing when encountering non-critical errors and provides detailed error reporting in the processing statistics.