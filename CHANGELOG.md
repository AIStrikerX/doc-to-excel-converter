# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]
### Added
- Planned GUI interface with drag-and-drop functionality
- Support for additional document formats (DOC, RTF)
- Advanced question validation and statistics
- Cloud processing capabilities
- Question bank merging and deduplication features

## [1.0.0] - 2024-01-XX
### Added
- Initial release of DOC to Excel Converter
- Core conversion functionality for .docx to .xlsx
- Bulk processing support for directories
- Smart parsing of MCQ question blocks
- Highlighting detection for correct vs incorrect explanations  
- Flexible configuration system with JSON config files
- Command-line interface with comprehensive options
- Python API for programmatic use
- Robust error handling and validation
- Comprehensive test suite
- GitHub Actions CI/CD pipeline
- Complete documentation and examples

### Features
- **Question Parsing**: Automatic detection of question numbers, options (A-E), correct answers, and explanations
- **Text Processing**: Separation of highlighted (correct) vs normal (incorrect) explanation text
- **Batch Processing**: Process single files or entire directories of Word documents
- **Excel Output**: Structured output with customizable column mapping and summary sheets
- **Configuration**: Multiple configuration options via JSON files, CLI arguments, and environment variables
- **Validation**: Data quality checks and error reporting for parsed questions
- **Extensibility**: Modular design with pluggable parsers and validators

### Technical Details
- Python 3.7+ compatibility
- Dependencies: pandas, python-docx, openpyxl
- Cross-platform support (Windows, macOS, Linux)
- Comprehensive error handling and logging
- Type hints for better code quality
- Unit tests with pytest
- Code formatting with Black
- Linting with flake8
- Type checking with mypy

### Documentation
- Complete README with usage examples
- API reference documentation
- Installation guide
- Contributing guidelines
- Sample data and templates
- GitHub issue templates

### Development
- Professional repository structure
- CI/CD with GitHub Actions
- Code quality checks (Black, flake8, mypy)
- Automated testing across multiple Python versions
- MIT license for open source use