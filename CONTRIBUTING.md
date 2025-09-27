# Contributing to DOC to Excel Converter

Thank you for your interest in contributing to the DOC to Excel Converter! This document provides guidelines for contributing to the project.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Setup](#development-setup)
- [Making Changes](#making-changes)
- [Testing](#testing)
- [Submitting Changes](#submitting-changes)
- [Style Guidelines](#style-guidelines)
- [Reporting Issues](#reporting-issues)

## Code of Conduct

This project adheres to a code of conduct. By participating, you are expected to uphold this code:

- Use welcoming and inclusive language
- Be respectful of differing viewpoints and experiences
- Gracefully accept constructive criticism
- Focus on what is best for the community
- Show empathy towards other community members

## Getting Started

### Prerequisites

- Python 3.7 or higher
- Git
- Basic understanding of Python and document processing

### Development Setup

1. Fork the repository on GitHub
2. Clone your fork locally:
   ```bash
   git clone https://github.com/yourusername/doc-to-excel-converter.git
   cd doc-to-excel-converter
   ```

3. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

4. Install development dependencies:
   ```bash
   pip install -r requirements.txt
   pip install -e .
   ```

5. Install additional development tools:
   ```bash
   pip install pytest pytest-cov black flake8 mypy
   ```

6. Verify the setup:
   ```bash
   python -m pytest tests/
   ```

## Making Changes

### Branch Naming

Create descriptive branch names:
- `feature/add-pdf-support` for new features
- `bugfix/fix-unicode-handling` for bug fixes
- `docs/update-api-reference` for documentation
- `refactor/improve-parsing-performance` for refactoring

### Commit Messages

Write clear, descriptive commit messages:

```
Add support for PDF document parsing

- Implement PDFParser class with text extraction
- Add PDF file validation in utils module
- Update CLI to accept .pdf files as input
- Add tests for PDF parsing functionality

Fixes #123
```

Format:
- First line: Brief summary (50 characters or less)
- Blank line
- Detailed description with bullet points
- Reference any relevant issues

### Types of Contributions

#### Bug Fixes
- Check existing issues before starting
- Include test cases that reproduce the bug
- Ensure fix doesn't break existing functionality

#### New Features
- Discuss major features in an issue first
- Follow existing code patterns
- Include comprehensive tests
- Update documentation

#### Documentation
- Fix typos and improve clarity
- Add examples for new features
- Keep API documentation up to date

#### Performance Improvements
- Include benchmarks showing improvement
- Ensure changes don't affect correctness
- Document any trade-offs

## Testing

### Running Tests

```bash
# Run all tests
python -m pytest

# Run with coverage
python -m pytest --cov=src

# Run specific test file
python -m pytest tests/test_converter.py

# Run with verbose output
python -m pytest -v
```

### Writing Tests

- Place tests in the `tests/` directory
- Follow the naming convention: `test_*.py`
- Use descriptive test names: `test_parse_question_with_missing_options`
- Include both positive and negative test cases
- Test edge cases and error conditions

Example test structure:
```python
class TestQuestionParser:
    def setUp(self):
        self.config = Config()
        self.parser = QuestionParser(self.config)
    
    def test_parse_valid_question(self):
        # Test parsing of a valid question
        pass
    
    def test_parse_question_with_missing_options(self):
        # Test error handling for invalid input
        pass
```

### Test Data

- Use mock data for unit tests
- Include sample .docx files for integration tests
- Keep test files small and focused
- Don't include sensitive or copyrighted content

## Submitting Changes

### Pull Request Process

1. Update documentation for any new features
2. Add tests for new functionality
3. Ensure all tests pass
4. Update CHANGELOG.md with your changes
5. Create a pull request with:
   - Clear title describing the change
   - Detailed description of what was changed and why
   - Link to any relevant issues
   - Screenshots for UI changes (if applicable)

### Pull Request Template

```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
- [ ] Unit tests added/updated
- [ ] Integration tests added/updated
- [ ] All tests pass locally

## Checklist
- [ ] Code follows project style guidelines
- [ ] Self-review completed
- [ ] Documentation updated
- [ ] Changes work on different operating systems
```

### Review Process

- At least one maintainer must review the PR
- Address feedback constructively
- Make requested changes in new commits
- Once approved, changes will be merged

## Style Guidelines

### Python Code Style

Follow PEP 8 with these specifics:

- Line length: 88 characters (Black default)
- Use type hints for function parameters and return values
- Use docstrings for all public functions and classes
- Prefer f-strings for string formatting

```python
def parse_question_block(
    self, paragraphs: List, start_idx: int, end_idx: int
) -> Dict[str, str]:
    """
    Parse a single question block from document paragraphs.
    
    Args:
        paragraphs: List of document paragraphs
        start_idx: Starting paragraph index
        end_idx: Ending paragraph index
        
    Returns:
        Dictionary containing parsed question data
        
    Raises:
        ValueError: If question format is invalid
    """
    pass
```

### Documentation Style

- Use Markdown for documentation files
- Include code examples in docstrings
- Keep README up to date with new features
- Use clear, concise language

### Code Organization

- Keep functions focused and small
- Use meaningful variable names
- Group related functionality in modules
- Follow the existing project structure

## Reporting Issues

### Bug Reports

Include the following information:

1. **Environment**: OS, Python version, package versions
2. **Steps to reproduce**: Minimal example that shows the bug
3. **Expected behavior**: What should happen
4. **Actual behavior**: What actually happens
5. **Sample files**: Attach sample .docx files if relevant (remove sensitive data)

### Feature Requests

Include:

1. **Use case**: Why is this feature needed?
2. **Proposed solution**: How should it work?
3. **Alternatives considered**: Other approaches you've thought of
4. **Additional context**: Screenshots, mockups, etc.

### Issue Labels

- `bug`: Something isn't working
- `enhancement`: New feature or improvement
- `documentation`: Documentation improvements
- `good first issue`: Good for newcomers
- `help wanted`: Extra attention needed

## Development Tips

### Local Testing

- Test with various document formats
- Verify cross-platform compatibility
- Test error conditions and edge cases
- Use the example scripts to verify changes

### Performance Considerations

- Profile code changes for large documents
- Consider memory usage for batch processing
- Test with various document sizes

### Documentation

- Update docstrings when changing function signatures
- Include examples in documentation
- Test that examples actually work

## Getting Help

- Check existing issues and documentation first
- Ask questions in GitHub Discussions
- Join our community chat (if available)
- Contact maintainers for major architectural questions

## Recognition

Contributors are recognized in:
- CONTRIBUTORS.md file
- Release notes for significant contributions
- GitHub contributor statistics

Thank you for contributing to DOC to Excel Converter!