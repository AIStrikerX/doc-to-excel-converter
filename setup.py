"""
Setup script for DOC to Excel Converter

This allows the package to be installed with pip install -e .
"""

from setuptools import setup, find_packages
import os

# Read README for long description
def read_readme():
    try:
        with open("README.md", "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        return "DOC to Excel Converter - Convert Word documents containing MCQs to Excel format"

# Read requirements
def read_requirements():
    try:
        with open("requirements.txt", "r", encoding="utf-8") as f:
            return [line.strip() for line in f if line.strip() and not line.startswith("#")]
    except FileNotFoundError:
        return ["pandas>=1.3.0", "python-docx>=0.8.11", "openpyxl>=3.0.9"]

setup(
    name="doc-to-excel-converter",
    version="1.0.0",
    
    # Package information
    description="Convert Word documents containing MCQs to Excel format",
    long_description=read_readme(),
    long_description_content_type="text/markdown",
    
    # Author information
    author="AIStrikerX",
    author_email="contact@aistrikerx.com",
    url="https://github.com/AIStrikerX/doc-to-excel-converter",
    
    # License and classifiers
    license="MIT",
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Education",
        "Intended Audience :: Healthcare Industry",
        "Topic :: Education",
        "Topic :: Office/Business",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: OS Independent",
    ],
    
    # Package configuration
    packages=find_packages(),
    python_requires=">=3.7",
    install_requires=read_requirements(),
    
    # Extra dependencies
    extras_require={
        "dev": [
            "pytest>=6.0.0",
            "pytest-cov>=2.12.0",
            "black>=21.0.0",
            "flake8>=3.9.0",
            "mypy>=0.910",
        ],
        "docs": [
            "sphinx>=4.0.0",
            "sphinx-rtd-theme>=0.5.0",
        ],
        "progress": [
            "tqdm>=4.62.0",
            "colorama>=0.4.4",
        ]
    },
    
    # Entry points for command-line tools
    entry_points={
        "console_scripts": [
            "doc2excel=main:main",
            "docx2xlsx=main:main",
        ],
    },
    
    # Package data
    package_data={
        "": ["*.md", "*.txt", "*.json"],
        "templates": ["*.docx"],
        "sample_data": ["*.docx"],
    },
    include_package_data=True,
    
    # Project URLs
    project_urls={
        "Bug Reports": "https://github.com/AIStrikerX/doc-to-excel-converter/issues",
        "Source": "https://github.com/AIStrikerX/doc-to-excel-converter",
        "Documentation": "https://github.com/AIStrikerX/doc-to-excel-converter/blob/main/docs/api.md",
    },
    
    # Keywords for PyPI search
    keywords="docx excel converter mcq questions medical education word document parser",
)