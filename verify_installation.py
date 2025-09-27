#!/usr/bin/env python3
"""
Installation verification script for DOC to Excel Converter

This script checks if the package can be properly installed and imported.
Run this after installation to verify everything is working correctly.

Usage: python verify_installation.py
"""

import sys
import os
import importlib

def check_python_version():
    """Check if Python version is supported."""
    version = sys.version_info
    if version >= (3, 7):
        print(f"✅ Python {version.major}.{version.minor}.{version.micro} (supported)")
        return True
    else:
        print(f"❌ Python {version.major}.{version.minor}.{version.micro} (requires >= 3.7)")
        return False

def check_dependencies():
    """Check if required dependencies are installed."""
    required_packages = [
        'pandas',
        'docx',  # python-docx
        'openpyxl'
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            importlib.import_module(package)
            print(f"✅ {package}")
        except ImportError:
            print(f"❌ {package} (not installed)")
            missing_packages.append(package)
    
    return missing_packages

def check_package_structure():
    """Check if package structure is correct."""
    required_files = [
        'src/__init__.py',
        'src/doc_converter.py',
        'src/config.py',
        'src/parsers.py',
        'src/utils.py',
        'src/validators.py',
        'main.py',
        'setup.py',
        'requirements.txt'
    ]
    
    missing_files = []
    
    for file_path in required_files:
        if os.path.exists(file_path):
            print(f"✅ {file_path}")
        else:
            print(f"❌ {file_path} (missing)")
            missing_files.append(file_path)
    
    return missing_files

def test_import():
    """Test if the main module can be imported."""
    try:
        # Add src to path
        sys.path.insert(0, 'src')
        
        # Try importing main components
        from config import Config
        print("✅ Config module imported")
        
        # Test basic functionality
        config = Config()
        print("✅ Config instance created")
        
        return True
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False
    except Exception as e:
        print(f"❌ Error testing imports: {e}")
        return False

def main():
    """Run all verification checks."""
    print("DOC to Excel Converter - Installation Verification")
    print("=" * 55)
    
    # Check Python version
    print("\n1. Checking Python version...")
    python_ok = check_python_version()
    
    # Check package structure
    print("\n2. Checking package structure...")
    missing_files = check_package_structure()
    
    # Check dependencies
    print("\n3. Checking dependencies...")
    missing_packages = check_dependencies()
    
    # Test imports
    print("\n4. Testing imports...")
    import_ok = test_import()
    
    # Summary
    print("\n" + "=" * 55)
    print("VERIFICATION SUMMARY")
    print("=" * 55)
    
    if python_ok and not missing_files and import_ok:
        if not missing_packages:
            print("🎉 Installation is complete and working!")
            print("\nYou can now use the converter:")
            print("  python main.py --help")
            print("  python examples/basic_usage.py")
        else:
            print("⚠️  Installation is mostly working, but some dependencies are missing.")
            print("\nTo install missing dependencies:")
            print("  pip install -r requirements.txt")
            print(f"\nMissing packages: {', '.join(missing_packages)}")
    else:
        print("❌ Installation has issues that need to be resolved.")
        
        if not python_ok:
            print("  - Upgrade Python to version 3.7 or higher")
        
        if missing_files:
            print(f"  - Missing files: {', '.join(missing_files)}")
        
        if not import_ok:
            print("  - Module import failed (check Python path and dependencies)")
    
    print("\n" + "=" * 55)

if __name__ == "__main__":
    main()