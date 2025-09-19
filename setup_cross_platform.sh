#!/bin/bash

echo "========================================================="
echo " Glucose Converter - Cross-Platform Setup"
echo "========================================================="
echo ""

# Detect OS
OS=""
if [[ "$OSTYPE" == "linux-gnu"* ]]; then
    OS="Linux"
elif [[ "$OSTYPE" == "darwin"* ]]; then
    OS="macOS"
elif [[ "$OSTYPE" == "cygwin" ]] || [[ "$OSTYPE" == "msys" ]] || [[ "$OSTYPE" == "win32" ]]; then
    OS="Windows"
else
    OS="Unknown"
fi

echo "Detected OS: $OS"
echo ""

# Check Python installation
if command -v python3 &> /dev/null; then
    echo "✓ Python3 found: $(python3 --version)"
elif command -v python &> /dev/null; then
    echo "✓ Python found: $(python --version)"
    # Create alias for consistency
    alias python3=python
else
    echo "✗ Python not found. Please install Python 3.8 or later"
    echo "  - Linux: sudo apt install python3 python3-pip"
    echo "  - macOS: brew install python3"
    echo "  - Windows: Download from https://python.org"
    exit 1
fi

echo ""
echo "Installing dependencies..."
echo ""

# Try different installation methods based on OS
if [[ "$OS" == "Linux" ]]; then
    echo "Attempting Linux installation methods..."
    
    # Try with --user flag first
    if python3 -m pip install --user openpyxl; then
        echo "✓ Installed with --user flag"
    # Try system package manager
    elif command -v apt &> /dev/null; then
        echo "Trying system package manager..."
        sudo apt install python3-openpyxl -y
    elif command -v dnf &> /dev/null; then
        sudo dnf install python3-openpyxl -y
    elif command -v pacman &> /dev/null; then
        sudo pacman -S python-openpyxl --noconfirm
    else
        echo "Please install manually: pip install openpyxl"
    fi
    
elif [[ "$OS" == "macOS" ]]; then
    echo "Installing for macOS..."
    python3 -m pip install openpyxl
    
elif [[ "$OS" == "Windows" ]]; then
    echo "Installing for Windows..."
    python -m pip install openpyxl
fi

# Optional: Install tkinterdnd2 for drag-and-drop
echo ""
echo "Installing optional drag-and-drop support..."
python3 -m pip install tkinterdnd2 --user 2>/dev/null || echo "Note: Drag-and-drop support not available"

echo ""
echo "========================================================="
echo " Setup Complete!"
echo "========================================================="
echo ""
echo "You can now run the application:"
echo "  - Enhanced CLI: python3 glucose_converter_enhanced.py --help"
echo "  - Enhanced GUI: python3 glucose_converter_gui_enhanced.py"
echo "  - Run tests: python3 test_glucose_converter.py"
echo ""
echo "Quick examples:"
echo "  python3 glucose_converter_enhanced.py --auto-detect"
echo "  python3 glucose_converter_enhanced.py input.csv --incremental"
echo "  python3 glucose_converter_enhanced.py --list-templates"
echo ""