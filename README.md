# Glucose Data Converter

A Python application that converts Contour Plus glucose meter CSV exports into beautifully formatted, color-coded XLSX files.

## Features

✅ **Automatic CSV to XLSX conversion** with professional formatting  
✅ **Color-coded glucose levels** for easy visualization:
- 🟦 Pale blue: Low glucose (< 4.0 mmol/L)
- 🟥 Red: High glucose (> 11.9 mmol/L)  
- 🟪 Light purple: Very high glucose (> 17.9 mmol/L)

✅ **Two interfaces available**:
- Command-line interface for automation
- GUI with drag-and-drop support for ease of use

✅ **Smart features**:
- Auto-detect latest CSV in Downloads folder
- Customizable glucose thresholds
- Statistics summary (average, min, max, distribution)
- Configuration file support
- Automatic file opening after conversion

## Installation Guide for Windows

### Method 1: Using the Executable (Easiest)

1. **Download the files** from this repository
2. **Run the build script** to create the executable:
   ```
   python build_exe.py
   ```
3. **Find the executable** in the `dist` folder: `GlucoseConverter.exe`
4. **Double-click** `GlucoseConverter.exe` to run the application

No Python installation required for the executable!

### Method 2: Running from Source Code

#### Step 1: Install Python

1. Download Python from [python.org](https://www.python.org/downloads/)
2. During installation, **CHECK** "Add Python to PATH"
3. Click "Install Now"

#### Step 2: Install Dependencies

Open Command Prompt (Win+R, type `cmd`, press Enter) and run:

```bash
cd path\to\glucose_results
pip install -r requirements.txt
```

Or install manually:
```bash
pip install openpyxl
pip install tkinterdnd2  # Optional, for drag-and-drop
```

#### Step 3: Run the Application

**Option A: GUI Version (Recommended)**
```bash
python glucose_converter_gui.py
```

**Option B: Command Line Version**
```bash
python glucose_converter.py ContourCSVReport.csv
```

## Usage Instructions

### GUI Application

1. **Launch** the GUI:
   - Double-click `GlucoseConverter.exe` (if using executable)
   - Or run: `python glucose_converter_gui.py`

2. **Load your CSV file**:
   - **Drag and drop** your CSV file onto the application window
   - Or click "Browse" to select the file
   - Or click "Auto-detect Latest CSV" to find the newest file in Downloads

3. **Configure settings** (optional):
   - Adjust glucose thresholds if needed
   - Select output folder (default: same as input file)
   - Check "Open file after conversion" for automatic opening

4. **Click "Convert to XLSX"**

5. **Find your formatted file** in the output folder with timestamp

### Command Line Usage

**Basic conversion:**
```bash
python glucose_converter.py path\to\your\file.csv
```

**Auto-detect latest CSV in Downloads:**
```bash
python glucose_converter.py --auto-detect
```

**Specify output file:**
```bash
python glucose_converter.py input.csv -o output.xlsx
```

**Use configuration file:**
```bash
python glucose_converter.py input.csv -c config.ini
```

**Create sample configuration:**
```bash
python glucose_converter.py --create-config
```

## Configuration File

Create a `glucose_config.ini` file to customize settings:

```ini
[Settings]
# Output folder (leave empty for same as input)
output_folder = C:\Users\YourName\Documents\Glucose

# Auto-open file after conversion
auto_open = true

# Glucose thresholds (mmol/L)
low_threshold = 4.0
high_threshold = 11.9
very_high_threshold = 17.9
```

## Output Format

The generated XLSX file includes:

1. **Formatted Data Table**:
   - Date/Time in readable format (d.m.Y H:i)
   - Color-coded glucose values
   - Meal markers and notes
   - Professional borders and styling

2. **Statistics Summary**:
   - Total number of readings
   - Average, minimum, and maximum glucose
   - Distribution across ranges (low/normal/high/very high)
   - Percentage calculations

## Troubleshooting

### "Python is not recognized"
- Reinstall Python and ensure "Add to PATH" is checked
- Or restart Command Prompt after installation

### "Module not found" error
```bash
pip install --upgrade pip
pip install openpyxl tkinterdnd2
```

### Drag and drop not working
- Install tkinterdnd2: `pip install tkinterdnd2`
- Or use the "Browse" button instead

### Permission denied error
- Run as Administrator
- Or save to a different folder with write permissions

### CSV file not loading
- Ensure the CSV is from Contour Plus app
- Check that the file is not open in another program
- Verify the CSV format matches expected structure

## File Structure

```
glucose_results/
├── glucose_converter.py        # Main converter script
├── glucose_converter_gui.py    # GUI application
├── build_exe.py                # Executable builder
├── requirements.txt            # Python dependencies
├── README.md                   # This file
└── dist/
    └── GlucoseConverter.exe    # Standalone executable (after build)
```

## Tips for Best Results

1. **Export from Contour Plus app**:
   - Open Contour Plus app
   - Go to Reports/Export
   - Select CSV format
   - Save to Downloads folder

2. **Use Auto-detect**:
   - Save CSV to Downloads with default name
   - Click "Auto-detect Latest CSV" in GUI
   - Or use `--auto-detect` in command line

3. **Batch Processing**:
   - Use command line version in scripts
   - Process multiple files with a loop

4. **Custom Thresholds**:
   - Adjust thresholds based on your target ranges
   - Save preferences in config file

## Support

For issues or questions:
1. Check the Troubleshooting section above
2. Ensure you have the latest version
3. Verify your CSV file format matches Contour Plus export

## License

This tool is provided as-is for personal use with Contour Plus glucose data.

---

**Note**: This tool is designed specifically for Contour Plus CSV exports. Other formats may require modification of the parsing logic.