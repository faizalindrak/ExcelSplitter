# Excel Splitter (Template-based)

A powerful GUI application for splitting Excel files based on unique values in a specified column, with template-based rendering to preserve formatting and optional PDF export functionality.

## 📋 Description

Excel Splitter is a Python-based desktop application that allows you to split large Excel files into multiple smaller files based on unique values in a key column. The application uses template-based rendering to maintain the original formatting, headers, and styling from your template file.

### Key Features

- **Template-based Splitting**: Preserve original Excel formatting and styling
- **Multiple Output Formats**: Generate both Excel (.xlsx) and PDF files
- **Flexible Key Column Selection**: Use column names or index numbers
- **Progress Tracking**: Real-time progress bar and status logging
- **Configuration Management**: Save and load settings via INI files
- **Error Handling**: Comprehensive error handling with detailed logging
- **Cross-platform**: Works on Windows, macOS, and Linux

## 🚀 Installation

### Prerequisites

- Python 3.10 or higher
- pip package manager (or uv for faster installation)

### Option 1: Using Pre-built Executable (Recommended)

1. Download the latest release from the [releases page](https://https://github.com/faizalindrak/ExcelSplitter/releases)
2. Extract the ZIP file to your desired location
3. Run `ExcelSplitter.exe`
4. The application is ready to use!

### Option 2: Running from Source

1. **Clone or download** this repository
2. **Navigate** to the project directory
3. **Install dependencies** using one of the following methods:

#### Using pip:
```bash
python -m venv .venv
.venv\Scripts\activate  # On Windows
# or source .venv/bin/activate  # On macOS/Linux

pip install -r requirements.txt
python main.py
```

#### Using uv (Recommended for faster installation):
```bash
# Install uv if not already installed
# On Windows PowerShell:
# iwr -useb https://astral.sh/uv/install.ps1 | iex

# Create virtual environment and install dependencies
uv venv .venv
uv pip sync requirements.txt

# Run the application
uv run python main.py
```

## 🛠️ Building from Source

### Quick Build (Windows)

Run one of the provided build scripts:

#### Using build.cmd (Auto-detects pip/uv):
```cmd
build.cmd
```

#### Using build-uv.bat (uv only):
```cmd
build-uv.bat
```

### Manual Build

1. **Install PyInstaller**:
   ```bash
   pip install pyinstaller
   ```

2. **Build the executable**:
   ```bash
   # Using the spec file (recommended)
   pyinstaller main.spec

   # Or direct command (alternative)
   pyinstaller --noconsole --onefile --clean --name ExcelSplitter main.py
   ```

3. **Find the executable**:
   - The built application will be in `dist/ExcelSplitter/ExcelSplitter.exe`
   - Copy this file to your desired location for distribution

### Build Configuration

The `main.spec` file contains build configuration:
- **One-file mode**: Single executable file
- **No console**: GUI application without console window
- **Icon support**: Add your custom icon by setting `ICON_PATH` in `main.spec`
- **UPX compression**: Enabled for smaller file size (requires UPX installed)

## 📖 Usage Guide

### Getting Started

1. **Launch** the application
2. **Select Source Excel**: Click "Browse..." to choose your data file
3. **Load Sheets**: Click "Load Sheets" to see available worksheets
4. **Select Sheet**: Choose the worksheet containing your data
5. **Load Headers**: Click "Load Headers" to see column options
6. **Select Key Column**: Choose the column to split by (by name or index)
7. **Select Template**: Choose a template Excel file for formatting
8. **Choose Output Folder**: Select where to save split files
9. **Configure Options**:
   - Header rows count (default: 5)
   - PDF engine (reportlab, libreoffice, or none)
   - LibreOffice path (if using libreoffice PDF engine)
10. **Generate**: Click "Generate" to start the splitting process

### Input Files

#### Source Excel File
- Contains the data to be split
- Can be .xlsx, .xls, .xlsm, or .xlsb format
- Should have a column with unique values for splitting

#### Template Excel File
- Used for formatting the output files
- Should be a .xlsx file with desired styling
- Header rows will be preserved in output files
- Column order can be customized by template structure

### Output Files

For each unique value in the key column, the application creates:
- **Excel file**: `SCHEDULE {key_value}.xlsx`
- **PDF file** (optional): `SCHEDULE {key_value}.pdf`

### PDF Export Options

#### ReportLab (Pure Python)
- **Pros**: No external dependencies, fast
- **Cons**: Limited formatting support
- **Installation**: `pip install reportlab`

#### LibreOffice
- **Pros**: Better formatting preservation, supports complex layouts
- **Cons**: Requires LibreOffice installation
- **Installation**: Download from [libreoffice.org](https://www.libreoffice.org/)
- **Auto-detection**: Application searches common installation paths

## ⚙️ Configuration

### Saving Settings
- Click "Save .ini" to save current configuration
- Settings include all file paths, options, and preferences
- Useful for repeated use with similar files

### Loading Settings
- Click "Load .ini" to restore previous configuration
- Quickly resume work with saved settings

## 🔧 Advanced Features

### Error Handling
- Comprehensive logging in the status text area
- Detailed error messages for troubleshooting
- Handles common Excel issues (corrupted files, permission errors)

### Performance Optimizations
- Efficient memory usage for large files
- Progress tracking for long operations
- Threaded processing to keep UI responsive

### Data Processing
- Handles categorical data conversion
- Supports various Excel formats
- Maintains data integrity during splitting

## 🐛 Troubleshooting

### Common Issues

#### "ReportLab not available"
- Install ReportLab: `pip install reportlab`
- Or switch to "libreoffice" or "none" PDF engine

#### "LibreOffice not found"
- Install LibreOffice from [libreoffice.org](https://www.libreoffice.org/)
- Or manually browse to `soffice.exe` location
- Or switch to "reportlab" PDF engine

#### "Permission denied" errors
- Close any open Excel files
- Run as administrator if needed
- Check file/folder permissions

#### "Memory error" with large files
- Close other applications
- Process files in smaller batches
- Consider using 64-bit Python

#### Build fails
- Ensure all dependencies are installed
- Check Python version (3.10+ required)
- Try running `pyinstaller --clean` to clear cache

### Debug Mode
Enable detailed logging by checking the status text area during processing. The application provides comprehensive debug information including:
- File size and accessibility
- Data type detection
- Processing time
- Error details

## 📋 Requirements

### Core Dependencies
- `customtkinter>=5.2.2` - Modern GUI framework
- `pandas>=2.3.2` - Data manipulation
- `openpyxl>=3.1.5` - Excel file handling
- `reportlab>=4.4.4` - PDF generation (optional)

### Build Dependencies
- `pyinstaller>=6.16.0` - Application packaging
- `pyinstaller-hooks-contrib>=2025.8` - Additional hooks

### Optional Dependencies
- `uv` - Faster Python package management
- `UPX` - Executable compression
- `LibreOffice` - Alternative PDF engine

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- Built with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
- Powered by [Pandas](https://pandas.pydata.org/) and [OpenPyXL](https://openpyxl.readthedocs.io/)
- PDF functionality via [ReportLab](https://www.reportlab.com/) and LibreOffice

## 📞 Support

For support and questions:
- Create an issue on GitHub
- Check the troubleshooting section above
- Review the debug logs in the application

---

**Version**: 1.0.0
**Last Updated**: 2024
**Python**: 3.10+