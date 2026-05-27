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

1. Download the latest release from the [releases page](https://github.com/faizalindrak/ExcelSplitter/releases)
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
uv pip install --upgrade -r requirements.txt

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
   - The built application will be in `dist/ExcelSplitter.exe`
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
7. **Choose Template Option**:
   - **Use Template File**: choose a separate template workbook and map template columns to source columns
   - **Use Source as Template**: split the selected source worksheet directly and keep its layout
8. **Review Column Mapping**: Click "Auto Map" and manually map any required template columns that were not detected
9. **Choose Output Folder**: Select where to save split files
10. **Configure Options**:
   - Header rows count (default: 5)
   - PDF engine (xlwings, libreoffice, or none)
   - LibreOffice path (if using libreoffice PDF engine)
11. **Generate**: Click "Generate" to start the splitting process

### Input Files

#### Source Excel File
- Contains the data to be split
- Can be .xlsx, .xls, .xlsm, or .xlsb format
- Should have a column with unique values for splitting

#### Template Excel File
- Used for formatting output files when using **Use Template File**
- Should be a .xlsx file with desired styling
- Header rows will be preserved in output files
- Column order comes from the template header row
- The template header row must contain column names so mapping can be completed
- If template headers do not match source headers, complete the Column Mapping card before generating

#### Source as Template
- Uses the selected source worksheet as the output template
- Outputs one workbook per key containing only that worksheet
- Preserves worksheet layout and styles as much as openpyxl supports
- Does not require a separate template file or column mapping

### Output Files

For each unique value in the key column, the application creates:
- **Excel file**: `{prefix} {key_value} {suffix}.xlsx`
- **PDF file** (optional): `{prefix} {key_value} {suffix}.pdf`

### Mail Merge

After a successful split, click **Mail Merge** to send generated files by email.

Recipient mapping is loaded from an Excel worksheet with one row per split key:

- `Key`: matches the split key value
- `To`: required recipient address list
- `CC`: optional
- `BCC`: optional

Multiple email addresses in `To`, `CC`, and `BCC` use semicolon separators.

Mail Merge supports in-app subject/body placeholders such as `{key}`, `{to}`, and columns from the recipient mapping worksheet. An optional `.html` file can be used as the email body template.

Before sending, the app shows a carousel preview so each email can be checked one by one. Strict validation blocks sending if recipients, attachments, subject, body, or Outlook availability are invalid.

The first sending provider is Microsoft Outlook desktop. Delay delivery sets Outlook's deferred delivery time, and throttle controls how quickly the app hands messages to Outlook.

### PDF Export Options

#### xlwings (Microsoft Excel)
- **Pros**: Uses Excel's native PDF export and preserves workbook formatting
- **Cons**: Requires Microsoft Excel and COM automation access
- **Installation**: `pip install xlwings`

#### LibreOffice
- **Pros**: Better formatting preservation, supports complex layouts
- **Cons**: Requires LibreOffice installation
- **Installation**: Download from [libreoffice.org](https://www.libreoffice.org/)
- **Auto-detection**: Application searches common installation paths

## ⚙️ Configuration

### Automatic Settings
- The app saves paths, template option, sheet/key selections, PDF options, filename prefix/suffix, and column mapping automatically with Qt QSettings
- Settings are restored on startup
- Use "Reset Settings" to clear saved settings and return to defaults

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

#### "Microsoft Excel tidak dapat diakses via COM"
- Install Microsoft Excel and ensure it can run normally
- Install xlwings: `pip install xlwings`
- Or switch to "libreoffice" or "none" PDF engine

#### "LibreOffice not found"
- Install LibreOffice from [libreoffice.org](https://www.libreoffice.org/)
- Or manually browse to `soffice.exe` location
- Or switch to "xlwings" or "none" PDF engine

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
- `PySide6>=6.5.0` - Qt GUI framework
- `PySide6-Fluent-Widgets>=1.5.3` - Fluent-style widget toolkit
- `pandas>=2.3.2` - Data manipulation
- `openpyxl>=3.1.5` - Excel file handling
- `xlwings>=0.33.6` - Microsoft Excel PDF export (optional)

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

- Built with [PySide6](https://doc.qt.io/qtforpython-6/) and [PySide6-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets)
- Powered by [Pandas](https://pandas.pydata.org/) and [OpenPyXL](https://openpyxl.readthedocs.io/)
- PDF functionality via [xlwings](https://www.xlwings.org/) and LibreOffice

## 📞 Support

For support and questions:
- Create an issue on GitHub
- Check the troubleshooting section above
- Review the debug logs in the application

---

**Version**: 1.0.0
**Last Updated**: 2025
**Python**: 3.10+
