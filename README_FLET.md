# Excel Splitter - Modern Flet UI

🎨 **Modern Fluent Design implementation** of the Excel Splitter application using Google's Flet framework.

## ✨ What's New

### 🎯 Modern UI Features
- **True Fluent Design**: Adaptive Material 3 + Fluent Design System styling
- **Modern Components**: Cards, navigation rails, filled buttons, progress rings
- **Responsive Layout**: Adapts to different screen sizes and themes
- **Smooth Animations**: Professional transitions and interactions
- **Dark/Light Mode**: System-aware theme switching

### 🔥 Enhanced User Experience
- **Native File Pickers**: OS-integrated file selection dialogs  
- **Real-time Validation**: Visual feedback for form inputs
- **Progress Tracking**: Beautiful progress indicators and status logging
- **Modern Notifications**: Toast-style success/error messages
- **Professional Layout**: Clean, organized tabbed interface

## 🚀 Quick Start

### 1. **Install and Run**
```bash
# Option 1: Simple run script (recommended)
python run_flet.py

# Option 2: Direct execution  
pip install flet>=0.24.0
python main_flet.py
```

### 2. **Updated Requirements**
The new version uses:
```
flet>=0.24.0              # Modern UI framework
pandas>=2.3.2             # Data processing (unchanged)
openpyxl>=3.1.5          # Excel handling (unchanged)  
reportlab>=4.4.4         # PDF generation (unchanged)
```

## 📱 Modern UI Tour

### **Input Tab** - Source Configuration
- 📄 **Source Excel File**: Modern file picker with validation
- 📊 **Sheet Selection**: Dynamic dropdown with refresh capability  
- 🔑 **Key Column**: Smart header/index selection

### **Template Tab** - Formatting Setup
- 🎨 **Template File**: Excel template for output styling
- 🔢 **Header Rows**: Configurable header row count
- 💡 **Info Cards**: Helpful usage tips and guidance

### **Output Tab** - Export Configuration  
- 📁 **Output Folder**: Native folder selection dialog
- 📄 **PDF Engine**: Choose between ReportLab or LibreOffice
- 🏷️ **File Naming**: Customizable prefix and suffix options

### **Generate Tab** - Processing & Results
- 📋 **Configuration Summary**: Review all settings at a glance
- ▶️ **Generate Button**: Start processing with modern styling
- 📊 **Progress Tracking**: Real-time progress ring and status log
- 📝 **Status Logging**: Scrollable, formatted activity feed

## 🆚 Comparison: Before vs After

| Feature | CustomTkinter (Old) | Flet (New) |
|---------|-------------------|------------|
| **Design Language** | Basic theming | True Fluent Design |
| **File Dialogs** | System default | Native OS integration |
| **Progress Feedback** | Basic progress bar | Modern progress ring + logging |
| **Validation** | Text-based errors | Visual feedback + colors |
| **Theme Support** | Manual switching | System-aware adaptive |
| **Responsiveness** | Fixed layout | Responsive + adaptive |
| **Animations** | None | Smooth transitions |
| **Cross-platform** | Desktop only | Desktop + Web ready |

## 🎯 Key Improvements

### **Visual Enhancements**
- ✨ Modern Material 3 + Fluent Design components
- 🎨 Professional color schemes and typography  
- 📱 Responsive layouts that adapt to screen size
- 🔄 Smooth animations and transitions
- 🌓 Automatic light/dark theme switching

### **Usability Improvements**  
- 🎯 Intuitive navigation with modern tab design
- 📁 Native OS file/folder selection dialogs
- ✅ Real-time input validation with visual feedback
- 📊 Enhanced progress tracking with detailed logging
- 💡 Contextual help and information cards

### **Technical Improvements**
- 🚀 Modern Python patterns and async support
- 🔧 Better error handling and user feedback
- 📦 Simplified dependency management
- 🌐 Future-ready for web deployment
- ⚡ Improved performance and responsiveness

## 🔧 Advanced Usage

### **Running as Web App**
```bash
# Flet can also run in web browser
python main_flet.py --web
```

### **Custom Themes**
The app automatically adapts to your system theme, but you can customize colors by modifying the theme configuration in `setup_page()`.

### **Development Mode**
```bash  
# Run with hot reload for development
flet run main_flet.py
```

## 🐛 Troubleshooting

### **Common Issues**

**"Module not found: flet"**
```bash
pip install flet>=0.24.0
```

**File dialogs not working**
- Ensure you're running on a supported platform
- Try using the run script: `python run_flet.py`

**Theme not applying**
- Restart the application
- Check system theme settings

### **Performance Tips**
- Close other applications for large Excel files
- Use SSD storage for better file I/O performance
- Ensure sufficient RAM for data processing

## 🔄 Migration Guide

### **From CustomTkinter Version**
1. **Backup** your current configuration files
2. **Install** Flet: `pip install flet>=0.24.0`  
3. **Run** the new version: `python main_flet.py`
4. **Import** your existing configuration (coming soon)

### **Side-by-Side Usage**
Both versions can coexist:
- **Original**: `python main.py` 
- **Modern**: `python main_flet.py`

## 📊 Feature Parity

✅ **All original features preserved**:
- Template-based Excel splitting
- PDF export (ReportLab + LibreOffice)
- Configuration save/load (coming soon)
- Multi-sheet support
- Custom file naming
- Progress tracking
- Error handling

➕ **Plus new modern features**:
- Fluent Design styling
- Native file dialogs  
- Visual validation
- Responsive layout
- Theme adaptation
- Enhanced logging

## 🎉 Benefits Summary

### **For End Users**
- 🎨 **Beautiful Interface**: Modern, professional appearance
- ⚡ **Faster Workflow**: Intuitive navigation and feedback
- 🔄 **Better Experience**: Smooth interactions and responsiveness
- 🛠️ **Less Friction**: Native dialogs and validation

### **For Developers**  
- 📚 **Modern Codebase**: Clean, maintainable Python code
- 🔧 **Better Architecture**: Component-based UI design
- 🌐 **Future-Proof**: Web deployment ready
- 🚀 **Active Framework**: Flet is actively developed by Google

---

**Ready to experience modern UI design?** Run `python run_flet.py` to get started! 🚀