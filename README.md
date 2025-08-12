# PowerPoint Add-in Tools

A comprehensive PowerPoint VSTO add-in that provides advanced tools for presentation creation and editing.

## ✨ **What's New in Version 1.0.0.3**

### 🔧 **Enhanced Error Handling & Logging**
- **Comprehensive Logging System**: File-based logging with automatic cleanup
- **Smart Error Recovery**: Better error handling with user-friendly messages
- **Debug Information**: Detailed logging for troubleshooting and development
- **Log Management**: Automatic log rotation and cleanup (30-day retention)

### ⚙️ **Configuration Management**
- **JSON Configuration**: Flexible configuration via `appsettings.json`
- **User Settings**: Personalized settings stored in user's AppData folder
- **Runtime Configuration**: Hot-reloadable configuration without restart
- **Default Presets**: Sensible defaults with easy customization

### 🚀 **Performance Improvements**
- **COM Object Management**: Proper disposal of Office COM objects
- **Memory Optimization**: Better resource management and cleanup
- **Async Operations**: Support for background operations
- **Operation Timeouts**: Configurable timeouts for long-running operations

### 🎯 **Enhanced Matrix Operations**
- **Improved Paste Functionality**: Better clipboard handling and error recovery
- **Cell Validation**: Robust matrix cell detection and validation
- **Progress Tracking**: Success/failure reporting for matrix operations
- **COM Safety**: Safe COM object handling with proper cleanup

## 🚀 **Features**

### 📋 File Operations
- New, Open, Save, Save As, Print presentations
- Smart file handling with automatic error recovery
- Support for multiple document formats (PPTX, PPT, PDF, HTML)

### 🎨 Wizards
- **Agenda Wizard**: Create structured agenda slides
- **Master Access**: Quick access to slide master editing
- **Element Wizard**: Insert smart design elements
- **Text Wizard**: Advanced text formatting tools
- **Format Wizard**: Apply consistent formatting across presentations
- **Map Wizard**: Insert various map templates and diagrams

### 📊 Smart Elements
- **Charts & Diagrams**: Insert and format charts
- **Tables**: Create regular and matrix tables with professional styling
- **Sticky Notes**: Add annotatable sticky notes
- **Citations**: Insert properly formatted citations
- **Standard Objects**: Quick access to common shapes and objects

### 🎯 Position & Alignment
- Comprehensive alignment tools (left, center, right, top, bottom, middle)
- Distribution tools for even spacing
- Dimension matching (width, height, both)
- Advanced positioning features

### 🔧 Shape Tools
- Process chain alignment
- Angle alignment for block arrows
- Pentagon header adjustment
- Rounded rectangle radius control

### 🎨 Formatting
- Fill, text, and outline color tools
- Text formatting (bold, italic, underline, bullets)
- Text wrapping controls

### 🔍 Navigation
- Zoom controls
- Fit to window functionality

## ⚙️ **Configuration**

The add-in now supports extensive configuration through `appsettings.json`:

```json
{
  "logging": {
    "enableFileLogging": true,
    "enableDebugLogging": true,
    "logRetentionDays": 30,
    "maxLogFileSizeMB": 10
  },
  "ui": {
    "defaultTaskPaneWidth": 300,
    "defaultTaskPaneHeight": 600,
    "autoShowTaskPanes": true,
    "rememberTaskPanePositions": true
  },
  "matrix": {
    "defaultText": "XXXX",
    "maxRows": 100,
    "maxColumns": 100,
    "autoResizeCells": true
  },
  "colorPalette": {
    "maxColors": 50,
    "defaultColorSize": 24,
    "autoSave": true
  }
}
```

## 📁 **Installation**

1. **Build the project** in Visual Studio 2019 or later
2. **Install the generated VSTO file** (`my-addin.vsto`)
3. **Open PowerPoint** - you'll see a "PowerPoint Add in Tools" tab in the ribbon

### **System Requirements**
- Windows 10/11
- Microsoft Office 2016 or later
- .NET Framework 4.7.2 or later
- PowerPoint application

## 🔧 **Development & Customization**

### **Project Structure**
```
PP_AddIn/
├── Core/                    # Core business logic
├── Services/               # Service layer (error handling, etc.)
├── Constants/              # Application constants and configuration
├── Models/                 # Data models
├── icons/                  # Icon resources
└── bin/Debug/             # Build output
```

### **Key Components**
- **ErrorHandlerService**: Centralized error handling and logging
- **ConfigurationManager**: Application configuration management
- **PaneManager**: Task pane lifecycle management
- **ServiceContainer**: Dependency injection container

### **Building from Source**
1. Clone the repository
2. Open `my-addin.sln` in Visual Studio
3. Restore NuGet packages
4. Build the solution
5. The VSTO file will be generated in `bin/Debug/`

## 📝 **Usage**

- **Ribbon Access**: Use the "PowerPoint Add in Tools" tab for quick access to all features
- **Task Pane**: Click "Show Tools" to open the comprehensive task pane with all tools
- **Keyboard Shortcuts**: Ctrl+V for matrix paste operations
- **Context Menus**: Right-click on shapes for additional options

## 🐛 **Troubleshooting**

### **Log Files**
Logs are stored in: `%APPDATA%\PowerPointAddIn\Logs\`
- Check the latest log file for error details
- Logs are automatically rotated and cleaned up

### **Common Issues**
1. **Add-in not loading**: Check Office version compatibility
2. **Task panes not showing**: Verify Office trust center settings
3. **Matrix operations failing**: Ensure shapes have proper matrix tags

### **Error Reporting**
- All errors are logged with detailed information
- User-friendly error messages are displayed
- Debug information is available in log files

## 🤝 **Contributing**

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📄 **License**

This project is licensed under the MIT License - see the LICENSE file for details.

## 🔗 **Support**

- **Issues**: Report bugs and feature requests via GitHub Issues
- **Documentation**: Check the IMPLEMENTATION_GUIDE.md for detailed technical information
- **Configuration**: Modify `appsettings.json` for customization

---

**Version**: 1.0.0.3  
**Last Updated**: December 2024  
**Compatibility**: PowerPoint 2016+ / Office 365
