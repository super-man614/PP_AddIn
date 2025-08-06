# PowerPoint Add-in Tools

A comprehensive PowerPoint VSTO add-in that provides advanced tools for presentation creation and editing.

## Features

### 📋 File Operations
- New, Open, Save, Save As, Print presentations
- Smart file handling with automatic error recovery

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

## Installation

1. Build the project in Visual Studio
2. Install the generated VSTO file
3. Open PowerPoint - you'll see a "PowerPoint Add in Tools" tab in the ribbon

## Usage

- **Ribbon Access**: Use the "PowerPoint Add in Tools" tab for quick access to all features
- **Task Pane**: Click "Show Tools" to open the comprehensive task pane with all tools
- **Images**: The add-in automatically handles image loading with emoji fallbacks for maximum compatibility

## Technical Details

- **Framework**: .NET Framework 4.7.2
- **Office Version**: Compatible with PowerPoint 2016+
- **Deployment**: VSTO Click-Once deployment
- **Architecture**: Clean separation with services layer for maintainability

## Development

The add-in follows clean architecture principles:
- **Services**: Business logic separation
- **Models**: Data structures for shapes and slides
- **UI**: Task pane and ribbon integration
- **Error Handling**: Comprehensive error management with graceful degradation
