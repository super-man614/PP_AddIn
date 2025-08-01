# PowerPoint Add-In Implementation Guide

This guide documents the comprehensive functionality implemented in the PowerPoint Add-In.

## Position Section - Comprehensive Alignment & Positioning Tools

The Position Section provides advanced object positioning capabilities that go far beyond PowerPoint's standard alignment tools. All functions support the "Master Object" concept where the last selected object serves as the reference point.

### 🎯 Basic Alignment Functions

#### **Standard Alignment**
- **Align Left** (←): Align objects to left edge
- **Align Center** (↔): Align objects to horizontal center  
- **Align Right** (→): Align objects to right edge
- **Align Top** (↑): Align objects to top edge
- **Align Bottom** (↓): Align objects to bottom edge
- **Align Middle** (↕): Align objects to vertical middle

**Master Object Logic:**
- Multiple objects: Align to the **last selected object** (master)
- Single object OR **Ctrl pressed**: Align to slide edges
- Hold **Ctrl** to force alignment to slide boundaries

### 🔗 Docking Functions

Advanced positioning that moves objects to "touch" edges rather than just align them.

- **Dock Left**: Move objects to touch the left side of master object (or slide edge with Ctrl)
- **Dock Right**: Move objects to touch the right side of master object (or slide edge with Ctrl) 
- **Dock Top**: Move objects to touch the top of master object (or slide top with Ctrl)
- **Dock Bottom**: Move objects to touch the bottom of master object (or slide bottom with Ctrl)

**Use Cases:**
- Creating connected layouts where objects touch each other
- Positioning objects at precise slide boundaries
- Building flowcharts and process diagrams

### 📐 Enhanced Distribution

- **Distribute** (≡): General distribution
- **Distribute Horizontal** (⇔): Horizontal distribution with Ctrl support
- **Distribute Vertical** (⇕): Vertical distribution with Ctrl support

**Distribution Modes:**
- **Standard**: Keep outer objects in place, distribute middle objects evenly
- **Ctrl Mode**: Distribute across entire slide width/height

### 📏 Size Matching

- **Match Both**: Match both width and height to master object
- **Match Height**: Match height only to master object
- **Match Width**: Match width only to master object

### 🔄 Transform Operations

- **Make Vertical**: Rotate objects to 90° (vertical orientation)
- **Make Horizontal**: Rotate objects to 0° (horizontal orientation)  
- **Swap Locations**: Exchange positions of exactly two selected objects

### ✨ Advanced Positioning Functions

#### **Golden Canon Alignment**
Implements the Golden Canon ratio for professional layout design.
- Creates 1:2 margin ratio (bottom margin twice the top margin)
- Master object should be taller than objects being positioned
- Perfect for typographic and design layouts

#### **Matrix Alignment**
Arrange objects in a precise grid pattern.
- Specify rows × columns (e.g., "3x2" for 3 rows, 2 columns)
- Objects filled row-wise from top to bottom
- Maintains selection order when placing objects
- Automatically calculates grid boundaries from selected objects

#### **Slice or Multiply Shape**
Transform a single shape into a grid of smaller shapes.
- Format: "rows x columns" (e.g., "2x3" creates 6 shapes)
- Optional spacing: "2x3 10" adds 10pt spacing between shapes
- Original shape becomes top-left piece
- Perfect for creating grids, tiles, or modular layouts

#### **Duplicate Right**
Quick duplication with automatic positioning.
- Duplicates all selected objects to the right
- Adds 10pt spacing automatically
- Maintains vertical alignment

#### **Center on Top Left Corner**
Precision positioning using master object's corner as reference.
- Centers objects on the top-left corner of master object
- Useful for creating layered designs or alignment markers

### 💾 Position Memory System

#### **Save Position and Size**
- Captures exact position (Left, Top) and dimensions (Width, Height) of selected objects
- Stores multiple objects in selection order
- Data persists during the PowerPoint session

#### **Apply Position and Size**  
- Applies saved positions and sizes to currently selected objects
- Matches objects in selection order
- Perfect for:
  - Creating object templates
  - Maintaining consistent layouts across slides
  - Copying precise positioning between different objects

### 🎮 Usage Tips

1. **Master Object**: Always select the reference object **last** when working with multiple objects
2. **Ctrl Key**: Use Ctrl to force operations relative to slide boundaries instead of master object
3. **Selection Order**: For matrix alignment and position memory, selection order matters
4. **Error Handling**: All functions include comprehensive error checking and user feedback

### 🔧 Technical Implementation

**Key Features:**
- Full support for PowerPoint's native alignment commands where applicable
- Custom algorithms for advanced positioning (Golden Canon, Matrix, Slicing)
- Robust error handling with informative user messages
- Memory system for position/size templates
- Ctrl key modifier support for alternate behaviors
- Tooltips with detailed usage instructions

**Compatibility:**
- Works with all PowerPoint shape types
- Supports single and multiple object selections
- Handles edge cases (single objects, slide boundaries, etc.)
- Maintains PowerPoint's undo/redo functionality

This comprehensive position system transforms PowerPoint into a precision design tool capable of professional-grade object positioning and layout management. 