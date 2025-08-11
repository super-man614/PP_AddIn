using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using my_addin.Core;

namespace my_addin
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private static CustomTaskPane _taskPaneInstance;
        private static ColorPaletteTaskPane _colorPaletteInstance;

        public static CustomTaskPane TaskPaneInstance
        {
            get { return _taskPaneInstance; }
            set { _taskPaneInstance = value; }
        }

        public static ColorPaletteTaskPane ColorPaletteInstance
        {
            get { return _colorPaletteInstance; }
            set { _colorPaletteInstance = value; }
        }

        public Ribbon()
        {
            System.Diagnostics.Debug.WriteLine("Ribbon constructor called - Ribbon instance created");
        }

        public string GetCustomUI(string ribbonID)
        {
            System.Diagnostics.Debug.WriteLine("GetCustomUI called - returning ribbon XML with presets");
            
            // Return ribbon XML with presets
            return @"<?xml version='1.0' encoding='UTF-8'?>
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
  <ribbon>
    <tabs>
      <tab id='PowerPointToolsTab' label='PowerPoint Tools'>
        <group id='TestGroup' label='Test'>
          <button id='TestRibbonButton' 
                  label='Test Ribbon' 
                  size='large'
                  onAction='TestRibbon_Click'
                  imageMso='HappyFace'
                  screentip='Test if ribbon is working' />
        </group>
        <group id='TaskPaneGroup' label='Task Pane'>
          <button id='ToggleTaskPaneButton' 
                  label='Show Tools' 
                  size='large'
                  onAction='ToggleTaskPane_Click'
                  imageMso='TaskPane'
                  screentip='Toggle the PowerPoint Tools task pane' />
        </group>
        <group id='PresetsGroup' label='Presets'>
          <splitButton id='Preset1Split'>
            <button id='Preset1Apply' label='Preset 1' imageMso='StyleGallery' onAction='Preset1_Apply'/>
            <menu id='Preset1Menu'>
              <button id='Preset1Save' label='Save from Selection' imageMso='Save' onAction='Preset1_Save'/>
              <button id='Preset1Clear' label='Clear' imageMso='Cancel' onAction='Preset1_Clear'/>
            </menu>
          </splitButton>
          <splitButton id='Preset2Split'>
            <button id='Preset2Apply' label='Preset 2' imageMso='StyleGallery' onAction='Preset2_Apply'/>
            <menu id='Preset2Menu'>
              <button id='Preset2Save' label='Save from Selection' imageMso='Save' onAction='Preset2_Save'/>
              <button id='Preset2Clear' label='Clear' imageMso='Cancel' onAction='Preset2_Clear'/>
            </menu>
          </splitButton>
          <splitButton id='Preset3Split'>
            <button id='Preset3Apply' label='Preset 3' imageMso='StyleGallery' onAction='Preset3_Apply'/>
            <menu id='Preset3Menu'>
              <button id='Preset3Save' label='Save from Selection' imageMso='Save' onAction='Preset3_Save'/>
              <button id='Preset3Clear' label='Clear' imageMso='Cancel' onAction='Preset3_Clear'/>
            </menu>
          </splitButton>
        </group>
        <group id='FormatToolsGroup' label='Format Tools'>
          <splitButton id='UniformSizesSplit'>
            <button id='UniformSizesButton' label='Uniform Sizes' imageMso='SizeToFit' onAction='UniformSizes_Click'/>
            <menu id='UniformSizesMenu'>
              <button id='MatchWidthButton' label='Width' imageMso='Width' onAction='MatchWidth_Click'/>
              <button id='MatchHeightButton' label='Height' imageMso='Height' onAction='MatchHeight_Click'/>
              <button id='MatchSizeButton' label='Size' imageMso='SizeToFit' onAction='MatchSize_Click'/>
            </menu>
          </splitButton>
          <splitButton id='MatchColorsSplit'>
            <button id='MatchColorsButton' label='Match Colors' imageMso='FontColorPicker' onAction='MatchColors_Click'/>
            <menu id='MatchColorsMenu'>
              <button id='MatchFillButton' label='Fill' imageMso='FillColor' onAction='MatchFill_Click'/>
              <button id='MatchFontButton' label='Font' imageMso='FontColor' onAction='MatchFont_Click'/>
              <button id='MatchOutlineButton' label='Outline' imageMso='LineColor' onAction='MatchOutline_Click'/>
            </menu>
          </splitButton>
          <splitButton id='AlignFontsSplit'>
            <button id='AlignFontsButton' label='Align Fonts' imageMso='FontDialog' onAction='AlignFonts_Click'/>
            <menu id='AlignFontsMenu'>
              <button id='MatchFontSizeButton' label='Size' imageMso='FontSize' onAction='MatchFontSize_Click'/>
              <button id='MatchFontFamilyButton' label='Family' imageMso='FontFamily' onAction='MatchFontFamily_Click'/>
              <button id='MatchFontColorButton' label='Color' imageMso='FontColor' onAction='MatchFontColor_Click'/>
            </menu>
          </splitButton>
        </group>
        <group id='ColorGroup' label='Colors'>
          <toggleButton id='ColorPaletteToggleButton' 
                        label='Color Palette' 
                        size='large'
                        onAction='ColorPaletteToggle_Click'
                        getPressed='GetColorPalettePressed'
                        imageMso='FontColorPicker'
                        screentip='Toggle Color Palette task pane' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            System.Diagnostics.Debug.WriteLine("Ribbon_Load called - ribbon is being initialized");
            ribbon = ribbonUI;
            System.Diagnostics.Debug.WriteLine("Ribbon UI stored successfully");
        }

        public void ToggleTaskPane_Click(Office.IRibbonControl control)
        {
            try
            {
                if (_taskPaneInstance == null || _taskPaneInstance.IsDisposed)
                {
                    System.Diagnostics.Debug.WriteLine("Creating new task pane instance...");
                    _taskPaneInstance = new CustomTaskPane();
                }

                // Ensure our main Tools pane docks to the right
                try
                {
                    _taskPaneInstance.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                }
                catch { }

                // Make sure Color Palette exists and is also on the right
                if (_colorPaletteInstance == null || _colorPaletteInstance.IsDisposed)
                {
                    System.Diagnostics.Debug.WriteLine("Ensuring Color Palette task pane exists...");
                    _colorPaletteInstance = new ColorPaletteTaskPane();
                }
                try
                {
                    _colorPaletteInstance.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                }
                catch { }

                // Show both panes
                _taskPaneInstance.Visible = true;
                _colorPaletteInstance.Visible = true;

                // Ensure Color Palette sits left-most among right-docked panes
                try { Core.PaneOrdering.EnsureColorPaletteLeftMost(_colorPaletteInstance); } catch { }
                try { Core.PaneManager.OnPaneVisibilityChanged(); } catch { }

                // Update the toggle button state
                if (ribbon != null)
                    ribbon.Invalidate();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ToggleTaskPane_Click: {ex.Message}");
            }
        }

        public void TestRibbon_Click(Office.IRibbonControl control)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("TestRibbon_Click executed successfully");
                // Removed unnecessary popup message
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in TestRibbon_Click: {ex.Message}");
            }
        }

        public void ColorPaletteToggle_Click(Office.IRibbonControl control)
        {
            try
            {
                if (_colorPaletteInstance == null || _colorPaletteInstance.IsDisposed)
                {
                    System.Diagnostics.Debug.WriteLine("Creating new Color Palette task pane...");
                    _colorPaletteInstance = new ColorPaletteTaskPane();
                }

                // Force docking to the right
                try
                {
                    _colorPaletteInstance.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                }
                catch { }

                // If our Tools pane exists, ensure it's on the right as well
                if (_taskPaneInstance != null && !_taskPaneInstance.IsDisposed)
                {
                    try { _taskPaneInstance.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight; } catch { }
                }

                // Toggle visibility
                _colorPaletteInstance.Toggle();

                // If both are visible, ensure Color Palette sits left-most
                if (_colorPaletteInstance.Visible && _taskPaneInstance != null && !_taskPaneInstance.IsDisposed && _taskPaneInstance.Visible)
                {
                    try { Core.PaneOrdering.EnsureColorPaletteLeftMost(_colorPaletteInstance); } catch { }
                }
                try { Core.PaneManager.OnPaneVisibilityChanged(); } catch { }

                // Update the toggle button state
                if (ribbon != null)
                    ribbon.Invalidate();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ColorPaletteToggle_Click: {ex.Message}");
            }
        }

        public bool GetColorPalettePressed(Office.IRibbonControl control)
        {
            return _colorPaletteInstance != null && !_colorPaletteInstance.IsDisposed && _colorPaletteInstance.Visible;
        }

        // Preset callbacks
        public void Preset1_Apply(Office.IRibbonControl control) { ShapePresetStorage.ApplyPreset(1); }
        public void Preset1_Save(Office.IRibbonControl control) { ShapePresetStorage.SavePreset(1); }
        public void Preset1_Clear(Office.IRibbonControl control) { ShapePresetStorage.ClearPreset(1); }
        public void Preset2_Apply(Office.IRibbonControl control) { ShapePresetStorage.ApplyPreset(2); }
        public void Preset2_Save(Office.IRibbonControl control) { ShapePresetStorage.SavePreset(2); }
        public void Preset2_Clear(Office.IRibbonControl control) { ShapePresetStorage.ClearPreset(2); }
        public void Preset3_Apply(Office.IRibbonControl control) { ShapePresetStorage.ApplyPreset(3); }
        public void Preset3_Save(Office.IRibbonControl control) { ShapePresetStorage.SavePreset(3); }
        public void Preset3_Clear(Office.IRibbonControl control) { ShapePresetStorage.ClearPreset(3); }

        // Format Tools callbacks
        public void UniformSizes_Click(Office.IRibbonControl control) { FormatTools.MatchSize(); }
        public void MatchWidth_Click(Office.IRibbonControl control) { FormatTools.MatchWidth(); }
        public void MatchHeight_Click(Office.IRibbonControl control) { FormatTools.MatchHeight(); }
        public void MatchSize_Click(Office.IRibbonControl control) { FormatTools.MatchSize(); }
        
        public void MatchColors_Click(Office.IRibbonControl control) { FormatTools.MatchFill(); }
        public void MatchFill_Click(Office.IRibbonControl control) { FormatTools.MatchFill(); }
        public void MatchFont_Click(Office.IRibbonControl control) { FormatTools.MatchFontColor(); }
        public void MatchOutline_Click(Office.IRibbonControl control) { FormatTools.MatchOutline(); }
        
        public void AlignFonts_Click(Office.IRibbonControl control) { FormatTools.MatchFontSize(); }
        public void MatchFontSize_Click(Office.IRibbonControl control) { FormatTools.MatchFontSize(); }
        public void MatchFontFamily_Click(Office.IRibbonControl control) { FormatTools.MatchFontFamily(); }
        public void MatchFontColor_Click(Office.IRibbonControl control) { FormatTools.MatchFontColor(); }

        // Wrap toggle
        public void WrapToggle_Click(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app?.ActiveWindow?.Selection;
                if (selection?.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange == null || selection.ShapeRange.Count < 1)
                    return;

                // Determine current wrap setting from first shape with text
                Office.MsoTriState? current = null;
                for (int i = 1; i <= selection.ShapeRange.Count; i++)
                {
                    var sh = selection.ShapeRange[i];
                    if (sh.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        current = sh.TextFrame2.WordWrap;
                        break;
                    }
                }
                var newVal = (current == Office.MsoTriState.msoTrue) ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;

                for (int i = 1; i <= selection.ShapeRange.Count; i++)
                {
                    var sh = selection.ShapeRange[i];
                    if (sh.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        sh.TextFrame2.WordWrap = newVal;
                    }
                }
            }
            catch { }
        }
    }
}