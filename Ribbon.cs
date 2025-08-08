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
            System.Diagnostics.Debug.WriteLine("GetCustomUI called - returning simplified ribbon XML");
            
            // Return simplified ribbon XML directly
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

        // Additional methods can be added here as needed
    }
}