using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace my_addin
{
    public partial class TaskPaneControl : UserControl
    {
        public TaskPaneControl()
        {
            InitializeComponent();
            OptimizeRemainingPanels(); // Optimize panels not yet optimized in Designer
            SetupEventHandlers();
            SetInitialValues();
            LoadButtonImages(); // Load custom images for buttons
            SetupTooltips(); // Setup tooltips for all buttons
        }

        /// <summary>
        /// Optimizes remaining button panels that still use individual Controls.Add calls
        /// This method handles panels not yet optimized in the Designer file
        /// </summary>
        private void OptimizeRemainingPanels()
        {
            // Wizard buttons panel optimization
            if (wizardButtonsPanel != null)
            {
                wizardButtonsPanel.SuspendLayout();
                var wizardControls = wizardButtonsPanel.Controls.Cast<Control>().ToArray();
                wizardButtonsPanel.Controls.Clear();
                wizardButtonsPanel.Controls.AddRange(wizardControls);
                wizardButtonsPanel.ResumeLayout(true);
            }
            
            // Text buttons panel optimization  
            if (textButtonsPanel != null)
            {
                textButtonsPanel.SuspendLayout();
                var textControls = textButtonsPanel.Controls.Cast<Control>().ToArray();
                textButtonsPanel.Controls.Clear();
                textButtonsPanel.Controls.AddRange(textControls);
                textButtonsPanel.ResumeLayout(true);
            }
            
            // Shape buttons panel optimization
            if (shapeButtonsPanel != null)
            {
                shapeButtonsPanel.SuspendLayout();
                var shapeControls = shapeButtonsPanel.Controls.Cast<Control>().ToArray();
                shapeButtonsPanel.Controls.Clear();
                shapeButtonsPanel.Controls.AddRange(shapeControls);
                shapeButtonsPanel.ResumeLayout(true);
            }
            
            // Color buttons panel optimization
            if (colorButtonsPanel != null)
            {
                colorButtonsPanel.SuspendLayout();
                var colorControls = colorButtonsPanel.Controls.Cast<Control>().ToArray();
                colorButtonsPanel.Controls.Clear();
                colorButtonsPanel.Controls.AddRange(colorControls);
                colorButtonsPanel.ResumeLayout(true);
            }
            
            // Navigation buttons panel optimization
            if (navigationButtonsPanel != null)
            {
                navigationButtonsPanel.SuspendLayout();
                var navControls = navigationButtonsPanel.Controls.Cast<Control>().ToArray();
                navigationButtonsPanel.Controls.Clear();
                navigationButtonsPanel.Controls.AddRange(navControls);
                navigationButtonsPanel.ResumeLayout(true);
            }
        }

        private void SetupEventHandlers()
        {
            // Size section events
            if (cmbSlideSize != null)
                cmbSlideSize.SelectedIndexChanged += CmbSlideSize_SelectedIndexChanged;
            if (btnApplySize != null)
                btnApplySize.Click += BtnApplySize_Click;
            
            // OPTIMIZED: Apply hover effects to ALL buttons at once using our utility
            // This replaces ~250 lines of duplicate hover event handlers!
            ButtonHoverUtility.EnableHoverEffectsForContainer(this, recursive: true);
            
            // Load current slides
            this.Load += TaskPaneControl_Load;
        }

        private void SetInitialValues()
        {
            // Set default combo box selection
            if (cmbSlideSize != null && cmbSlideSize.Items.Count > 1)
                cmbSlideSize.SelectedIndex = 1; // Default to 16:9
        }

        /// <summary>
        /// Loads custom images for buttons from the icons directory
        /// </summary>
        private void LoadButtonImages()
        {
            try
            {
                string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string assemblyDir = System.IO.Path.GetDirectoryName(assemblyPath);
                
                // Load image for btnNew
                string newIconPath = System.IO.Path.Combine(assemblyDir, "icons", "icons8-open-file-48.png");
                if (System.IO.File.Exists(newIconPath))
                {
                    btnNew.BackgroundImage = Image.FromFile(newIconPath);
                    btnNew.BackgroundImageLayout = ImageLayout.Stretch;
                    btnNew.Text = ""; // Clear text to show image
                }

                // Load wizard button images
                LoadWizardButtonImages(assemblyDir);
            }
            catch (Exception ex)
            {
                // Silently fail - buttons will use emoji fallbacks
                System.Diagnostics.Debug.WriteLine($"Error loading button images: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads images for wizard buttons
        /// </summary>
        private void LoadWizardButtonImages(string assemblyDir)
        {
            try
            {
                // Wizard button images
                var wizardButtons = new Dictionary<Button, string>
                {
                    { btnAgenda, "icons/wizzards/agenda.png" },
                    { btnMaster, "icons/wizzards/master.png" },
                    { btnElement, "icons/wizzards/element.png" },
                    { btnText, "icons/wizzards/text.png" },
                    { btnFormat, "icons/wizzards/format.png" },
                    { btnMap, "icons/wizzards/map.png" }
                };

                foreach (var button in wizardButtons)
                {
                    string iconPath = System.IO.Path.Combine(assemblyDir, button.Value);
                    if (System.IO.File.Exists(iconPath))
                    {
                        button.Key.BackgroundImage = Image.FromFile(iconPath);
                        button.Key.BackgroundImageLayout = ImageLayout.Stretch;
                        button.Key.Text = ""; // Clear text to show image
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading wizard button images: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets up tooltips for all buttons
        /// </summary>
        private void SetupTooltips()
        {
            // Create tooltip control
            var tooltip = new ToolTip();
            tooltip.AutoPopDelay = 5000;
            tooltip.InitialDelay = 1000;
            tooltip.ReshowDelay = 500;
            tooltip.ShowAlways = true;

            // Presentation section tooltips
            tooltip.SetToolTip(btnNew, "New Presentation");
            tooltip.SetToolTip(btnOpen, "Open Presentation");
            tooltip.SetToolTip(btnSave, "Save Presentation");
            tooltip.SetToolTip(btnSaveAs, "Save As");
            tooltip.SetToolTip(btnPrint, "Print");
            //tooltip.SetToolTip(btnShare, "Share");

            // Wizards section tooltips
            tooltip.SetToolTip(btnAgenda, "Agenda");
            tooltip.SetToolTip(btnMaster, "Master");
            tooltip.SetToolTip(btnElement, "Element");
            tooltip.SetToolTip(btnText, "Text");
            tooltip.SetToolTip(btnFormat, "Format");
            tooltip.SetToolTip(btnMap, "Map");

            // Smart Elements section tooltips
            tooltip.SetToolTip(btnChart, "Chart");
            tooltip.SetToolTip(btnDiagram, "Diagram");
            tooltip.SetToolTip(btnTable, "Table");
            tooltip.SetToolTip(btnMatrixTable, "Matrix Table");
            tooltip.SetToolTip(btnStickyNote, "Sticky Note");
            tooltip.SetToolTip(btnCitation, "Citation");
            tooltip.SetToolTip(btnStandardObjects, "Standard Objects");

            // Position section tooltips
            tooltip.SetToolTip(btnAlignLeft, "Align Left");
            tooltip.SetToolTip(btnAlignCenter, "Align Center");
            tooltip.SetToolTip(btnAlignRight, "Align Right");
            tooltip.SetToolTip(btnDistribute, "Distribute");
            tooltip.SetToolTip(btnMatchBoth, "Match Both");
            tooltip.SetToolTip(btnMatchHeight, "Match Height");
            tooltip.SetToolTip(btnMatchWidth, "Match Width");

            // Shape section tooltips
            tooltip.SetToolTip(btnAlignProcessChain, "Align Process Chain");
            tooltip.SetToolTip(btnAlignAngles, "Align Angles");
            tooltip.SetToolTip(btnAlignToProcessArrow, "Align to Process Arrow");
            tooltip.SetToolTip(btnAdjustPentagonHeader, "Adjust Pentagon Header");
            tooltip.SetToolTip(btnAlignBlockArrows, "Align Block Arrows");
            tooltip.SetToolTip(btnAlignRoundedRectangleArrows, "Align Rounded Rectangle Arrows");

            // Transform section tooltips
            tooltip.SetToolTip(btnMakeVertical, "Make Vertical");
            tooltip.SetToolTip(btnMakeHorizontal, "Make Horizontal");
            tooltip.SetToolTip(btnSwapLocations, "Swap Locations");

            // Size section tooltips
            tooltip.SetToolTip(btnApplySize, "Apply Size");

            // Colors section tooltips
            tooltip.SetToolTip(btnFillColor, "Fill Color");
            tooltip.SetToolTip(btnTextColor, "Text Color");
            tooltip.SetToolTip(btnOutlineColor, "Outline Color");

            // Text section tooltips
            tooltip.SetToolTip(btnBold, "Bold");
            tooltip.SetToolTip(btnItalic, "Italic");
            tooltip.SetToolTip(btnUnderline, "Underline");
            tooltip.SetToolTip(btnBullets, "Bullets");
            tooltip.SetToolTip(btnWrapText, "Wrap Text");
            tooltip.SetToolTip(btnNoWrapText, "No Wrap Text");

            // Navigation section tooltips
            tooltip.SetToolTip(btnZoomIn, "Zoom In");
            tooltip.SetToolTip(btnZoomOut, "Zoom Out");
            tooltip.SetToolTip(btnFitToWindow, "Fit to Window");

            // Expert Tools section tooltips
            tooltip.SetToolTip(btnFreeWebinar, "Free Webinar");
        }

        #region Event Handlers

        private void TaskPaneControl_Load(object sender, EventArgs e)
        {
            // Initialize the interface
        }

        #region Presentation Section

        private void BtnNew_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                app.Presentations.Add();
                MessageBox.Show("New presentation created!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating presentation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnNew_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnNew_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnOpen_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnOpen_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnSave_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnSave_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnSaveAs_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnSaveAs_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnPrint_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnPrint_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnShare_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnShare_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        // Wizards section button hover methods
        private void BtnAgenda_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnAgenda_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnMaster_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnMaster_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnElement_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnElement_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnText_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnText_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnFormat_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnFormat_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnMap_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnMap_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        // Smart Elements section button hover methods
        private void BtnChart_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnChart_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnDiagram_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnDiagram_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnTable_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnTable_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        // Matrix table section button hover methods
        private void BtnMatrixTable_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnMatrixTable_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        // Sticky notes section button hover methods
        private void BtnStickyNote_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }
        private void BtnStickyNote_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        // Citation and standard objects section button hover methods
        private void BtnCitation_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }
        private void BtnCitation_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }
        private void BtnStandardObjects_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }
        private void BtnStandardObjects_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        // Position section button hover methods
        private void BtnAlignLeft_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnAlignLeft_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnAlignCenter_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnAlignCenter_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnAlignRight_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnAlignRight_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnDistribute_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnDistribute_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        // Dimension matching section button hover methods
        private void BtnMatchBoth_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnMatchBoth_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnMatchHeight_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnMatchHeight_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnMatchWidth_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnMatchWidth_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        // Rotation and swap section button hover methods
        private void BtnMakeVertical_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }
        private void BtnMakeVertical_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }
        private void BtnMakeHorizontal_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }
        private void BtnMakeHorizontal_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }
        private void BtnSwapLocations_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }
        private void BtnSwapLocations_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        // Shape section button hover methods - Removed old handlers as they're now handled by ButtonHoverUtility

        // Color section button hover methods
        private void BtnFillColor_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnFillColor_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnTextColor_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnTextColor_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnOutlineColor_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnOutlineColor_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        // Text section button hover methods
        private void BtnBold_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnBold_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnItalic_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnItalic_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnUnderline_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnUnderline_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnBullets_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnBullets_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            try
            {
                var openDialog = new OpenFileDialog();
                openDialog.Filter = "PowerPoint Files (*.pptx;*.ppt)|*.pptx;*.ppt|All Files (*.*)|*.*";
                openDialog.Title = "Open Presentation";
                
                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    var app = Globals.ThisAddIn.Application;
                    app.Presentations.Open(openDialog.FileName);
                    MessageBox.Show("Presentation opened successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening presentation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    app.ActivePresentation.Save();
                    MessageBox.Show("Presentation saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("No active presentation to save.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving presentation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSaveAs_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    var saveDialog = new SaveFileDialog();
                    saveDialog.Filter = "PowerPoint Presentation (*.pptx)|*.pptx|PowerPoint 97-2003 (*.ppt)|*.ppt";
                    saveDialog.Title = "Save Presentation As";
                    
                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        app.ActivePresentation.SaveAs(saveDialog.FileName);
                        MessageBox.Show("Presentation saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("No active presentation to save.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving presentation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    app.ActivePresentation.PrintOut();
                    MessageBox.Show("Print dialog opened!", "Print", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("No active presentation to print.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error printing presentation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnShare_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("Share functionality would integrate with OneDrive/SharePoint.\nFor now, use File > Share in PowerPoint.", 
                              "Share", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing share options: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Wizards Section

        private void BtnAgenda_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    var slide = app.ActivePresentation.Slides.Add(app.ActivePresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
                    slide.Shapes.Title.TextFrame.TextRange.Text = "Agenda";
                    slide.Shapes.Placeholders[2].TextFrame.TextRange.Text = "• Introduction\n• Main Topics\n• Discussion\n• Next Steps";
                    MessageBox.Show("Agenda slide created!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating agenda: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnMaster_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    app.ActiveWindow.ViewType = PowerPoint.PpViewType.ppViewSlideMaster;
                    MessageBox.Show("Switched to Slide Master view!", "Master", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing slide master: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnElement_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("Smart Element wizard would help create interactive elements.\nTry inserting SmartArt from the Insert tab.", 
                              "Elements", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing elements: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnTextWizard_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    var slide = app.ActivePresentation.Slides.Add(app.ActivePresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
                    slide.Shapes.Title.TextFrame.TextRange.Text = "Text Content";
                    MessageBox.Show("Text slide created! You can now add your content.", "Text Wizard", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating text slide: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnFormat_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("Format wizard would help apply consistent formatting.\nUse the Design tab for themes and Format Painter for copying formats.", 
                              "Format", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing format options: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnMap_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("Map wizard would insert interactive maps.\nTry Insert > Online Pictures and search for 'map' or use Insert > Icons.", 
                              "Map", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing map options: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Smart Elements Section

        private void BtnChart_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null && app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    var slide = app.ActiveWindow.Selection.SlideRange[1];
                    slide.Shapes.AddChart2(Style: -1, Type: Office.XlChartType.xlColumnClustered);
                    MessageBox.Show("Chart added successfully!", "Chart", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select a slide first, then insert chart from Insert > Chart.", "Chart", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting chart: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDiagram_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("For diagrams, use Insert > SmartArt to create professional diagrams and flowcharts.", 
                              "Diagram", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing diagram options: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnTable_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null && app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    // Show table dropdown control
                    var tableDropdown = new TableDropdownControl();
                    
                    // Position the dropdown below the table button
                    var btnLocation = btnTable.PointToScreen(Point.Empty);
                    tableDropdown.Location = new Point(btnLocation.X, btnLocation.Y + btnTable.Height);
                    
                    if (tableDropdown.ShowDialog() == DialogResult.OK)
                    {
                        var slide = app.ActiveWindow.Selection.SlideRange[1];
                        
                        switch (tableDropdown.ActionType)
                        {
                            case "GridSelect":
                                CreateTableFromGrid(slide, tableDropdown.SelectedRows, tableDropdown.SelectedColumns);
                                break;
                            case "InsertTable":
                                ShowInsertTableDialog(slide);
                                break;
                            case "ExcelSpreadsheet":
                                InsertExcelSpreadsheet(slide);
                                break;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select a slide first to insert a table.", "Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting table: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateTableFromGrid(PowerPoint.Slide slide, int rows, int columns)
        {
            try
            {
                // Calculate optimal position and size
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;
                
                // Table dimensions
                float tableWidth = slideWidth * 0.8f;
                float tableHeight = Math.Min(slideHeight * 0.6f, rows * 40f);
                
                // Center the table
                float left = (slideWidth - tableWidth) / 2;
                float top = (slideHeight - tableHeight) / 2;
                
                // Create native PowerPoint table
                var tableShape = slide.Shapes.AddTable(rows, columns, left, top, tableWidth, tableHeight);
                var table = tableShape.Table;
                
                // Apply basic styling and clear any default header text
                ApplyBasicStyling(table, false); // No header by default for grid selection
                ClearTableHeaderText(table);
                
                MessageBox.Show($"Table ({rows}x{columns}) created successfully!", "Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create table from grid: {ex.Message}");
            }
        }

        private void ShowInsertTableDialog(PowerPoint.Slide slide)
        {
            try
            {
                                 var simpleDialog = new SimpleTableDialog();
                 if (simpleDialog.ShowDialog() == DialogResult.OK)
                 {
                     int rows = simpleDialog.Rows;
                     int columns = simpleDialog.Columns;
                     
                     // Calculate optimal position and size
                     float slideWidth = slide.Master.Width;
                     float slideHeight = slide.Master.Height;
                     
                     // Table dimensions
                     float tableWidth = slideWidth * 0.8f;
                     float tableHeight = Math.Min(slideHeight * 0.6f, rows * 40f);
                     
                     // Center the table
                     float left = (slideWidth - tableWidth) / 2;
                     float top = (slideHeight - tableHeight) / 2;
                     
                     // Create native PowerPoint table
                     var tableShape = slide.Shapes.AddTable(rows, columns, left, top, tableWidth, tableHeight);
                     var table = tableShape.Table;
                     
                     // Apply basic styling and clear any default header text
                     ApplyBasicStyling(table, false);
                     ClearTableHeaderText(table);
                     
                     MessageBox.Show($"Table ({rows}x{columns}) created successfully!", "Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error showing insert table dialog: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InsertExcelSpreadsheet(PowerPoint.Slide slide)
        {
            try
            {
                // Calculate position for Excel object
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;
                
                float objWidth = slideWidth * 0.7f;
                float objHeight = slideHeight * 0.5f;
                float left = (slideWidth - objWidth) / 2;
                float top = (slideHeight - objHeight) / 2;
                
                // Insert Excel spreadsheet as OLE object
                var excelObject = slide.Shapes.AddOLEObject(
                    left, top, objWidth, objHeight,
                    "Excel.Sheet", "", 
                    Office.MsoTriState.msoFalse, "", 0, "", 
                    Office.MsoTriState.msoFalse);
                
                excelObject.Name = "ExcelSpreadsheet";
                
                MessageBox.Show("Excel spreadsheet inserted successfully!\n\nDouble-click the object to edit in Excel.", "Excel Spreadsheet", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting Excel spreadsheet: {ex.Message}\n\nNote: Excel must be installed for this feature to work.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnMatrixTable_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null && app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    // Get user input for rows and columns
                    var matrixDialog = new MatrixTableDialog();
                    if (matrixDialog.ShowDialog() == DialogResult.OK)
                    {
                        int rows = matrixDialog.Rows;
                        int columns = matrixDialog.Columns;
                        bool hasHeader = matrixDialog.HasHeader;
                        
                        var slide = app.ActiveWindow.Selection.SlideRange[1];
                        CreateMatrixTable(slide, rows, columns, hasHeader);
                        
                        MessageBox.Show($"✅ Matrix table ({rows}x{columns}) created successfully!\n\n🎯 Matrix tables are perfect for:\n• Decision-making matrices\n• Feature comparisons\n• Structured analysis\n• SWOT analysis\n\n💡 All cells are filled with 'XXXX' placeholder text.\n✏️ Click on any cell to replace with your content.\n🎨 Clean uniform design with transparent cells and gray borders.", "Matrix Table Created", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select a slide first to insert matrix table.", "Matrix Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating matrix table: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateMatrixTable(PowerPoint.Slide slide, int rows, int columns, bool hasHeader)
        {
            try
            {
                // Calculate optimal position and size for matrix
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;
                
                // Matrix dimensions (make it more square and centered)
                float tableSize = Math.Min(slideWidth * 0.75f, slideHeight * 0.7f);
                float tableWidth = tableSize;
                float tableHeight = tableSize;
                
                // Center the matrix table
                float left = (slideWidth - tableWidth) / 2;
                float top = (slideHeight - tableHeight) / 2;
                
                // Create native PowerPoint table
                var tableShape = slide.Shapes.AddTable(rows, columns, left, top, tableWidth, tableHeight);
                var table = tableShape.Table;
                
                                        // Apply uniform styling (ignoring header setting for consistent appearance)
                        ApplyMatrixTableStyle(table, false);
                
                // Set default matrix content
                SetMatrixDefaultContent(table, hasHeader);
                
                // Focus on the first editable cell
                int startRow = hasHeader ? 2 : 1;
                int startCol = hasHeader ? 2 : 1;
                if (rows >= startRow && columns >= startCol)
                {
                    table.Cell(startRow, startCol).Shape.TextFrame.TextRange.Select();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create matrix table: {ex.Message}");
            }
        }

        private void CreateProfessionalTable(PowerPoint.Slide slide, int rows, int columns, bool hasHeader)
        {
            try
            {
                // Calculate optimal position and size
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;
                
                // Table dimensions (leave margins)
                float tableWidth = slideWidth * 0.8f; // 80% of slide width
                float tableHeight = Math.Min(slideHeight * 0.6f, rows * 40f); // Max 60% of height or based on rows
                
                // Center the table
                float left = (slideWidth - tableWidth) / 2;
                float top = (slideHeight - tableHeight) / 2;
                
                // Create native PowerPoint table
                var tableShape = slide.Shapes.AddTable(rows, columns, left, top, tableWidth, tableHeight);
                var table = tableShape.Table;
                
                // Apply PowerPoint's built-in table style
                ApplyDefaultTableStyle(table, hasHeader);
                
                // Set default content if it's a header table
                if (hasHeader && rows > 0)
                {
                    for (int col = 1; col <= columns; col++)
                    {
                        table.Cell(1, col).Shape.TextFrame.TextRange.Text = $"Header {col}";
                    }
                }
                
                // Focus on the first editable cell
                int startRow = hasHeader ? 2 : 1;
                if (rows >= startRow && startRow <= rows)
                {
                    table.Cell(startRow, 1).Shape.TextFrame.TextRange.Select();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create professional table: {ex.Message}");
            }
        }

        private void ApplyDefaultTableStyle(PowerPoint.Table table, bool hasHeader)
        {
            try
            {
                // Apply enhanced styling for regular tables
                ApplyEnhancedBasicStyling(table, hasHeader);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Table styling failed: {ex.Message}");
                // Fallback to basic styling
                ApplyBasicStyling(table, hasHeader);
            }
        }

        private void ApplyMatrixTableStyle(PowerPoint.Table table, bool hasHeader)
        {
            try
            {
                // Apply uniform styling to match the image - all cells identical
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        var cell = table.Cell(row, col);
                        var shape = cell.Shape;
                        
                        // Set uniform cell height for square appearance
                        table.Rows[row].Height = Math.Max(50f, table.Rows[row].Height);
                        
                        // No background color - transparent cells
                        shape.Fill.Visible = Office.MsoTriState.msoFalse;
                        
                        // Uniform text styling for all cells - dark gray text
                        shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(64, 64, 64));
                        shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoFalse;
                        shape.TextFrame.TextRange.Font.Size = 11;
                        
                        // Center alignment for all cells
                        shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                        shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                        
                        // Uniform borders for all cells - dark gray borders
                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                        shape.Line.Weight = 1.0f;
                        shape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(64, 64, 64));
                        shape.Line.DashStyle = Office.MsoLineDashStyle.msoLineSolid;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Matrix styling failed: {ex.Message}");
                // Fallback to uniform styling
                ApplyUniformMatrixStyling(table);
            }
        }

        private void SetMatrixDefaultContent(PowerPoint.Table table, bool hasHeader)
        {
            try
            {
                // Fill ALL cells with "XXXX" regardless of header setting
                // This matches the uniform XXXX pattern shown in the image
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        table.Cell(row, col).Shape.TextFrame.TextRange.Text = "XXXX";
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Setting matrix content failed: {ex.Message}");
            }
        }

        private void ApplyUniformMatrixStyling(PowerPoint.Table table)
        {
            try
            {
                // Simple uniform styling that matches the image exactly
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        var cell = table.Cell(row, col);
                        var shape = cell.Shape;
                        
                        // No background color - transparent cells
                        shape.Fill.Visible = Office.MsoTriState.msoFalse;
                        
                        // Dark gray text
                        shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(64, 64, 64));
                        shape.TextFrame.TextRange.Font.Size = 11;
                        
                        // Center alignment
                        shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                        
                        // Dark gray borders
                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                        shape.Line.Weight = 1f;
                        shape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(64, 64, 64));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Uniform matrix styling failed: {ex.Message}");
            }
        }

        private void ApplyEnhancedBasicStyling(PowerPoint.Table table, bool hasHeader)
        {
            try
            {
                // Enhanced basic styling with background colors
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        var cell = table.Cell(row, col);
                        var shape = cell.Shape;
                        
                        // Ensure fill is visible
                        shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        shape.Fill.Solid();
                        
                        if (hasHeader && (row == 1 || col == 1))
                        {
                            // Header styling
                            shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(79, 129, 189));
                            shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                        }
                        else
                        {
                            // Data cell styling
                            shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        }
                        
                        // Basic borders
                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                        shape.Line.Weight = 1f;
                        shape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                        
                        // Set basic text properties
                        shape.TextFrame.TextRange.Font.Size = 11;
                        shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Enhanced basic styling failed: {ex.Message}");
                // Ultimate fallback
                ApplyBasicStyling(table, hasHeader);
            }
        }

        private void ApplyBasicStyling(PowerPoint.Table table, bool hasHeader)
        {
            try
            {
                // Very basic styling that should work universally
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        var shape = table.Cell(row, col).Shape;
                        
                        // Only apply the most basic formatting
                        if (hasHeader && row == 1)
                        {
                            shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                        }
                        
                        // Set basic text properties
                        shape.TextFrame.TextRange.Font.Size = 11;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Basic styling failed: {ex.Message}");
            }
        }

        private void BtnStickyNote_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null && app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    // Get user input for sticky note
                    var stickyDialog = new StickyNoteDialog();
                    if (stickyDialog.ShowDialog() == DialogResult.OK)
                    {
                        string noteText = stickyDialog.NoteText;
                        Color noteColor = stickyDialog.NoteColor;
                        
                        var slide = app.ActiveWindow.Selection.SlideRange[1];
                        CreateStickyNote(slide, noteText, noteColor);
                        
                        MessageBox.Show("Sticky note added successfully!\n\nYou can move, resize, or edit the note as needed.", "Sticky Note", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select a slide first to add a sticky note.", "Sticky Note", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating sticky note: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateStickyNote(PowerPoint.Slide slide, string noteText, Color noteColor)
        {
            try
            {
                // Sticky note dimensions
                float noteWidth = 160f;
                float noteHeight = 120f;
                
                // Position in top-right corner with some margin
                float slideWidth = slide.Master.Width;
                float left = slideWidth - noteWidth - 50f;
                float top = 50f;
                
                // Create the main sticky note shape (rounded rectangle)
                var stickyNote = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                    left, top, noteWidth, noteHeight);
                
                // Apply sticky note styling
                ApplyStickyNoteStyling(stickyNote, noteText, noteColor);
                
                // Add a subtle shadow effect to make it look more realistic
                ApplyStickyNoteShadow(stickyNote);
                
                // Name the shape for easy identification
                stickyNote.Name = $"StickyNote_{DateTime.Now.Ticks}";
                
                // Select the sticky note so user can see it
                stickyNote.Select();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create sticky note: {ex.Message}");
            }
        }

        private void ApplyStickyNoteStyling(PowerPoint.Shape stickyNote, string noteText, Color noteColor)
        {
            try
            {
                // Set the background color
                stickyNote.Fill.ForeColor.RGB = ColorTranslator.ToOle(noteColor);
                
                // Configure the text
                if (stickyNote.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    var textFrame = stickyNote.TextFrame;
                    var textRange = textFrame.TextRange;
                    
                    // Set the text content
                    textRange.Text = noteText;
                    
                    // Configure text formatting (using basic Font properties)
                    textRange.Font.Name = "Segoe UI";
                    textRange.Font.Size = 10;
                    textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
                    
                    // Set margins for a realistic sticky note look
                    textFrame.MarginLeft = 10f;
                    textFrame.MarginRight = 10f;
                    textFrame.MarginTop = 8f;
                    textFrame.MarginBottom = 8f;
                }
                
                // Configure border (very subtle)
                stickyNote.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(200, 200, 200));
                stickyNote.Line.Weight = 0.5f;
                
                // Add a very slight rotation for realistic look
                stickyNote.Rotation = 2f; // 2 degrees clockwise
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Sticky note styling failed: {ex.Message}");
            }
        }

        private void ApplyStickyNoteShadow(PowerPoint.Shape stickyNote)
        {
            try
            {
                // Add shadow effect for more realistic appearance
                var shadowFormat = stickyNote.Shadow;
                shadowFormat.Visible = Office.MsoTriState.msoTrue;
                shadowFormat.Type = Office.MsoShadowType.msoShadow6; // Use a simple shadow type
                shadowFormat.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
                shadowFormat.OffsetX = 3f;
                shadowFormat.OffsetY = 3f;
            }
            catch (Exception ex)
            {
                // Shadow is optional, don't fail if it doesn't work
                System.Diagnostics.Debug.WriteLine($"Shadow effect failed: {ex.Message}");
            }
        }

        private void BtnCitation_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null && app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    // Simple input dialog for citation text
                    string citationText = Microsoft.VisualBasic.Interaction.InputBox(
                        "Enter citation text:",
                        "Add Citation",
                        "Source: [Author, Year, Title]");
                    
                    if (!string.IsNullOrEmpty(citationText))
                    {
                        var slide = app.ActiveWindow.Selection.SlideRange[1];
                        CreateCitation(slide, citationText);
                        
                        MessageBox.Show("Citation added to bottom left!\n\nYou can move, resize, or edit the citation as needed.", "Citation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select a slide first to add a citation.", "Citation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating citation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateCitation(PowerPoint.Slide slide, string citationText)
        {
            try
            {
                // Citation dimensions and positioning
                float citationWidth = 400f;
                float citationHeight = 30f;
                
                // Position in bottom left with margin
                float left = 20f;
                float slideHeight = slide.Master.Height;
                float top = slideHeight - citationHeight - 20f;
                
                // Create citation text box
                var citation = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    left, top, citationWidth, citationHeight);
                
                // Apply citation styling
                ApplyCitationStyling(citation, citationText);
                
                // Name the shape for easy identification
                citation.Name = $"Citation_{DateTime.Now.Ticks}";
                
                // Select the citation so user can see it
                citation.Select();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create citation: {ex.Message}");
            }
        }

        private void ApplyCitationStyling(PowerPoint.Shape citation, string citationText)
        {
            try
            {
                // Configure the text
                if (citation.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    var textFrame = citation.TextFrame;
                    var textRange = textFrame.TextRange;
                    
                    // Set the citation text
                    textRange.Text = citationText;
                    
                    // Configure citation text formatting
                    textRange.Font.Name = "Segoe UI";
                    textRange.Font.Size = 9; // Smaller font for citations
                    textRange.Font.Italic = Office.MsoTriState.msoTrue; // Italic for academic style
                    textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(64, 64, 64)); // Dark gray
                    
                    // Set minimal margins
                    textFrame.MarginLeft = 2f;
                    textFrame.MarginRight = 2f;
                    textFrame.MarginTop = 2f;
                    textFrame.MarginBottom = 2f;
                }
                
                // Make the text box transparent (no fill, no border)
                citation.Fill.Visible = Office.MsoTriState.msoFalse;
                citation.Line.Visible = Office.MsoTriState.msoFalse;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Citation styling failed: {ex.Message}");
            }
        }

        private void BtnStandardObjects_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null && app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    // Show standard objects dialog
                    var objectsDialog = new StandardObjectsDialog();
                    if (objectsDialog.ShowDialog() == DialogResult.OK)
                    {
                        string selectedObject = objectsDialog.SelectedObject;
                        
                        if (!string.IsNullOrEmpty(selectedObject))
                        {
                            var slide = app.ActiveWindow.Selection.SlideRange[1];
                            CreateStandardObject(slide, selectedObject);
                            
                            MessageBox.Show($"Standard object added successfully!\n\nYou can move, resize, or edit the object as needed.", "Standard Objects", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select a slide first to add a standard object.", "Standard Objects", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating standard object: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateStandardObject(PowerPoint.Slide slide, string objectType)
        {
            try
            {
                // Parse the object type (remove emoji and get the actual type)
                string type = objectType.Substring(objectType.IndexOf(' ') + 1);
                
                switch (type)
                {
                    case "Title & Subtitle Layout":
                        CreateTitleSubtitleLayout(slide);
                        break;
                    case "Header Text Box":
                        CreateHeaderTextBox(slide);
                        break;
                    case "Content Text Box":
                        CreateContentTextBox(slide);
                        break;
                    case "Callout Box":
                        CreateCalloutBox(slide, "💡", Color.FromArgb(255, 255, 102), "Important Note");
                        break;
                    case "Warning Box":
                        CreateCalloutBox(slide, "⚠️", Color.FromArgb(255, 182, 193), "Warning");
                        break;
                    case "Success Box":
                        CreateCalloutBox(slide, "✅", Color.FromArgb(144, 238, 144), "Success");
                        break;
                    case "Information Box":
                        CreateCalloutBox(slide, "ℹ️", Color.FromArgb(173, 216, 230), "Information");
                        break;
                    case "Date Stamp":
                        CreateDateStamp(slide);
                        break;
                    case "Page Number":
                        CreatePageNumber(slide);
                        break;
                    case "Navigation Arrow":
                        CreateNavigationArrow(slide);
                        break;
                    case "Company Logo Placeholder":
                        CreateLogoPlaceholder(slide);
                        break;
                    default:
                        CreateGenericTextBox(slide, type);
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create standard object: {ex.Message}");
            }
        }

        private void CreateTitleSubtitleLayout(PowerPoint.Slide slide)
        {
            // Create title text box
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 80, 600, 80);
            var titleTextRange = title.TextFrame.TextRange;
            titleTextRange.Text = "Your Title Here";
            titleTextRange.Font.Name = "Segoe UI";
            titleTextRange.Font.Size = 36;
            titleTextRange.Font.Bold = Office.MsoTriState.msoTrue;
            titleTextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(68, 114, 196));
            title.Fill.Visible = Office.MsoTriState.msoFalse;
            title.Line.Visible = Office.MsoTriState.msoFalse;
            title.Name = "TitleTextBox";

            // Create subtitle text box
            var subtitle = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 180, 600, 40);
            var subtitleTextRange = subtitle.TextFrame.TextRange;
            subtitleTextRange.Text = "Your subtitle or description here";
            subtitleTextRange.Font.Name = "Segoe UI";
            subtitleTextRange.Font.Size = 18;
            subtitleTextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(100, 100, 100));
            subtitle.Fill.Visible = Office.MsoTriState.msoFalse;
            subtitle.Line.Visible = Office.MsoTriState.msoFalse;
            subtitle.Name = "SubtitleTextBox";
        }

        private void CreateHeaderTextBox(PowerPoint.Slide slide)
        {
            var header = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 30, 600, 50);
            var textRange = header.TextFrame.TextRange;
            textRange.Text = "Header Text";
            textRange.Font.Name = "Segoe UI";
            textRange.Font.Size = 24;
            textRange.Font.Bold = Office.MsoTriState.msoTrue;
            textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(68, 114, 196));
            header.Fill.Visible = Office.MsoTriState.msoFalse;
            header.Line.Visible = Office.MsoTriState.msoFalse;
            header.Name = "HeaderTextBox";
        }

        private void CreateContentTextBox(PowerPoint.Slide slide)
        {
            var content = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 150, 600, 200);
            var textRange = content.TextFrame.TextRange;
            textRange.Text = "Your content goes here. You can add multiple lines of text, bullet points, or any other content you need for your presentation.";
            textRange.Font.Name = "Segoe UI";
            textRange.Font.Size = 14;
            textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
            content.Fill.Visible = Office.MsoTriState.msoFalse;
            content.Line.Visible = Office.MsoTriState.msoFalse;
            content.Name = "ContentTextBox";
        }

        private void CreateCalloutBox(PowerPoint.Slide slide, string icon, Color bgColor, string defaultText)
        {
            var callout = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, 100, 200, 400, 80);
            callout.Fill.ForeColor.RGB = ColorTranslator.ToOle(bgColor);
            callout.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(150, 150, 150));
            callout.Line.Weight = 1f;
            
            var textRange = callout.TextFrame.TextRange;
            textRange.Text = $"{icon} {defaultText}: Add your important message here";
            textRange.Font.Name = "Segoe UI";
            textRange.Font.Size = 12;
            textRange.Font.Bold = Office.MsoTriState.msoTrue;
            textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
            
            callout.TextFrame.MarginLeft = 15f;
            callout.TextFrame.MarginRight = 15f;
            callout.TextFrame.MarginTop = 10f;
            callout.TextFrame.MarginBottom = 10f;
            
            callout.Name = $"{defaultText}Box";
        }

        private void CreateDateStamp(PowerPoint.Slide slide)
        {
            var dateStamp = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 550, 500, 150, 30);
            var textRange = dateStamp.TextFrame.TextRange;
            textRange.Text = DateTime.Now.ToString("MMM dd, yyyy");
            textRange.Font.Name = "Segoe UI";
            textRange.Font.Size = 10;
            textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(100, 100, 100));
            dateStamp.Fill.Visible = Office.MsoTriState.msoFalse;
            dateStamp.Line.Visible = Office.MsoTriState.msoFalse;
            dateStamp.Name = "DateStamp";
        }

        private void CreatePageNumber(PowerPoint.Slide slide)
        {
            var pageNum = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 650, 500, 50, 30);
            var textRange = pageNum.TextFrame.TextRange;
            textRange.Text = slide.SlideIndex.ToString();
            textRange.Font.Name = "Segoe UI";
            textRange.Font.Size = 12;
            textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(100, 100, 100));
            pageNum.Fill.Visible = Office.MsoTriState.msoFalse;
            pageNum.Line.Visible = Office.MsoTriState.msoFalse;
            pageNum.Name = "PageNumber";
        }

        private void CreateNavigationArrow(PowerPoint.Slide slide)
        {
            var arrow = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightArrow, 300, 400, 100, 40);
            arrow.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(68, 114, 196));
            arrow.Line.Visible = Office.MsoTriState.msoFalse;
            arrow.Name = "NavigationArrow";
        }

        private void CreateLogoPlaceholder(PowerPoint.Slide slide)
        {
            var logo = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 600, 30, 80, 80);
            logo.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(240, 240, 240));
            logo.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(200, 200, 200));
            logo.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;
            
            var textRange = logo.TextFrame.TextRange;
            textRange.Text = "LOGO";
            textRange.Font.Name = "Segoe UI";
            textRange.Font.Size = 10;
            textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(150, 150, 150));
            
            logo.Name = "LogoPlaceholder";
        }

        private void CreateGenericTextBox(PowerPoint.Slide slide, string objectType)
        {
            var textBox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 200, 400, 100);
            var textRange = textBox.TextFrame.TextRange;
            textRange.Text = $"{objectType}\n\nClick to edit this text box and add your content.";
            textRange.Font.Name = "Segoe UI";
            textRange.Font.Size = 12;
            textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
            textBox.Fill.Visible = Office.MsoTriState.msoFalse;
            textBox.Line.Visible = Office.MsoTriState.msoFalse;
            textBox.Name = objectType.Replace(" ", "");
        }

        #endregion

        #region Position Section

        private void BtnAlignLeft_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    app.ActiveWindow.Selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignLefts, Office.MsoTriState.msoFalse);
                    MessageBox.Show("Objects aligned to the left!", "Align", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select objects to align.", "Align", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning objects: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAlignCenter_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    app.ActiveWindow.Selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignCenters, Office.MsoTriState.msoFalse);
                    MessageBox.Show("Objects aligned to center!", "Align", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select objects to align.", "Align", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning objects: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAlignRight_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    app.ActiveWindow.Selection.ShapeRange.Align(Office.MsoAlignCmd.msoAlignRights, Office.MsoTriState.msoFalse);
                    MessageBox.Show("Objects aligned to the right!", "Align", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select objects to align.", "Align", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning objects: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDistribute_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    app.ActiveWindow.Selection.ShapeRange.Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, Office.MsoTriState.msoFalse);
                    MessageBox.Show("Objects distributed evenly!", "Distribute", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select multiple objects to distribute.", "Distribute", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error distributing objects: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnMatchBoth_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    if (shapes.Count >= 2)
                    {
                        // Use the last selected shape as reference
                        var referenceShape = shapes[shapes.Count];
                        float referenceHeight = referenceShape.Height;
                        float referenceWidth = referenceShape.Width;

                        // Apply dimensions to all other shapes
                        for (int i = 1; i < shapes.Count; i++)
                        {
                            shapes[i].Height = referenceHeight;
                            shapes[i].Width = referenceWidth;
                        }

                        MessageBox.Show($"Matched height and width to the last selected shape!", "Match Dimensions", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Please select at least two shapes to match dimensions.", "Match Dimensions", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select shapes to match dimensions.", "Match Dimensions", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching dimensions: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnMatchHeight_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    if (shapes.Count >= 2)
                    {
                        // Use the last selected shape as reference
                        var referenceShape = shapes[shapes.Count];
                        float referenceHeight = referenceShape.Height;

                        // Apply height to all other shapes
                        for (int i = 1; i < shapes.Count; i++)
                        {
                            shapes[i].Height = referenceHeight;
                        }

                        MessageBox.Show($"Matched height to the last selected shape!", "Match Height", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Please select at least two shapes to match height.", "Match Height", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select shapes to match height.", "Match Height", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching height: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnMatchWidth_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    if (shapes.Count >= 2)
                    {
                        // Use the last selected shape as reference
                        var referenceShape = shapes[shapes.Count];
                        float referenceWidth = referenceShape.Width;

                        // Apply width to all other shapes
                        for (int i = 1; i < shapes.Count; i++)
                        {
                            shapes[i].Width = referenceWidth;
                        }

                        MessageBox.Show($"Matched width to the last selected shape!", "Match Width", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Please select at least two shapes to match width.", "Match Width", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select shapes to match width.", "Match Width", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error matching width: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnMakeVertical_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    int count = 0;
                    
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        shapes[i].Rotation = 90f;
                        count++;
                    }
                    
                    MessageBox.Show($"Rotated {count} shape(s) to vertical (90°)!", "Make Vertical", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select shapes to rotate vertically.", "Make Vertical", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error rotating shapes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnMakeHorizontal_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    int count = 0;
                    
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        shapes[i].Rotation = 0f;
                        count++;
                    }
                    
                    MessageBox.Show($"Rotated {count} shape(s) to horizontal (0°)!", "Make Horizontal", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select shapes to rotate horizontally.", "Make Horizontal", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error rotating shapes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSwapLocations_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    if (shapes.Count == 2)
                    {
                        // Get positions of both shapes
                        float shape1Left = shapes[1].Left;
                        float shape1Top = shapes[1].Top;
                        float shape2Left = shapes[2].Left;
                        float shape2Top = shapes[2].Top;

                        // Swap the positions
                        shapes[1].Left = shape2Left;
                        shapes[1].Top = shape2Top;
                        shapes[2].Left = shape1Left;
                        shapes[2].Top = shape1Top;

                        MessageBox.Show("Swapped positions of the two selected shapes!", "Swap Locations", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Please select exactly two shapes to swap their locations.", "Swap Locations", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select two shapes to swap their locations.", "Swap Locations", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error swapping shape locations: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Size Section

        private void CmbSlideSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (nudWidth != null && nudHeight != null)
            {
                switch (cmbSlideSize.SelectedIndex)
            {
                case 0: // Standard 4:3
                    nudWidth.Value = 10m;
                    nudHeight.Value = 7.5m;
                    break;
                case 1: // Widescreen 16:9
                        nudWidth.Value = 13.3m;
                    nudHeight.Value = 7.5m;
                    break;
                    case 2: // Custom - don't change values
                    break;
                }
            }
        }

        private void BtnApplySize_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null && nudWidth != null && nudHeight != null)
                {
                    app.ActivePresentation.PageSetup.SlideWidth = (float)nudWidth.Value * 72; // Convert inches to points
                    app.ActivePresentation.PageSetup.SlideHeight = (float)nudHeight.Value * 72;
                    MessageBox.Show("Slide size applied successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please open a presentation first.", "No Presentation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying slide size: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Shape Section

        private void BtnAlignProcessChain_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    // Align selected shapes in a process chain
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    if (shapes.Count > 1)
                    {
                        // Sort shapes by X position
                        var sortedShapes = shapes.Cast<PowerPoint.Shape>().OrderBy(s => s.Left).ToArray();
                        
                        // Align them horizontally with equal spacing
                        float totalWidth = sortedShapes.Sum(s => s.Width);
                        float spacing = (app.ActiveWindow.Width - totalWidth) / (sortedShapes.Length + 1);
                        float currentX = spacing;
                        
                        foreach (var shape in sortedShapes)
                        {
                            shape.Left = currentX;
                            currentX += shape.Width + spacing;
                        }
                        
                        MessageBox.Show("Process chain aligned!", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Please select multiple shapes to align in a process chain.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select shapes to align in a process chain.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning process chain: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAlignAngles_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    if (shapes.Count > 1)
                    {
                        // Align shapes at their angles (corners)
                        var firstShape = shapes[1];
                        float targetAngle = firstShape.Rotation;
                        
                        foreach (PowerPoint.Shape shape in shapes)
                        {
                            shape.Rotation = targetAngle;
                        }
                        
                        MessageBox.Show("Shapes aligned at angles!", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Please select multiple shapes to align at angles.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select shapes to align at angles.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning angles: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAlignToProcessArrow_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    if (shapes.Count > 1)
                    {
                        // Find the first arrow shape and align other shapes to it
                        PowerPoint.Shape arrowShape = null;
                        foreach (PowerPoint.Shape shape in shapes)
                        {
                            if (shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRightArrow ||
                                shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeLeftArrow ||
                                shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeUpArrow ||
                                shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeDownArrow)
                            {
                                arrowShape = shape;
                                break;
                            }
                        }
                        
                        if (arrowShape != null)
                        {
                            // Align other shapes to the arrow's center
                            foreach (PowerPoint.Shape shape in shapes)
                            {
                                if (shape != arrowShape)
                                {
                                    shape.Top = arrowShape.Top + (arrowShape.Height - shape.Height) / 2;
                                }
                            }
                            
                            MessageBox.Show("Shapes aligned to process arrow!", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Please include an arrow shape in your selection.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select multiple shapes including an arrow.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select shapes to align to process arrow.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning to process arrow: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAdjustPentagonHeader_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    foreach (PowerPoint.Shape shape in shapes)
                    {
                        if (shape.AutoShapeType == Office.MsoAutoShapeType.msoShapePentagon)
                        {
                            // Adjust pentagon header properties
                            shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightBlue);
                            shape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.DarkBlue);
                            shape.Line.Weight = 2;
                            
                            // Add text if not present
                            if (string.IsNullOrEmpty(shape.TextFrame.TextRange.Text))
                            {
                                shape.TextFrame.TextRange.Text = "Header";
                                shape.TextFrame.TextRange.Font.Size = 14;
                                shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                            }
                        }
                    }
                    
                    MessageBox.Show("Pentagon headers adjusted!", "Shape Adjustment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select pentagon shapes to adjust.", "Shape Adjustment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adjusting pentagon headers: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAlignBlockArrows_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    if (shapes.Count > 1)
                    {
                        // Find block arrows and align them
                        var blockArrows = new List<PowerPoint.Shape>();
                        foreach (PowerPoint.Shape shape in shapes)
                        {
                            if (shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRightArrow ||
                                shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeLeftArrow ||
                                shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeUpArrow ||
                                shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeDownArrow ||
                                shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRightArrowCallout ||
                                shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeLeftArrowCallout)
                            {
                                blockArrows.Add(shape);
                            }
                        }
                        
                        if (blockArrows.Count > 1)
                        {
                            // Align block arrows vertically
                            float centerY = blockArrows.Average(s => s.Top + s.Height / 2);
                            foreach (var arrow in blockArrows)
                            {
                                arrow.Top = centerY - arrow.Height / 2;
                            }
                            
                            MessageBox.Show("Block arrows aligned!", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Please select multiple block arrow shapes.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select multiple shapes including block arrows.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select shapes to align block arrows.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning block arrows: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAlignRoundedRectangleArrows_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    if (shapes.Count > 1)
                    {
                        // Find rounded rectangles and align them
                        var roundedRects = new List<PowerPoint.Shape>();
                        foreach (PowerPoint.Shape shape in shapes)
                        {
                            if (shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRoundedRectangle ||
                                shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRoundedRectangularCallout)
                            {
                                roundedRects.Add(shape);
                            }
                        }
                        
                        if (roundedRects.Count > 1)
                        {
                            // Align rounded rectangles horizontally with equal spacing
                            var sortedRects = roundedRects.OrderBy(s => s.Left).ToArray();
                            float totalWidth = sortedRects.Sum(s => s.Width);
                            float spacing = (app.ActiveWindow.Width - totalWidth) / (sortedRects.Length + 1);
                            float currentX = spacing;
                            
                            foreach (var rect in sortedRects)
                            {
                                rect.Left = currentX;
                                currentX += rect.Width + spacing;
                            }
                            
                            MessageBox.Show("Rounded rectangle arrows aligned!", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Please select multiple rounded rectangle shapes.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select multiple shapes including rounded rectangles.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select shapes to align rounded rectangle arrows.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning rounded rectangle arrows: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Color Section

        private void BtnFillColor_Click(object sender, EventArgs e)
        {
            try
            {
                var colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    var app = Globals.ThisAddIn.Application;
                    if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        app.ActiveWindow.Selection.ShapeRange.Fill.ForeColor.RGB = ColorTranslator.ToOle(colorDialog.Color);
                        MessageBox.Show("Fill color applied!", "Color", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Please select an object to change its fill color.", "Color", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error changing fill color: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnTextColor_Click(object sender, EventArgs e)
        {
            try
            {
                var colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    var app = Globals.ThisAddIn.Application;
                    if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    {
                        app.ActiveWindow.Selection.TextRange.Font.Color.RGB = ColorTranslator.ToOle(colorDialog.Color);
                        MessageBox.Show("Text color applied!", "Color", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Please select text to change its color.", "Color", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error changing text color: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnOutlineColor_Click(object sender, EventArgs e)
        {
            try
            {
                var colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    var app = Globals.ThisAddIn.Application;
                    if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        app.ActiveWindow.Selection.ShapeRange.Line.ForeColor.RGB = ColorTranslator.ToOle(colorDialog.Color);
                        MessageBox.Show("Outline color applied!", "Color", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Please select an object to change its outline color.", "Color", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    }
                }
                catch (Exception ex)
            {
                MessageBox.Show($"Error changing outline color: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Text Section

        private void BtnBold_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    app.ActiveWindow.Selection.TextRange.Font.Bold = 
                        app.ActiveWindow.Selection.TextRange.Font.Bold == Office.MsoTriState.msoTrue ? 
                        Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
                    MessageBox.Show("Bold formatting toggled!", "Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select text to format.", "Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error formatting text: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnItalic_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    app.ActiveWindow.Selection.TextRange.Font.Italic = 
                        app.ActiveWindow.Selection.TextRange.Font.Italic == Office.MsoTriState.msoTrue ? 
                        Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
                    MessageBox.Show("Italic formatting toggled!", "Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select text to format.", "Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error formatting text: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnUnderline_Click(object sender, EventArgs e)
        {
            try
                {
                    var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    app.ActiveWindow.Selection.TextRange.Font.Underline = 
                        app.ActiveWindow.Selection.TextRange.Font.Underline == Office.MsoTriState.msoTrue ? 
                        Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
                    MessageBox.Show("Underline formatting toggled!", "Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select text to format.", "Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
            {
                MessageBox.Show($"Error formatting text: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnBullets_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    var textRange = app.ActiveWindow.Selection.TextRange;
                    textRange.ParagraphFormat.Bullet.Visible = 
                        textRange.ParagraphFormat.Bullet.Visible == Office.MsoTriState.msoTrue ? 
                        Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
                    MessageBox.Show("Bullet formatting toggled!", "Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select text to add bullets.", "Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error formatting bullets: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Text Wrapping Section

        private void BtnWrapText_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }
        private void BtnWrapText_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnNoWrapText_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }
        private void BtnNoWrapText_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnWrapText_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    int count = 0;
                    
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        if (shapes[i].HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            shapes[i].TextFrame2.WordWrap = Office.MsoTriState.msoTrue;
                            count++;
                        }
                    }
                    
                    if (count > 0)
                    {
                        MessageBox.Show($"Text wrapping enabled for {count} shape(s)!", "Text Wrap", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("No text shapes found in selection.", "Text Wrap", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select shapes to enable text wrapping.", "Text Wrap", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error enabling text wrap: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnNoWrapText_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    int count = 0;
                    
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        if (shapes[i].HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            shapes[i].TextFrame2.WordWrap = Office.MsoTriState.msoFalse;
                            count++;
                        }
                    }
                    
                    if (count > 0)
                    {
                        MessageBox.Show($"Text wrapping disabled for {count} shape(s)!", "No Text Wrap", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("No text shapes found in selection.", "No Text Wrap", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select shapes to disable text wrapping.", "No Text Wrap", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error disabling text wrap: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Navigation & View Section

        private void BtnZoomIn_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.View.Zoom < 400)
                {
                    app.ActiveWindow.View.Zoom = app.ActiveWindow.View.Zoom + 10;
                    MessageBox.Show($"Zoom: {app.ActiveWindow.View.Zoom}%", "Zoom", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error zooming: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnZoomOut_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.View.Zoom > 10)
                {
                    app.ActiveWindow.View.Zoom = app.ActiveWindow.View.Zoom - 10;
                    MessageBox.Show($"Zoom: {app.ActiveWindow.View.Zoom}%", "Zoom", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error zooming: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnFitToWindow_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                app.ActiveWindow.View.ZoomToFit = Office.MsoTriState.msoTrue;
                MessageBox.Show("Zoom fit to window!", "Zoom", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error fitting to window: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Expert Tools Section

        private void BtnFreeWebinar_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("🎓 This would open a free PowerPoint training webinar!\n\nLearn advanced PowerPoint techniques and tips from experts.", 
                              "Free Webinar", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error accessing webinar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #endregion

        private void lblWidth_Click(object sender, EventArgs e)
        {

        }

        // Navigation View button hover handlers
        private void BtnNavButton_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }
        private void BtnNavButton_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void ClearTableHeaderText(PowerPoint.Table table)
        {
            try
            {
                // Clear the default header text in the first row
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    var cell = table.Cell(1, col);
                    cell.Shape.TextFrame.TextRange.Text = "";
                }
            }
            catch (Exception ex)
            {
                // If clearing header text fails, just continue without error
                // This ensures table creation still works even if header clearing fails
                System.Diagnostics.Debug.WriteLine($"Could not clear header text: {ex.Message}");
            }
        }
    }
}