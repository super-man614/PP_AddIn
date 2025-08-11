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
            try
            {
            InitializeComponent();
            OptimizeRemainingPanels(); // Optimize panels not yet optimized in Designer
            SetupEventHandlers();
            SetInitialValues();
                // Delay image loading until after control is fully loaded
                this.Load += TaskPaneControl_LoadImages;
            SetupTooltips(); // Setup tooltips for all buttons
            this.KeyDown += TaskPaneControl_KeyDown;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in TaskPaneControl constructor: {ex.Message}");
                // Still initialize basic functionality even if there are errors
                try
                {
                    InitializeComponent();
                }
                catch
                {
                    // If even basic initialization fails, log it
                    System.Diagnostics.Debug.WriteLine("Critical error: InitializeComponent failed");
                }
            }
        }

        private void TaskPaneControl_LoadImages(object sender, EventArgs e)
        {
            try
            {
                // Remove the event handler to prevent multiple calls
                this.Load -= TaskPaneControl_LoadImages;
                
                // Load images after the control is fully loaded
                LoadButtonImages();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading images: {ex.Message}");
            }
        }

        private PowerPoint.Slide GetActiveSlideOrNull()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null || app.ActivePresentation == null || app.ActiveWindow == null)
                {
                    return null;
                }

                var selection = app.ActiveWindow.Selection;
                if (selection != null && selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides && selection.SlideRange != null && selection.SlideRange.Count > 0)
                {
                    return selection.SlideRange[1];
                }

                var viewSlide = app.ActiveWindow.View?.Slide;
                if (viewSlide != null)
                {
                    return viewSlide;
                }

                var pres = app.ActivePresentation;
                return pres?.Slides?.Count > 0 ? pres.Slides[1] : null;
            }
            catch
            {
                return null;
            }
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
            try
        {
            // Size section events
            if (cmbSlideSize != null)
                cmbSlideSize.SelectedIndexChanged += CmbSlideSize_SelectedIndexChanged;
            if (btnApplySize != null)
                btnApplySize.Click += BtnApplySize_Click;
            if (cmbAlign != null)
                cmbAlign.SelectedIndexChanged += CmbAlign_SelectedIndexChanged;
            if (cmbStretch != null)
                cmbStretch.SelectedIndexChanged += CmbStretch_SelectedIndexChanged;
            if (cmbFill != null)
                cmbFill.SelectedIndexChanged += CmbFill_SelectedIndexChanged;
            if (btnMagicResizer != null)
                btnMagicResizer.Click += BtnMagicResizer_Click;
            
            // OPTIMIZED: Apply hover effects to ALL buttons at once using our utility
            // This replaces ~250 lines of duplicate hover event handlers!
            ButtonHoverUtility.EnableHoverEffectsForContainer(this, recursive: true);
            
            // Load current slides
            this.Load += TaskPaneControl_Load;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting up event handlers: {ex.Message}");
            }
        }

        private void SetInitialValues()
        {
            try
        {
            // Populate slide size combo box with smart presets
            if (cmbSlideSize != null)
            {
                cmbSlideSize.Items.Clear();
                cmbSlideSize.Items.AddRange(sizePresets.Keys.ToArray());
                cmbSlideSize.Items.Add("Custom");
                cmbSlideSize.SelectedIndex = 1; // Default to 16:9 Widescreen
            }
            
            // Populate align combo box
            if (cmbAlign != null)
            {
                cmbAlign.Items.Clear();
                cmbAlign.Items.AddRange(new object[] {
                    "Align Width",
                    "Align Height",
                    "Align Width & Height"
                });
            }
            
            // Populate stretch combo box
            if (cmbStretch != null)
            {
                cmbStretch.Items.Clear();
                cmbStretch.Items.AddRange(new object[] {
                    "Stretch Left",
                    "Stretch Right",
                    "Stretch Up",
                    "Stretch Down"
                });
            }
            
            // Populate fill combo box
            if (cmbFill != null)
            {
                cmbFill.Items.Clear();
                cmbFill.Items.AddRange(new object[] {
                    "Fill Left",
                    "Fill Right",
                    "Fill Up",
                    "Fill Down"
                });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting initial values: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads custom images for buttons from the icons directory
        /// </summary>
        private void LoadButtonImages()
        {
            try
            {
                // Load presentation button images
                LoadPresentationButtonImages();

                // Load wizard button images
                LoadWizardButtonImages();

                // Load smart elements button images
                LoadSmartElementsButtonImages();

                // Load position button images
                LoadPositionButtonImages();

                // Shape, Color, and Text sections use emoji text instead of images
                SetupEmojiButtons();
            }
            catch (Exception ex)
            {
                // Silently fail - buttons will use emoji fallbacks
                System.Diagnostics.Debug.WriteLine($"Error loading button images: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads an image for a specific button with multiple fallback paths
        /// Falls back to emoji text if image is not found
        /// </summary>
        private void LoadImageForButton(Button button, string imageName, string subfolder = "")
        {
            try
            {
                // Add null check for button
                if (button == null)
                {
                    System.Diagnostics.Debug.WriteLine($"Button is null when trying to load image {imageName}");
                    return;
                }

                string imagePath = FindImagePath(imageName, subfolder);
                if (!string.IsNullOrEmpty(imagePath) && System.IO.File.Exists(imagePath))
                {
                    try
                    {
                        // Successfully load the image
                        button.BackgroundImage = Image.FromFile(imagePath);
                        button.BackgroundImageLayout = ImageLayout.Stretch;
                        button.Text = ""; // Clear text to show image
                        System.Diagnostics.Debug.WriteLine($"Loaded image: {imageName}");
                    }
                    catch (Exception imgEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to load image file {imagePath}: {imgEx.Message}");
                        // Keep any existing text/emoji as fallback
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Image not found: {imageName} in {subfolder} - keeping text/emoji fallback");
                    // Image not found - button will use its designed text/emoji
                    button.BackgroundImage = null; // Ensure no background image
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in LoadImageForButton for {imageName}: {ex.Message}");
                // Ensure button remains functional with text fallback
                if (button != null)
                {
                    button.BackgroundImage = null;
                }
            }
        }

        /// <summary>
        /// Finds the correct path for an image file using multiple fallback locations
        /// Optimized for VSTO deployment scenarios
        /// </summary>
        private string FindImagePath(string imageName, string subfolder = "")
        {
            var searchPaths = new List<string>();
            
            try
            {
                // Get the assembly location (where the .dll is deployed)
                string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string assemblyDir = System.IO.Path.GetDirectoryName(assemblyPath);
                
                // For VSTO deployment, images are typically in the same folder as the assembly
                if (!string.IsNullOrEmpty(subfolder))
                {
                    searchPaths.Add(System.IO.Path.Combine(assemblyDir, "icons", subfolder, imageName));
                    searchPaths.Add(System.IO.Path.Combine(assemblyDir, subfolder, imageName));
                }
                searchPaths.Add(System.IO.Path.Combine(assemblyDir, "icons", imageName));
                searchPaths.Add(System.IO.Path.Combine(assemblyDir, imageName));
                
                // Additional fallback locations for different deployment scenarios
                var additionalBaseDirs = new List<string>();
                
                // VSTO deployment folder (typically under user's Local folder)
                if (assemblyDir.Contains("Users") && assemblyDir.Contains("Local"))
                {
                    additionalBaseDirs.Add(assemblyDir);
                }
                
                // Application data folder
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string vstoAppPath = System.IO.Path.Combine(appDataPath, "Apps", "2.0");
                if (System.IO.Directory.Exists(vstoAppPath))
                {
                    // Look for VSTO deployment folder
                    var vstoDirs = System.IO.Directory.GetDirectories(vstoAppPath, "*", System.IO.SearchOption.AllDirectories)
                        .Where(d => d.Contains("my-addin") || System.IO.File.Exists(System.IO.Path.Combine(d, "my-addin.dll")));
                    additionalBaseDirs.AddRange(vstoDirs);
                }
                
                // Add these paths to search list
                foreach (var baseDir in additionalBaseDirs)
                {
                    if (!string.IsNullOrEmpty(subfolder))
                    {
                        searchPaths.Add(System.IO.Path.Combine(baseDir, "icons", subfolder, imageName));
                        searchPaths.Add(System.IO.Path.Combine(baseDir, subfolder, imageName));
                    }
                    searchPaths.Add(System.IO.Path.Combine(baseDir, "icons", imageName));
                    searchPaths.Add(System.IO.Path.Combine(baseDir, imageName));
                }
                
                // Development fallbacks
                if (assemblyDir.Contains("bin"))
                {
                    var projectDir = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(assemblyDir));
                    if (!string.IsNullOrEmpty(projectDir))
                    {
                        if (!string.IsNullOrEmpty(subfolder))
                        {
                            searchPaths.Add(System.IO.Path.Combine(projectDir, "icons", subfolder, imageName));
                        }
                        searchPaths.Add(System.IO.Path.Combine(projectDir, "icons", imageName));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error building search paths: {ex.Message}");
            }
            
            // Search through all paths
            foreach (var path in searchPaths)
            {
                try
                {
                    if (System.IO.File.Exists(path))
                    {
                        System.Diagnostics.Debug.WriteLine($"Found image at: {path}");
                        return path;
                    }
                }
                catch
                {
                    // Continue searching if path access fails
                    continue;
                }
            }
            
            System.Diagnostics.Debug.WriteLine($"Image not found: {imageName} in subfolder: {subfolder}");
            System.Diagnostics.Debug.WriteLine($"Searched paths: {string.Join("; ", searchPaths)}");
            
            return null;
        }

        /// <summary>
        /// Loads images for wizard buttons
        /// </summary>
        private void LoadWizardButtonImages()
        {
            try
            {
                // Wizard button images mapping
                var wizardButtons = new Dictionary<Button, string>
                {
                    { btnAgenda, "agenda.png" },
                    { btnMaster, "master.png" },
                    { btnElement, "element.png" },
                    { btnText, "text.png" },
                    { btnFormat, "format.png" },
                    { btnMap, "map.png" }
                };

                foreach (var button in wizardButtons)
                {
                    if (button.Key != null)
                    {
                        LoadImageForButton(button.Key, button.Value, "wizzards", true);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading wizard button images: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets up emoji text for Shape, Color, and Text sections (no image loading)
        /// </summary>
        private void SetupEmojiButtons()
        {
            try
            {
                // Shape section buttons - ensure emoji text is visible and no background image
                // (Text already set in Designer, just ensure no background images)
                var shapeButtons = new Button[]
                {
                    btnAlignProcessChain,    // "🔗"
                    btnAlignAngles,          // "📐"
                    btnAlignToProcessArrow,  // "➡️"
                    btnAdjustPentagonHeader, // "🔷"
                    btnAlignBlockArrows,     // "▶️"
                    btnAlignRoundedRectangleRadius // "🔲"
                };

                foreach (var button in shapeButtons)
                {
                    if (button != null)
                    {
                        button.BackgroundImage = null; // Remove any background image
                        button.UseVisualStyleBackColor = false;
                        
                        // Ensure emoji font is set correctly
                        if (button.Font == null || !button.Font.Name.Contains("Emoji"))
                        {
                            button.Font = new System.Drawing.Font("Segoe UI Emoji", 7F);
                        }
                        
                        // Ensure text is visible and not empty
                        if (string.IsNullOrEmpty(button.Text))
                        {
                            // Set default emoji if text is missing
                            switch (button.Name)
                            {
                                case "btnAlignProcessChain":
                                    button.Text = "📐";
                                    break;
                                case "btnAlignAngles": 
                                    button.Text = "📐";
                                    break;
                                case "btnAlignToProcessArrow":
                                    button.Text = "➡️";
                                    break;
                                case "btnAdjustPentagonHeader":
                                    button.Text = "🔷";
                                    break;
                                case "btnAlignBlockArrows":
                                    button.Text = "▶️";
                                    break;
                                case "btnAlignRoundedRectangleRadius":
                                    button.Text = "🔲";
                                    break;
                            }
                        }
                        
                        System.Diagnostics.Debug.WriteLine($"✅ Shape button {button.Name}: text='{button.Text}', font={button.Font.Name}");
                    }
                }

                // Color section buttons - ensure emoji text is visible
                // (Text already set in Designer, just ensure no background images)
                var colorButtons = new Button[]
                {
                    btnFillColor,    // "🎨"
                    btnTextColor,    // "A"
                    btnOutlineColor  // "◯"
                };

                foreach (var button in colorButtons)
                {
                    if (button != null)
                    {
                        button.BackgroundImage = null; // Remove any background image
                        button.UseVisualStyleBackColor = false;
                        // Text is already set in Designer file
                    }
                }

                // Text section buttons - ensure emoji text is visible
                // (Text already set in Designer, just ensure no background images)
                    var textButtons = new Button[] 
                    { 
                    btnBold,        // "B"
                    btnItalic,      // "I"
                    btnUnderline,   // "U"
                    btnBullets,     // "•"
                    btnWrapText,    // "📦"
                    btnNoWrapText   // "📄"
                    };

                    foreach (var button in textButtons)
                    {
                        if (button != null)
                        {
                        button.BackgroundImage = null; // Remove any background image
                        button.UseVisualStyleBackColor = false;
                        // Text is already set in Designer file
                    }
                }

                System.Diagnostics.Debug.WriteLine("Emoji buttons setup completed for Shape, Color, and Text sections");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting up emoji buttons: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets up tooltips for all buttons
        /// </summary>
        private void SetupTooltips()
        {
            try
        {
            // Create tooltip control
            var tooltip = new ToolTip();
            tooltip.AutoPopDelay = 5000;
            tooltip.InitialDelay = 1000;
            tooltip.ReshowDelay = 500;
            tooltip.ShowAlways = true;

                // Presentation section tooltips - with null checks
                if (btnNew != null) tooltip.SetToolTip(btnNew, "New Presentation");
                if (btnOpen != null) tooltip.SetToolTip(btnOpen, "Open Presentation");
                if (btnSave != null) tooltip.SetToolTip(btnSave, "Save Presentation");
                if (btnSaveAs != null) tooltip.SetToolTip(btnSaveAs, "Save As");
                if (btnPrint != null) tooltip.SetToolTip(btnPrint, "Print");
                //if (btnShare != null) tooltip.SetToolTip(btnShare, "Share");

                // Wizards section tooltips - with null checks
                if (btnAgenda != null) tooltip.SetToolTip(btnAgenda, "Agenda");
                if (btnMaster != null) tooltip.SetToolTip(btnMaster, "Master");
                if (btnElement != null) tooltip.SetToolTip(btnElement, "Element");
                if (btnText != null) tooltip.SetToolTip(btnText, "Text");
                if (btnFormat != null) tooltip.SetToolTip(btnFormat, "Format");
                if (btnMap != null) tooltip.SetToolTip(btnMap, "Map");

                // Smart Elements section tooltips - with null checks
                if (btnChart != null) tooltip.SetToolTip(btnChart, "Chart");
                if (btnDiagram != null) tooltip.SetToolTip(btnDiagram, "Diagram");
                if (btnTable != null) tooltip.SetToolTip(btnTable, "Table");
                if (btnMatrixTable != null) tooltip.SetToolTip(btnMatrixTable, "Matrix Table");
                if (btnStickyNote != null) tooltip.SetToolTip(btnStickyNote, "Sticky Note");
                if (btnCitation != null) tooltip.SetToolTip(btnCitation, "Citation");
                if (btnStandardObjects != null) tooltip.SetToolTip(btnStandardObjects, "Standard Objects");

                // Position section tooltips - with null checks
                if (btnAlignLeft != null) tooltip.SetToolTip(btnAlignLeft, "Align Left\nAlign objects to left edge (Ctrl: to slide edge)");
                if (btnAlignCenter != null) tooltip.SetToolTip(btnAlignCenter, "Align Center\nAlign objects to center (Ctrl: to slide center)");
                if (btnAlignRight != null) tooltip.SetToolTip(btnAlignRight, "Align Right\nAlign objects to right edge (Ctrl: to slide edge)");
                if (btnAlignTop != null) tooltip.SetToolTip(btnAlignTop, "Align Top\nAlign objects to top edge (Ctrl: to slide top)");
                if (btnAlignBottom != null) tooltip.SetToolTip(btnAlignBottom, "Align Bottom\nAlign objects to bottom edge (Ctrl: to slide bottom)");
                if (btnAlignMiddle != null) tooltip.SetToolTip(btnAlignMiddle, "Align Middle\nAlign objects to vertical middle (Ctrl: to slide middle)");
                if (btnDockLeft != null) tooltip.SetToolTip(btnDockLeft, "Dock Left\nMove objects to touch left edge (Ctrl: to slide edge)");
                if (btnDockRight != null) tooltip.SetToolTip(btnDockRight, "Dock Right\nMove objects to touch right edge (Ctrl: to slide edge)");
                if (btnDockTop != null) tooltip.SetToolTip(btnDockTop, "Dock Top\nMove objects to touch top edge (Ctrl: to slide top)");
                if (btnDockBottom != null) tooltip.SetToolTip(btnDockBottom, "Dock Bottom\nMove objects to touch bottom edge (Ctrl: to slide bottom)");
                if (btnDistribute != null) tooltip.SetToolTip(btnDistribute, "Distribute\nDistribute objects evenly");
                if (btnDistributeHorizontal != null) tooltip.SetToolTip(btnDistributeHorizontal, "Distribute Horizontal\nDistribute horizontally (Ctrl: across slide)");
                if (btnDistributeVertical != null) tooltip.SetToolTip(btnDistributeVertical, "Distribute Vertical\nDistribute vertically (Ctrl: across slide)");
                if (btnMatchBoth != null) tooltip.SetToolTip(btnMatchBoth, "Match Both\nMatch width and height to master object");
                if (btnMatchHeight != null) tooltip.SetToolTip(btnMatchHeight, "Match Height\nMatch height to master object");
                if (btnMatchWidth != null) tooltip.SetToolTip(btnMatchWidth, "Match Width\nMatch width to master object");
                if (btnGoldenCanon != null) tooltip.SetToolTip(btnGoldenCanon, "Golden Canon\nAlign in golden ratio (1:2 margin ratio)");
                if (btnAlignMatrix != null) tooltip.SetToolTip(btnAlignMatrix, "Align Matrix\nArrange objects in matrix grid");
                if (btnSliceShape != null) tooltip.SetToolTip(btnSliceShape, "Slice Shape\nSlice or multiply shape into grid");
                if (btnDuplicateRight != null) tooltip.SetToolTip(btnDuplicateRight, "Duplicate Right\nDuplicate objects to the right");
                if (btnCenterTopLeft != null) tooltip.SetToolTip(btnCenterTopLeft, "Center on Top Left\nCenter objects on master's top-left corner");
                if (btnSavePosition != null) tooltip.SetToolTip(btnSavePosition, "Save Position\nSave position and size of selected objects");
                if (btnApplyPosition != null) tooltip.SetToolTip(btnApplyPosition, "Apply Position\nApply saved position and size to selected objects");
                if (btnRemoveMarginObjects != null) tooltip.SetToolTip(btnRemoveMarginObjects, "Remove Margin Objects\nRemove objects outside the main slide layout");

                // Shape section tooltips - with null checks
                if (btnAlignProcessChain != null) tooltip.SetToolTip(btnAlignProcessChain, "Align Process Chain");
                if (btnAlignAngles != null) tooltip.SetToolTip(btnAlignAngles, "Align Angles");
                if (btnAlignToProcessArrow != null) tooltip.SetToolTip(btnAlignToProcessArrow, "Align to Process Arrow");
                if (btnAdjustPentagonHeader != null) tooltip.SetToolTip(btnAdjustPentagonHeader, "Adjust Pentagon Header");
                if (btnAlignBlockArrows != null) tooltip.SetToolTip(btnAlignBlockArrows, "Align Block Arrows");
                if (btnAlignRoundedRectangleRadius != null) tooltip.SetToolTip(btnAlignRoundedRectangleRadius, "Align Rounded Rectangle Radius");

                // Transform section tooltips - with null checks
                if (btnMakeVertical != null) tooltip.SetToolTip(btnMakeVertical, "Make Vertical");
                if (btnMakeHorizontal != null) tooltip.SetToolTip(btnMakeHorizontal, "Make Horizontal");
                if (btnSwapLocations != null) tooltip.SetToolTip(btnSwapLocations, "Swap Locations");

                // Size section tooltips - with null checks
                if (btnApplySize != null) tooltip.SetToolTip(btnApplySize, "Apply size with intelligent content scaling");

                // Colors section tooltips - with null checks
                if (btnFillColor != null) tooltip.SetToolTip(btnFillColor, "Fill Color");
                if (btnTextColor != null) tooltip.SetToolTip(btnTextColor, "Text Color");
                if (btnOutlineColor != null) tooltip.SetToolTip(btnOutlineColor, "Outline Color");

                // Text section tooltips - with null checks
                if (btnBold != null) tooltip.SetToolTip(btnBold, "Bold");
                if (btnItalic != null) tooltip.SetToolTip(btnItalic, "Italic");
                if (btnUnderline != null) tooltip.SetToolTip(btnUnderline, "Underline");
                if (btnBullets != null) tooltip.SetToolTip(btnBullets, "Bullet Points");
                if (btnWrapText != null) tooltip.SetToolTip(btnWrapText, "Wrap Text");
                if (btnNoWrapText != null) tooltip.SetToolTip(btnNoWrapText, "No Wrap Text");

                // Navigation section tooltips - with null checks
                if (btnZoomIn != null) tooltip.SetToolTip(btnZoomIn, "Zoom In");
                if (btnZoomOut != null) tooltip.SetToolTip(btnZoomOut, "Zoom Out");
                if (btnFitToWindow != null) tooltip.SetToolTip(btnFitToWindow, "Fit to Window");

                // Expert Tools section tooltips - with null checks
                if (btnFreeWebinar != null) tooltip.SetToolTip(btnFreeWebinar, "Free Webinar");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting up tooltips: {ex.Message}");
            }
        }

        #region Event Handlers

        private void TaskPaneControl_Load(object sender, EventArgs e)
        {
            // Initialize the interface safely without trying to access PowerPoint presentation on startup
            try
            {
                // Set default slide size selection without accessing PowerPoint
                if (cmbSlideSize != null && cmbSlideSize.Items.Count > 1)
                {
                    cmbSlideSize.SelectedIndex = 1; // Default to 16:9 Widescreen
                }
                System.Diagnostics.Debug.WriteLine("Task pane loaded successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error during task pane load: {ex.Message}");
            }
        }

        #region Presentation Section

        private void BtnNew_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                app.Presentations.Add();
                // Presentation created successfully
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating presentation: {ex.Message}");
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
                    // Presentation opened successfully
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error opening presentation: {ex.Message}");
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
                    // Presentation saved successfully
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No active presentation to save.");
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
                        // Presentation saved successfully
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No active presentation to save.");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving presentation: {ex.Message}");
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
                var app = Globals.ThisAddIn.Application;
                var slideForElement = GetActiveSlideOrNull();
                if (slideForElement != null)
                {
                    // Show element selection options
                    var elementDialog = new Form();
                    elementDialog.Text = "Smart Element Wizard";
                    elementDialog.Size = new Size(400, 300);
                    elementDialog.StartPosition = FormStartPosition.CenterParent;
                    
                    var listBox = new ListBox();
                    listBox.Size = new Size(350, 200);
                    listBox.Location = new Point(25, 25);
                    listBox.Items.AddRange(new string[] {
                        "📊 Process Flow (SmartArt)",
                        "🔄 Cycle Diagram",
                        "📈 Hierarchy Chart",
                        "📋 List with Icons",
                        "⚡ Decision Tree",
                        "🎯 Target Diagram",
                        "📐 Matrix Layout",
                        "🌟 Feature Highlight"
                    });
                    
                    var btnOK = new Button();
                    btnOK.Text = "Create Element";
                    btnOK.Size = new Size(100, 30);
                    btnOK.Location = new Point(200, 240);
                    btnOK.DialogResult = DialogResult.OK;
                    
                    var btnCancel = new Button();
                    btnCancel.Text = "Cancel";
                    btnCancel.Size = new Size(80, 30);
                    btnCancel.Location = new Point(310, 240);
                    btnCancel.DialogResult = DialogResult.Cancel;
                    
                    elementDialog.Controls.AddRange(new Control[] { listBox, btnOK, btnCancel });
                    elementDialog.AcceptButton = btnOK;
                    elementDialog.CancelButton = btnCancel;
                    
                    if (elementDialog.ShowDialog() == DialogResult.OK && listBox.SelectedIndex >= 0)
                    {
                        var slide = GetActiveSlideOrNull();
                        if (slide != null)
                        {
                            CreateSmartElement(slide, listBox.SelectedIndex);
                        }
                    }
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating element: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateSmartElement(PowerPoint.Slide slide, int elementType)
        {
            try
            {
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;
                float centerX = slideWidth / 2;
                float centerY = slideHeight / 2;
                
                                 switch (elementType)
                 {
                     case 0: // Process Flow
                         CreateProcessFlowDiagram(slide, centerX - 200, centerY - 100);
                         break;
                     case 1: // Cycle Diagram
                         CreateCycleDiagram(slide, centerX - 150, centerY - 150);
                         break;
                     case 2: // Hierarchy Chart
                         CreateHierarchyChart(slide, centerX - 200, centerY - 150);
                         break;
                    case 3: // List with Icons
                        CreateIconList(slide, centerX - 150, centerY - 100);
                        break;
                    case 4: // Decision Tree
                        CreateDecisionTree(slide, centerX - 200, centerY - 150);
                        break;
                    case 5: // Target Diagram
                        CreateTargetDiagram(slide, centerX - 100, centerY - 100);
                        break;
                    case 6: // Matrix Layout
                        CreateMatrixLayout(slide, centerX - 150, centerY - 100);
                        break;
                    case 7: // Feature Highlight
                        CreateFeatureHighlight(slide, centerX - 150, centerY - 75);
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create smart element: {ex.Message}");
            }
        }

        private void CreateIconList(PowerPoint.Slide slide, float left, float top)
        {
            string[] items = { "First Item", "Second Item", "Third Item" };
            string[] icons = { "🔹", "🔸", "🔹" };
            
            for (int i = 0; i < items.Length; i++)
            {
                var textBox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 
                    left, top + (i * 40), 300, 30);
                textBox.TextFrame.TextRange.Text = $"{icons[i]} {items[i]}";
                textBox.TextFrame.TextRange.Font.Size = 16;
            }
        }

        private void CreateDecisionTree(PowerPoint.Slide slide, float left, float top)
        {
            // Create decision tree with shapes and connectors
            var rootBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left + 150, top, 100, 50);
            rootBox.TextFrame.TextRange.Text = "Decision?";
            rootBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightBlue);
            
            var yesBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top + 100, 100, 50);
            yesBox.TextFrame.TextRange.Text = "Yes";
            yesBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGreen);
            
            var noBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left + 300, top + 100, 100, 50);
            noBox.TextFrame.TextRange.Text = "No";
            noBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightCoral);
        }

        private void CreateTargetDiagram(PowerPoint.Slide slide, float left, float top)
        {
            // Create concentric circles for target diagram
            var outerCircle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, top, 200, 200);
            outerCircle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
            outerCircle.TextFrame.TextRange.Text = "Goal";
            
            var innerCircle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left + 50, top + 50, 100, 100);
            innerCircle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Yellow);
            innerCircle.TextFrame.TextRange.Text = "Target";
        }

        private void CreateMatrixLayout(PowerPoint.Slide slide, float left, float top)
        {
            // Create 2x2 matrix layout
            string[] labels = { "Quadrant 1", "Quadrant 2", "Quadrant 3", "Quadrant 4" };
            Color[] colors = { Color.LightBlue, Color.LightGreen, Color.LightYellow, Color.LightCoral };
            
            for (int i = 0; i < 4; i++)
            {
                int row = i / 2;
                int col = i % 2;
                var box = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                    left + (col * 150), top + (row * 100), 140, 90);
                box.TextFrame.TextRange.Text = labels[i];
                box.Fill.ForeColor.RGB = ColorTranslator.ToOle(colors[i]);
            }
        }

        private void CreateFeatureHighlight(PowerPoint.Slide slide, float left, float top)
        {
            // Create feature highlight box
            var highlightBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, 300, 150);
            highlightBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(255, 255, 102));
            highlightBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Orange);
            highlightBox.Line.Weight = 3;
            highlightBox.TextFrame.TextRange.Text = "⭐ Key Feature\n\nHighlight important information here";
            highlightBox.TextFrame.TextRange.Font.Size = 14;
            highlightBox.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
        }

        private void CreateProcessFlowDiagram(PowerPoint.Slide slide, float left, float top)
        {
            // Create process flow with connected boxes and arrows
            string[] steps = { "Start", "Process", "Decision", "End" };
            
            for (int i = 0; i < steps.Length; i++)
            {
                Office.MsoAutoShapeType shapeType = Office.MsoAutoShapeType.msoShapeRectangle;
                if (i == 0) shapeType = Office.MsoAutoShapeType.msoShapeOval; // Start
                if (i == 2) shapeType = Office.MsoAutoShapeType.msoShapeDiamond; // Decision
                if (i == 3) shapeType = Office.MsoAutoShapeType.msoShapeOval; // End
                
                var step = slide.Shapes.AddShape(shapeType, left + (i * 100), top, 80, 60);
                step.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightBlue);
                step.TextFrame.TextRange.Text = steps[i];
                step.TextFrame.TextRange.Font.Size = 12;
                
                // Add connecting arrows
                if (i < steps.Length - 1)
                {
                    var arrow = slide.Shapes.AddConnector(Office.MsoConnectorType.msoConnectorStraight, 
                        left + (i * 100) + 80, top + 30, left + ((i + 1) * 100), top + 30);
                    arrow.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadTriangle;
                }
            }
        }

        private void CreateCycleDiagram(PowerPoint.Slide slide, float left, float top)
        {
            // Create circular process diagram
            string[] phases = { "Phase 1", "Phase 2", "Phase 3", "Phase 4" };
            float radius = 120f;
            float centerX = left + 150;
            float centerY = top + 150;
            
            for (int i = 0; i < phases.Length; i++)
            {
                double angle = (2 * Math.PI * i) / phases.Length;
                float x = centerX + (float)(radius * Math.Cos(angle)) - 40;
                float y = centerY + (float)(radius * Math.Sin(angle)) - 20;
                
                var phase = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, x, y, 80, 40);
                phase.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGreen);
                phase.TextFrame.TextRange.Text = phases[i];
                phase.TextFrame.TextRange.Font.Size = 10;
            }
            
            // Add center circle
            var center = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, centerX - 30, centerY - 30, 60, 60);
            center.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Yellow);
            center.TextFrame.TextRange.Text = "Cycle";
        }

        private void CreateHierarchyChart(PowerPoint.Slide slide, float left, float top)
        {
            // Create organizational chart structure
            // Top level
            var ceo = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left + 150, top, 100, 50);
            ceo.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Gold);
            ceo.TextFrame.TextRange.Text = "CEO";
            ceo.TextFrame.TextRange.Font.Size = 12;
            ceo.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            
            // Second level
            string[] managers = { "Manager A", "Manager B", "Manager C" };
            for (int i = 0; i < managers.Length; i++)
            {
                var manager = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                    left + (i * 130), top + 100, 100, 40);
                manager.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightBlue);
                manager.TextFrame.TextRange.Text = managers[i];
                manager.TextFrame.TextRange.Font.Size = 10;
                
                // Connect to CEO
                var connector = slide.Shapes.AddConnector(Office.MsoConnectorType.msoConnectorStraight,
                    left + 200, top + 50, left + (i * 130) + 50, top + 100);
                connector.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
            }
            
            // Third level (under first manager only)
            for (int i = 0; i < 2; i++)
            {
                var employee = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                    left + (i * 60), top + 200, 80, 30);
                employee.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
                employee.TextFrame.TextRange.Text = $"Employee {i + 1}";
                employee.TextFrame.TextRange.Font.Size = 9;
                
                // Connect to manager
                var connector = slide.Shapes.AddConnector(Office.MsoConnectorType.msoConnectorStraight,
                    left + 50, top + 140, left + (i * 60) + 40, top + 200);
                connector.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
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
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    // Show format options dialog
                    var formatDialog = new Form();
                    formatDialog.Text = "Format Wizard";
                    formatDialog.Size = new Size(450, 350);
                    formatDialog.StartPosition = FormStartPosition.CenterParent;
                    
                    var label = new Label();
                    label.Text = "Select formatting action:";
                    label.Size = new Size(400, 20);
                    label.Location = new Point(25, 25);
                    
                    var listBox = new ListBox();
                    listBox.Size = new Size(400, 220);
                    listBox.Location = new Point(25, 50);
                    listBox.Items.AddRange(new string[] {
                        "🎨 Apply Corporate Color Scheme",
                        "📝 Standardize All Text Fonts",
                        "📐 Align All Objects to Grid",
                        "🔲 Apply Consistent Shape Styles",
                        "📊 Format All Charts Uniformly",
                        "📋 Standardize Table Formatting",
                        "🎯 Quick Professional Theme",
                        "🌈 Color Harmony Correction",
                        "📏 Consistent Spacing & Margins",
                        "✨ Add Drop Shadows to All Shapes"
                    });
                    
                    var btnOK = new Button();
                    btnOK.Text = "Apply Format";
                    btnOK.Size = new Size(100, 30);
                    btnOK.Location = new Point(250, 280);
                    btnOK.DialogResult = DialogResult.OK;
                    
                    var btnCancel = new Button();
                    btnCancel.Text = "Cancel";
                    btnCancel.Size = new Size(80, 30);
                    btnCancel.Location = new Point(360, 280);
                    btnCancel.DialogResult = DialogResult.Cancel;
                    
                    formatDialog.Controls.AddRange(new Control[] { label, listBox, btnOK, btnCancel });
                    formatDialog.AcceptButton = btnOK;
                    formatDialog.CancelButton = btnCancel;
                    
                    if (formatDialog.ShowDialog() == DialogResult.OK && listBox.SelectedIndex >= 0)
                    {
                        ApplyFormatting(app.ActivePresentation, listBox.SelectedIndex);
                    }
                }
                else
                {
                    MessageBox.Show("Please open a presentation first to apply formatting.", "Format Wizard", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying formatting: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyFormatting(PowerPoint.Presentation presentation, int formatType)
        {
            try
            {
                switch (formatType)
                {
                    case 0: // Corporate Color Scheme
                        ApplyCorporateColors(presentation);
                        break;
                    case 1: // Standardize Fonts
                        StandardizeFonts(presentation);
                        break;
                    case 2: // Align to Grid
                        AlignObjectsToGrid(presentation);
                        break;
                    case 3: // Shape Styles
                        ApplyConsistentShapeStyles(presentation);
                        break;
                    case 4: // Chart Formatting
                        FormatAllCharts(presentation);
                        break;
                    case 5: // Table Formatting
                        StandardizeTableFormatting(presentation);
                        break;
                    case 6: // Professional Theme
                        ApplyProfessionalTheme(presentation);
                        break;
                    case 7: // Color Harmony
                        CorrectColorHarmony(presentation);
                        break;
                    case 8: // Spacing & Margins
                        FixSpacingAndMargins(presentation);
                        break;
                    case 9: // Drop Shadows
                        AddDropShadowsToShapes(presentation);
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to apply formatting: {ex.Message}");
            }
        }

        private void ApplyCorporateColors(PowerPoint.Presentation presentation)
        {
            // Apply a professional blue-based color scheme
            Color primaryColor = Color.FromArgb(0, 102, 204);   // Professional blue
            
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        shape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(51, 51, 51));
                    }
                    
                    if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                    {
                        shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(primaryColor);
                    }
                }
            }
        }

        private void StandardizeFonts(PowerPoint.Presentation presentation)
        {
            string headerFont = "Calibri";
            string bodyFont = "Calibri";
            
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        var textRange = shape.TextFrame.TextRange;
                        textRange.Font.Name = bodyFont;
                        
                        // Make title shapes larger
                        if (shape.Type == Office.MsoShapeType.msoPlaceholder && 
                            shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderTitle)
                        {
                            textRange.Font.Name = headerFont;
                            textRange.Font.Size = 24;
                            textRange.Font.Bold = Office.MsoTriState.msoTrue;
                        }
                    }
                }
            }
        }

        private void AlignObjectsToGrid(PowerPoint.Presentation presentation)
        {
            float gridSize = 20f; // 20 point grid
            
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    // Snap to grid
                    shape.Left = (float)(Math.Round(shape.Left / gridSize) * gridSize);
                    shape.Top = (float)(Math.Round(shape.Top / gridSize) * gridSize);
                }
            }
        }

        private void ApplyConsistentShapeStyles(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == Office.MsoShapeType.msoAutoShape)
                    {
                        shape.Line.Weight = 1.5f;
                        shape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
                        shape.Fill.Transparency = 0.1f;
                    }
                }
            }
        }

        private void FormatAllCharts(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.HasChart == Office.MsoTriState.msoTrue)
                    {
                        // Apply consistent chart formatting
                        shape.Chart.ChartStyle = 42; // Professional style
                    }
                }
            }
        }

        private void StandardizeTableFormatting(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        var table = shape.Table;
                        
                        // Apply consistent table styling
                        for (int row = 1; row <= table.Rows.Count; row++)
                        {
                            for (int col = 1; col <= table.Columns.Count; col++)
                            {
                                var cell = table.Cell(row, col);
                                cell.Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
                                cell.Shape.Line.Weight = 1;
                                cell.Shape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
                            }
                        }
                    }
                }
            }
        }

        private void ApplyProfessionalTheme(PowerPoint.Presentation presentation)
        {
            // Apply a built-in professional theme
            try
            {
                // This would apply a built-in theme - simplified version
                ApplyCorporateColors(presentation);
                StandardizeFonts(presentation);
                AlignObjectsToGrid(presentation);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Theme application failed: {ex.Message}");
            }
        }

        private void CorrectColorHarmony(PowerPoint.Presentation presentation)
        {
            // Apply color harmony rules
            Color[] harmonicColors = {
                Color.FromArgb(0, 102, 204),    // Blue
                Color.FromArgb(102, 153, 255),  // Light Blue
                Color.FromArgb(153, 204, 255),  // Lighter Blue
                Color.FromArgb(255, 165, 0),    // Orange (complementary)
                Color.FromArgb(255, 215, 0)     // Gold (triadic)
            };
            
            int colorIndex = 0;
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                    {
                        shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(harmonicColors[colorIndex % harmonicColors.Length]);
                        colorIndex++;
                    }
                }
            }
        }

        private void FixSpacingAndMargins(PowerPoint.Presentation presentation)
        {
            float margin = 50f;
            
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                var shapes = slide.Shapes.Cast<PowerPoint.Shape>().ToList();
                
                for (int i = 0; i < shapes.Count; i++)
                {
                    var shape = shapes[i];
                    
                    // Ensure minimum margins
                    if (shape.Left < margin) shape.Left = margin;
                    if (shape.Top < margin) shape.Top = margin;
                }
            }
        }

        private void AddDropShadowsToShapes(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == Office.MsoShapeType.msoAutoShape || 
                        shape.Type == Office.MsoShapeType.msoTextBox)
                    {
                        try
                        {
                            shape.Shadow.Type = Office.MsoShadowType.msoShadow6;
                            shape.Shadow.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
                            shape.Shadow.Transparency = 0.5f;
                        }
                        catch
                        {
                            // Some shapes might not support shadows
                        }
                    }
                }
            }
        }

        private void BtnMap_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var slide = GetActiveSlideOrNull();
                if (slide != null)
                {
                    // Show map options dialog
                    var mapDialog = new Form();
                    mapDialog.Text = "Map Wizard";
                    mapDialog.Size = new Size(400, 320);
                    mapDialog.StartPosition = FormStartPosition.CenterParent;
                    
                    var label = new Label();
                    label.Text = "Select map type to insert:";
                    label.Size = new Size(350, 20);
                    label.Location = new Point(25, 25);
                    
                    var listBox = new ListBox();
                    listBox.Size = new Size(350, 200);
                    listBox.Location = new Point(25, 50);
                    listBox.Items.AddRange(new string[] {
                        "🌍 World Map Outline",
                        "🇺🇸 USA Map with States",
                        "🇪🇺 Europe Map",
                        "📍 Location Pin Template",
                        "🗺️ Process Journey Map",
                        "🏢 Office Floor Plan Template",
                        "🌆 City Skyline Template",
                        "🔴 Hotspot Map Template",
                        "📊 Geographic Data Visualization"
                    });
                    
                    var btnOK = new Button();
                    btnOK.Text = "Insert Map";
                    btnOK.Size = new Size(100, 30);
                    btnOK.Location = new Point(200, 260);
                    btnOK.DialogResult = DialogResult.OK;
                    
                    var btnCancel = new Button();
                    btnCancel.Text = "Cancel";
                    btnCancel.Size = new Size(80, 30);
                    btnCancel.Location = new Point(290, 260);
                    btnCancel.DialogResult = DialogResult.Cancel;
                    
                    mapDialog.Controls.AddRange(new Control[] { label, listBox, btnOK, btnCancel });
                    mapDialog.AcceptButton = btnOK;
                    mapDialog.CancelButton = btnCancel;
                    
                    if (mapDialog.ShowDialog() == DialogResult.OK && listBox.SelectedIndex >= 0)
                    {
                        var targetSlide = GetActiveSlideOrNull();
                        if (targetSlide != null)
                        {
                            CreateMapTemplate(targetSlide, listBox.SelectedIndex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating map: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateMapTemplate(PowerPoint.Slide slide, int mapType)
        {
            try
            {
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;
                float centerX = slideWidth / 2;
                float centerY = slideHeight / 2;
                
                switch (mapType)
                {
                    case 0: // World Map Outline
                        CreateWorldMapOutline(slide, centerX - 200, centerY - 150);
                        break;
                    case 1: // USA Map
                        CreateUSAMap(slide, centerX - 150, centerY - 100);
                        break;
                    case 2: // Europe Map
                        CreateEuropeMap(slide, centerX - 125, centerY - 100);
                        break;
                    case 3: // Location Pin
                        CreateLocationPinTemplate(slide, centerX - 100, centerY - 100);
                        break;
                    case 4: // Process Journey
                        CreateProcessJourneyMap(slide, centerX - 200, centerY - 100);
                        break;
                    case 5: // Floor Plan
                        CreateFloorPlanTemplate(slide, centerX - 150, centerY - 100);
                        break;
                    case 6: // City Skyline
                        CreateCitySkylineTemplate(slide, centerX - 200, centerY - 75);
                        break;
                    case 7: // Hotspot Map
                        CreateHotspotMap(slide, centerX - 150, centerY - 100);
                        break;
                    case 8: // Geographic Data Viz
                        CreateGeographicDataVisualization(slide, centerX - 175, centerY - 125);
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create map template: {ex.Message}");
            }
        }

        private void CreateWorldMapOutline(PowerPoint.Slide slide, float left, float top)
        {
            // Simplified world map using basic shapes
            var worldShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, 400, 300);
            worldShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightBlue);
            worldShape.Line.Weight = 2;
            worldShape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.DarkBlue);
            
            // Add continent shapes (simplified)
            var continent1 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left + 50, top + 80, 80, 60);
            continent1.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Green);
            continent1.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.DarkGreen);
            
            var continent2 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left + 200, top + 100, 120, 80);
            continent2.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Green);
            continent2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.DarkGreen);
            
            // Add title
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top - 30, 400, 25);
            title.TextFrame.TextRange.Text = "World Map";
            title.TextFrame.TextRange.Font.Size = 18;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
        }

        private void CreateUSAMap(PowerPoint.Slide slide, float left, float top)
        {
            // Simplified USA outline
            var usaShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 300, 200);
            usaShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(176, 224, 230));
            usaShape.Line.Weight = 2;
            usaShape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Navy);
            
            // Add state labels
            var label1 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left + 50, top + 50, 60, 20);
            label1.TextFrame.TextRange.Text = "CA";
            label1.TextFrame.TextRange.Font.Size = 12;
            
            var label2 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left + 200, top + 80, 60, 20);
            label2.TextFrame.TextRange.Text = "NY";
            label2.TextFrame.TextRange.Font.Size = 12;
        }

        private void CreateEuropeMap(PowerPoint.Slide slide, float left, float top)
        {
            // Simplified Europe outline
            var europeShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 250, 200);
            europeShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(255, 228, 181));
            europeShape.Line.Weight = 2;
            europeShape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Brown);
            
            // Add country markers
            CreateLocationPin(slide, left + 100, top + 80, "🇩🇪 Germany");
            CreateLocationPin(slide, left + 50, top + 60, "🇫🇷 France");
            CreateLocationPin(slide, left + 150, top + 100, "🇮🇹 Italy");
        }

        private void CreateLocationPinTemplate(PowerPoint.Slide slide, float left, float top)
        {
            // Create location pin with callout
            var pin = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left + 90, top + 150, 20, 20);
            pin.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            
            var pinTop = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeIsoscelesTriangle, left + 95, top + 130, 10, 20);
            pinTop.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            pinTop.Rotation = 180;
            
            // Add callout
            var callout = slide.Shapes.AddCallout(Office.MsoCalloutType.msoCalloutTwo, left + 120, top + 100, 150, 60);
            callout.TextFrame.TextRange.Text = "📍 Location Name\nDescription here";
            callout.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Yellow);
        }

        private void CreateLocationPin(PowerPoint.Slide slide, float left, float top, string label)
        {
            var pin = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, top, 12, 12);
            pin.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            
            var labelBox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left + 15, top - 5, 80, 20);
            labelBox.TextFrame.TextRange.Text = label;
            labelBox.TextFrame.TextRange.Font.Size = 10;
        }

        private void CreateProcessJourneyMap(PowerPoint.Slide slide, float left, float top)
        {
            // Create journey steps with connecting arrows
            string[] steps = { "Start", "Step 1", "Step 2", "End" };
            
            for (int i = 0; i < steps.Length; i++)
            {
                var step = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, 
                    left + (i * 100), top, 80, 40);
                step.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGreen);
                step.TextFrame.TextRange.Text = steps[i];
                
                if (i < steps.Length - 1)
                {
                    var arrow = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightArrow, 
                        left + (i * 100) + 85, top + 15, 10, 10);
                    arrow.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Blue);
                }
            }
        }

        private void CreateFloorPlanTemplate(PowerPoint.Slide slide, float left, float top)
        {
            // Create simple floor plan outline
            var building = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 300, 200);
            building.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
            building.Line.Weight = 2;
            
            // Add rooms
            var room1 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left + 20, top + 20, 80, 60);
            room1.Fill.Transparency = 0.5f;
            room1.TextFrame.TextRange.Text = "Office 1";
            
            var room2 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left + 120, top + 20, 80, 60);
            room2.Fill.Transparency = 0.5f;
            room2.TextFrame.TextRange.Text = "Office 2";
            
            var corridor = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left + 20, top + 100, 180, 40);
            corridor.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            corridor.TextFrame.TextRange.Text = "Corridor";
        }

        private void CreateCitySkylineTemplate(PowerPoint.Slide slide, float left, float top)
        {
            // Create city skyline with buildings of different heights
            int[] buildingHeights = { 80, 120, 100, 150, 90, 110, 130 };
            
            for (int i = 0; i < buildingHeights.Length; i++)
            {
                var building = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                    left + (i * 40), top + (150 - buildingHeights[i]), 35, buildingHeights[i]);
                building.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(105, 105, 105));
                building.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
                
                // Add windows
                for (int w = 0; w < 3; w++)
                {
                    var window = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                        left + (i * 40) + 5 + (w * 8), top + (150 - buildingHeights[i]) + 10, 6, 8);
                    window.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Yellow);
                }
            }
        }

        private void CreateHotspotMap(PowerPoint.Slide slide, float left, float top)
        {
            // Create base map
            var baseMap = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 300, 200);
            baseMap.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
            
            // Add hotspots with different intensities
            var hotspot1 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left + 50, top + 50, 30, 30);
            hotspot1.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            hotspot1.Fill.Transparency = 0.3f;
            
            var hotspot2 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left + 150, top + 80, 40, 40);
            hotspot2.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Orange);
            hotspot2.Fill.Transparency = 0.3f;
            
            var hotspot3 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left + 200, top + 120, 25, 25);
            hotspot3.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Yellow);
            hotspot3.Fill.Transparency = 0.3f;
        }

        private void CreateGeographicDataVisualization(PowerPoint.Slide slide, float left, float top)
        {
            // Create choropleth-style map with data regions
            var baseRegion = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 350, 250);
            baseRegion.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightBlue);
            
            // Add data regions with different colors representing data values
            Color[] dataColors = { Color.Red, Color.Orange, Color.Yellow, Color.LightGreen, Color.Green };
            string[] dataLabels = { "High", "Med-High", "Medium", "Med-Low", "Low" };
            
            for (int i = 0; i < 5; i++)
            {
                var region = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                    left + (i * 60) + 20, top + 50, 50, 60);
                region.Fill.ForeColor.RGB = ColorTranslator.ToOle(dataColors[i]);
                region.TextFrame.TextRange.Text = dataLabels[i];
                region.TextFrame.TextRange.Font.Size = 10;
            }
            
            // Add legend
            var legend = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 
                left + 280, top + 180, 100, 80);
            legend.TextFrame.TextRange.Text = "Legend:\n🔴 High\n🟠 Med-High\n🟡 Medium\n🟢 Low";
            legend.TextFrame.TextRange.Font.Size = 10;
        }

        #endregion

        #region Smart Elements Section

        private void BtnChart_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var slide = GetActiveSlideOrNull();
                if (slide != null)
                {
                    slide.Shapes.AddChart2(Style: -1, Type: Office.XlChartType.xlColumnClustered);
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
                var slide = GetActiveSlideOrNull();
                if (slide != null)
                {
                    // Show table dropdown control
                    var tableDropdown = new TableDropdownControl();
                    
                    // Position the dropdown below the table button
                    var btnLocation = btnTable.PointToScreen(Point.Empty);
                    tableDropdown.Location = new Point(btnLocation.X, btnLocation.Y + btnTable.Height);
                    
                    if (tableDropdown.ShowDialog() == DialogResult.OK)
                    {
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
                var slide = GetActiveSlideOrNull();
                if (slide != null)
                {
                    // Get user input for rows and columns
                    var matrixDialog = new MatrixTableDialog();
                    if (matrixDialog.ShowDialog() == DialogResult.OK)
                    {
                        int rows = matrixDialog.Rows;
                        int columns = matrixDialog.Columns;
                        bool hasHeader = matrixDialog.HasHeader;
                        
                        var targetSlide = GetActiveSlideOrNull();
                        if (targetSlide != null)
                        {
                            CreateMatrixTable(targetSlide, rows, columns, hasHeader);
                        }
                    }
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
                CreateCustomMatrix(slide, rows, columns);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating matrix table: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CreateCustomMatrix(PowerPoint.Slide slide, int rows, int columns)
        {
            try
            {
                // Slide size
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;

                // Define consistent spacing between cells
                float cellSpacing = 8f;

                // Grid size with margin
                float availableWidth = slideWidth * 0.85f;
                float availableHeight = slideHeight * 0.75f;

                // Total spacing needed
                float totalHorizontalSpacing = (columns - 1) * cellSpacing;
                float totalVerticalSpacing = (rows - 1) * cellSpacing;

                // Available space for actual cells
                float cellAreaWidth = availableWidth - totalHorizontalSpacing;
                float cellAreaHeight = availableHeight - totalVerticalSpacing;

                // Cell size
                float maxCellWidth = cellAreaWidth / columns;
                float maxCellHeight = cellAreaHeight / rows;
                float cellSize = Math.Min(maxCellWidth, maxCellHeight);
                cellSize = Math.Max(cellSize, 30f); // Enforce minimum readable size

                // Actual grid size
                float actualGridWidth = (columns * cellSize) + totalHorizontalSpacing;
                float actualGridHeight = (rows * cellSize) + totalVerticalSpacing;

                // Grid start position
                float startLeft = (slideWidth - actualGridWidth) / 2;
                float startTop = (slideHeight - actualGridHeight) / 2;

                // Track all cells in row-major order for later selection
                List<PowerPoint.Shape> allCells = new List<PowerPoint.Shape>();

                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < columns; col++)
                    {
                        float left = startLeft + (col * (cellSize + cellSpacing));
                        float top = startTop + (row * (cellSize + cellSpacing));

                        // Create individual shape (text-enabled rectangle)
                        PowerPoint.Shape cell = slide.Shapes.AddShape(
                            Type: Office.MsoAutoShapeType.msoShapeRectangle,
                            Left: left,
                            Top: top,
                            Width: cellSize,
                            Height: cellSize);

                        // Default text
                        cell.TextFrame.TextRange.Text = "XXXX";
                        cell.TextFrame.TextRange.Font.Size = Math.Max(8f, cellSize / 4f);
                        cell.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        cell.TextFrame.TextRange.Font.Name = "Segoe UI";

                        // Align text
                        cell.TextFrame.HorizontalAnchor = Office.MsoHorizontalAnchor.msoAnchorCenter;
                        cell.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                        cell.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                        // Remove text margins for tight layout
                        cell.TextFrame.MarginBottom = 0f;
                        cell.TextFrame.MarginTop = 0f;
                        cell.TextFrame.MarginLeft = 0f;
                        cell.TextFrame.MarginRight = 0f;

                        // Border
                        cell.Line.Visible = Office.MsoTriState.msoTrue;
                        cell.Line.Weight = 1.0f;
                        cell.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

                        // Background
                        cell.Fill.Visible = Office.MsoTriState.msoTrue;
                        cell.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                        // Track cell
                        allCells.Add(cell);

                        try
                        {
                            // Tag as matrix cell with row/col for paste handling
                            cell.Tags.Add("MATRIX", "1");
                            cell.Tags.Add("MATRIX_ROW", row.ToString());
                            cell.Tags.Add("MATRIX_COL", col.ToString());
                        }
                        catch { /* ignore tag failures */ }
                    }
                }

                // Select all shapes in row-major order (important for Excel paste)
                string[] shapeNames = allCells.Select(s => s.Name).ToArray();
                PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(shapeNames);
                shapeRange.Select();
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
                var slideForSticky = GetActiveSlideOrNull();
                if (slideForSticky != null)
                {
                    // Get user input for sticky note
                    var stickyDialog = new StickyNoteDialog();
                    if (stickyDialog.ShowDialog() == DialogResult.OK)
                    {
                        string noteText = stickyDialog.NoteText;
                        Color noteColor = stickyDialog.NoteColor;
                        
                        var slide = GetActiveSlideOrNull();
                        if (slide != null)
                        {
                            CreateStickyNote(slide, noteText, noteColor);
                        }
                    }
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
                var slide = GetActiveSlideOrNull();
                if (slide != null)
                {
                    // Simple input dialog for citation text
                    string citationText = Microsoft.VisualBasic.Interaction.InputBox(
                        "Enter citation text:",
                        "Add Citation",
                        "Source: [Author, Year, Title]");
                    
                    if (!string.IsNullOrEmpty(citationText))
                    {
                        var targetSlide = GetActiveSlideOrNull();
                        if (targetSlide != null)
                        {
                            CreateCitation(targetSlide, citationText);
                        }
                    }
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
                var slide = GetActiveSlideOrNull();
                if (slide != null)
                {
                    // Show standard objects dialog
                    var objectsDialog = new StandardObjectsDialog();
                    if (objectsDialog.ShowDialog() == DialogResult.OK)
                    {
                        string selectedObject = objectsDialog.SelectedObject;
                        
                        if (!string.IsNullOrEmpty(selectedObject))
                        {
                            var targetSlide = GetActiveSlideOrNull();
                            if (targetSlide != null)
                            {
                                CreateStandardObject(targetSlide, selectedObject);
                            }
                        }
                    }
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
                    }
                }
            }
            catch (Exception)
            {
                // Silently handle errors to avoid popup dialogs
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
                    }
                }
            }
            catch (Exception)
            {
                // Silently handle errors to avoid popup dialogs
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
                    }
                }
            }
            catch (Exception)
            {
                // Silently handle errors to avoid popup dialogs
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
                        // Swap locations based on center points
                        float shape1CenterX = shapes[1].Left + shapes[1].Width / 2f;
                        float shape1CenterY = shapes[1].Top + shapes[1].Height / 2f;
                        float shape2CenterX = shapes[2].Left + shapes[2].Width / 2f;
                        float shape2CenterY = shapes[2].Top + shapes[2].Height / 2f;

                        // Move shape 1's center to shape 2's center
                        shapes[1].Left = shape2CenterX - shapes[1].Width / 2f;
                        shapes[1].Top = shape2CenterY - shapes[1].Height / 2f;

                        // Move shape 2's center to shape 1's center
                        shapes[2].Left = shape1CenterX - shapes[2].Width / 2f;
                        shapes[2].Top = shape1CenterY - shapes[2].Height / 2f;
                    }
                }
            }
            catch (Exception)
            {
                // Silently handle errors to avoid popup dialogs
            }
        }

        #endregion

        #region Size Section

        // Smart size presets for different use cases
        private readonly Dictionary<string, (decimal width, decimal height)> sizePresets = new Dictionary<string, (decimal, decimal)>
        {
            { "4:3 Standard", (10m, 7.5m) },
            { "16:9 Widescreen", (13.3m, 7.5m) },
            { "16:10 Widescreen", (12.8m, 8m) },
            { "A4 Portrait", (8.27m, 11.69m) },
            { "A4 Landscape", (11.69m, 8.27m) },
            { "Letter Portrait", (8.5m, 11m) },
            { "Letter Landscape", (11m, 8.5m) },
            { "A3 Portrait", (11.69m, 16.54m) },
            { "A3 Landscape", (16.54m, 11.69m) },
            { "Banner", (8m, 1m) },
            { "Social Media", (1.91m, 1m) },
            { "Instagram Post", (1m, 1m) },
            { "Instagram Story", (1m, 1.78m) },
            { "YouTube Thumbnail", (1.78m, 1m) },
            { "LinkedIn Post", (1.91m, 1m) },
            { "Twitter Post", (1.91m, 1m) },
            { "Facebook Post", (1.91m, 1m) }
        };

        private void CmbSlideSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (nudWidth != null && nudHeight != null && cmbSlideSize.SelectedItem != null)
            {
                string selectedSize = cmbSlideSize.SelectedItem.ToString();
                if (sizePresets.ContainsKey(selectedSize))
                {
                    var (width, height) = sizePresets[selectedSize];
                    nudWidth.Value = width;
                    nudHeight.Value = height;
                    
                    // Auto-apply if it's a preset (not custom)
                    if (selectedSize != "Custom")
                    {
                        ApplySlideSizeWithSmartScaling(width, height, selectedSize);
                    }
                }
            }
        }

        private void BtnApplySize_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app != null && app.ActivePresentation != null && nudWidth != null && nudHeight != null)
                {
                    decimal width = nudWidth.Value;
                    decimal height = nudHeight.Value;
                    string sizeName = cmbSlideSize.SelectedItem?.ToString() ?? "Custom";
                    
                    ApplySlideSizeWithSmartScaling(width, height, sizeName);
                }
                else if (app == null)
                {
                    MessageBox.Show("PowerPoint is not available. Please ensure PowerPoint is running.", "PowerPoint Not Available", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show("Please open or create a presentation first.", "No Active Presentation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying slide size: {ex.Message}");
            }
        }

        /// <summary>
        /// Applies slide size with intelligent content scaling
        /// </summary>
        private void ApplySlideSizeWithSmartScaling(decimal width, decimal height, string sizeName)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app == null)
                {
                    MessageBox.Show("PowerPoint application is not available.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var presentation = app.ActivePresentation;
                if (presentation == null)
                {
                    MessageBox.Show("No active presentation found. Please open or create a presentation first.", "No Presentation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // Store current dimensions for scaling calculations
                float oldWidth = presentation.PageSetup.SlideWidth;
                float oldHeight = presentation.PageSetup.SlideHeight;
                
                // Apply new size
                presentation.PageSetup.SlideWidth = (float)width * 72; // Convert inches to points
                presentation.PageSetup.SlideHeight = (float)height * 72;
                
                // Calculate scaling factors
                float scaleX = presentation.PageSetup.SlideWidth / oldWidth;
                float scaleY = presentation.PageSetup.SlideHeight / oldHeight;
                
                // Apply intelligent content scaling
                ScaleContentIntelligently(scaleX, scaleY);
                
                // Update current slide size display
                UpdateCurrentSizeDisplay();
                
                System.Diagnostics.Debug.WriteLine($"Slide size changed to {sizeName} ({width}\" × {height}\")");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying smart slide size: {ex.Message}");
            }
        }

        /// <summary>
        /// Intelligently scales content when slide size changes
        /// </summary>
        private void ScaleContentIntelligently(float scaleX, float scaleY)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app?.ActivePresentation == null) return;
                
                var presentation = app.ActivePresentation;
                
                // Use the smaller scale factor to maintain aspect ratios
                float scaleFactor = Math.Min(scaleX, scaleY);
                
                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        // Scale position and size
                        shape.Left *= scaleFactor;
                        shape.Top *= scaleFactor;
                        shape.Width *= scaleFactor;
                        shape.Height *= scaleFactor;
                        
                        // Scale font size proportionally
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            try
                            {
                                var textRange = shape.TextFrame.TextRange;
                                if (textRange.Font.Size > 0)
                                {
                                    textRange.Font.Size *= scaleFactor;
                                }
                            }
                            catch
                            {
                                // Ignore font scaling errors for shapes without text
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error scaling content: {ex.Message}");
            }
        }

        /// <summary>
        /// Updates the current size display to show the actual slide dimensions
        /// </summary>
        private void UpdateCurrentSizeDisplay()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app != null && app.ActivePresentation != null)
                {
                    var presentation = app.ActivePresentation;
                    float widthInches = presentation.PageSetup.SlideWidth / 72f;
                    float heightInches = presentation.PageSetup.SlideHeight / 72f;
                    
                    // Update numeric controls to reflect current size
                    if (nudWidth != null && nudHeight != null)
                    {
                        nudWidth.Value = (decimal)Math.Round(widthInches, 2);
                        nudHeight.Value = (decimal)Math.Round(heightInches, 2);
                    }
                    
                    // Update combo box to show closest preset
                    UpdateComboBoxToClosestPreset(widthInches, heightInches);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No active presentation available for size display update");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating size display: {ex.Message}");
            }
        }

        /// <summary>
        /// Updates the combo box to show the closest matching preset
        /// </summary>
        private void UpdateComboBoxToClosestPreset(float width, float height)
        {
            try
            {
                if (cmbSlideSize == null) return;
                
                string closestPreset = "Custom";
                double minDifference = double.MaxValue;
                
                foreach (var preset in sizePresets)
                {
                    double diff = Math.Abs((double)preset.Value.width - (double)width) + Math.Abs((double)preset.Value.height - (double)height);
                    if (diff < minDifference)
                    {
                        minDifference = diff;
                        closestPreset = preset.Key;
                    }
                }
                
                // Only update if it's a close match (within 0.5 inches)
                if (minDifference < 0.5)
                {
                    cmbSlideSize.SelectedItem = closestPreset;
                }
                else
                {
                    cmbSlideSize.SelectedItem = "Custom";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating combo box: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads current slide size when task pane loads
        /// </summary>
        private void LoadCurrentSlideSize()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app?.ActivePresentation != null)
                {
                    float widthInches = app.ActivePresentation.PageSetup.SlideWidth / 72f;
                    float heightInches = app.ActivePresentation.PageSetup.SlideHeight / 72f;

                    if (nudWidth != null && nudHeight != null)
                    {
                        nudWidth.Value = (decimal)Math.Round(widthInches, 2);
                        nudHeight.Value = (decimal)Math.Round(heightInches, 2);
                    }

                    UpdateCurrentSizeDisplay();
                    System.Diagnostics.Debug.WriteLine($"Loaded slide size: {widthInches}\" x {heightInches}\"");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No active presentation - slide size not loaded");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading current slide size: {ex.Message}");
                // Silent failure - no user-facing errors
            }
        }

        /// <summary>
        /// Applies size to all slides in the presentation
        /// </summary>
        private void ApplySizeToAllSlides()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null && nudWidth != null && nudHeight != null)
                {
                    decimal width = nudWidth.Value;
                    decimal height = nudHeight.Value;
                    
                    var presentation = app.ActivePresentation;
                    presentation.PageSetup.SlideWidth = (float)width * 72;
                    presentation.PageSetup.SlideHeight = (float)height * 72;
                    
                    System.Diagnostics.Debug.WriteLine($"Size applied to all {presentation.Slides.Count} slides!");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying size to all slides: {ex.Message}");
            }
        }

        /// <summary>
        /// Suggests optimal slide size based on content analysis
        /// </summary>
        private void SuggestOptimalSize()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    var presentation = app.ActivePresentation;
                    
                    // Analyze content to suggest optimal size
                    string suggestion = AnalyzeContentAndSuggestSize(presentation);
                    
                    System.Diagnostics.Debug.WriteLine($"Size Suggestion: {suggestion}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error suggesting size: {ex.Message}");
            }
        }

        /// <summary>
        /// Analyzes presentation content and suggests optimal slide size
        /// </summary>
        private string AnalyzeContentAndSuggestSize(PowerPoint.Presentation presentation)
        {
            try
            {
                int totalSlides = presentation.Slides.Count;
                int slidesWithImages = 0;
                int slidesWithCharts = 0;
                int slidesWithTables = 0;
                int slidesWithText = 0;
                int slidesWithVideos = 0;
                int slidesWithSmartArt = 0;
                
                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    bool hasImages = false, hasCharts = false, hasTables = false, hasText = false, 
                         hasVideos = false, hasSmartArt = false;
                    
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.Type == Office.MsoShapeType.msoPicture || 
                            shape.Type == Office.MsoShapeType.msoLinkedPicture)
                            hasImages = true;
                        else if (shape.HasChart == Office.MsoTriState.msoTrue)
                            hasCharts = true;
                        else if (shape.HasTable == Office.MsoTriState.msoTrue)
                            hasTables = true;
                        else if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                            hasText = true;
                        else if (shape.Type == Office.MsoShapeType.msoMedia)
                            hasVideos = true;
                        else if (shape.Type == Office.MsoShapeType.msoSmartArt)
                            hasSmartArt = true;
                    }
                    
                    if (hasImages) slidesWithImages++;
                    if (hasCharts) slidesWithCharts++;
                    if (hasTables) slidesWithTables++;
                    if (hasText) slidesWithText++;
                    if (hasVideos) slidesWithVideos++;
                    if (hasSmartArt) slidesWithSmartArt++;
                }
                
                // Enhanced suggestion logic with more detailed analysis
                double imageRatio = (double)slidesWithImages / totalSlides;
                double chartRatio = (double)slidesWithCharts / totalSlides;
                double tableRatio = (double)slidesWithTables / totalSlides;
                double textRatio = (double)slidesWithText / totalSlides;
                double videoRatio = (double)slidesWithVideos / totalSlides;
                double smartArtRatio = (double)slidesWithSmartArt / totalSlides;
                
                // Priority-based suggestions
                if (videoRatio > 0.3)
                    return "16:9 Widescreen - Optimal for video content and modern displays";
                else if (imageRatio > 0.7)
                    return "16:9 Widescreen - Best for image-heavy presentations and visual storytelling";
                else if (chartRatio > 0.5)
                    return "16:9 Widescreen - Perfect for data visualization and charts";
                else if (smartArtRatio > 0.4)
                    return "16:9 Widescreen - Ideal for SmartArt and modern diagrams";
                else if (tableRatio > 0.6)
                    return "A4 Landscape - Excellent for table-heavy content and detailed data";
                else if (textRatio > 0.8)
                    return "4:3 Standard - Traditional format for text-heavy academic presentations";
                else if (totalSlides > 20)
                    return "16:9 Widescreen - Recommended for large presentations";
                else if (totalSlides < 5)
                    return "16:9 Widescreen - Perfect for short, impactful presentations";
                else
                    return "16:9 Widescreen - Modern standard for most presentations";
            }
            catch (Exception)
            {
                return "16:9 Widescreen - Recommended default size";
            }
        }

        /// <summary>
        /// Smart size optimization for different use cases
        /// </summary>
        private void OptimizeSizeForUseCase(string useCase)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    var presentation = app.ActivePresentation;
                    string suggestion = "";
                    
                    switch (useCase.ToLower())
                    {
                        case "presentation":
                            suggestion = "16:9 Widescreen - Standard for modern presentations";
                            ApplySlideSizeWithSmartScaling(13.3m, 7.5m, "16:9 Widescreen");
                            break;
                        case "print":
                            suggestion = "A4 Landscape - Optimized for printing";
                            ApplySlideSizeWithSmartScaling(11.69m, 8.27m, "A4 Landscape");
                            break;
                        case "social":
                            suggestion = "Social Media - Perfect for social platforms";
                            ApplySlideSizeWithSmartScaling(1.91m, 1m, "Social Media");
                            break;
                        case "mobile":
                            suggestion = "Instagram Story - Mobile-optimized";
                            ApplySlideSizeWithSmartScaling(1m, 1.78m, "Instagram Story");
                            break;
                        case "webinar":
                            suggestion = "16:9 Widescreen - Ideal for online presentations";
                            ApplySlideSizeWithSmartScaling(13.3m, 7.5m, "16:9 Widescreen");
                            break;
                        default:
                            suggestion = AnalyzeContentAndSuggestSize(presentation);
                            break;
                    }
                    
                    MessageBox.Show($"Optimized for {useCase}:\n{suggestion}", 
                        "Smart Optimization", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error optimizing size: {ex.Message}");
            }
        }

        /// <summary>
        /// Auto-fit content to current slide size
        /// </summary>
        private void AutoFitContentToSlide()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActivePresentation != null)
                {
                    var presentation = app.ActivePresentation;
                    int adjustedShapes = 0;
                    
                    foreach (PowerPoint.Slide slide in presentation.Slides)
                    {
                        foreach (PowerPoint.Shape shape in slide.Shapes)
                        {
                            // Check if shape is outside slide bounds
                            if (shape.Left < 0 || shape.Top < 0 || 
                                shape.Left + shape.Width > presentation.PageSetup.SlideWidth ||
                                shape.Top + shape.Height > presentation.PageSetup.SlideHeight)
                            {
                                // Auto-fit the shape to slide bounds
                                if (shape.Left < 0) shape.Left = 0;
                                if (shape.Top < 0) shape.Top = 0;
                                
                                if (shape.Left + shape.Width > presentation.PageSetup.SlideWidth)
                                    shape.Left = presentation.PageSetup.SlideWidth - shape.Width;
                                if (shape.Top + shape.Height > presentation.PageSetup.SlideHeight)
                                    shape.Top = presentation.PageSetup.SlideHeight - shape.Height;
                                
                                adjustedShapes++;
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error auto-fitting content: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        // Align shapes at their angles (corners) - use last selected shape as master
                        var masterShape = shapes[shapes.Count];
                        float targetAngle = masterShape.Rotation;
                        
                        foreach (PowerPoint.Shape shape in shapes)
                        {
                            shape.Rotation = targetAngle;
                        }

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

        private void BtnAlignRoundedRectangleRadius_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    if (shapes.Count > 1)
                    {
                        // Find rounded rectangles and align their radius
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
                            // Use the last selected shape as the "Master" for radius
                            var masterShape = roundedRects.Last();
                            
                            // Apply the master's radius to all other rounded rectangles
                            foreach (var rect in roundedRects)
                            {
                                if (rect != masterShape)
                                {
                                    // Note: PowerPoint doesn't expose corner radius directly, 
                                    // so this is a simplified implementation
                                    try
                                    {
                                        if (masterShape.Adjustments.Count > 0)
                                        {
                                            rect.Adjustments[1] = masterShape.Adjustments[1];
                                        }
                                    }
                                    catch
                                    {
                                        // Fallback: adjust using alternative method
                                        rect.AutoShapeType = masterShape.AutoShapeType;
                                    }
                                }
                            }

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
                    MessageBox.Show("Please select shapes to align rounded rectangle radius.", "Shape Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    }
                }
            }
            catch (Exception)
            {
                // Silently handle errors to avoid popup dialogs
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
                    }
                }
            }
            catch (Exception)
            {
                // Silently handle errors to avoid popup dialogs
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
                    }
                }
            }
            catch (Exception)
            {
                // Silently handle errors to avoid popup dialogs
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
                    
                    // Check if bullets are currently visible
                    bool bulletsVisible = textRange.ParagraphFormat.Bullet.Visible == Office.MsoTriState.msoTrue;
                    
                    if (bulletsVisible)
                    {
                        // Turn off bullets
                        textRange.ParagraphFormat.Bullet.Visible = Office.MsoTriState.msoFalse;
                    }
                    else
                    {
                        // Turn on bullet points (not numbered list)
                        textRange.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletUnnumbered;
                        textRange.ParagraphFormat.Bullet.Visible = Office.MsoTriState.msoTrue;
                        textRange.ParagraphFormat.Bullet.Character = 8226; // Unicode bullet •
                        textRange.ParagraphFormat.Bullet.Font.Name = "Symbol";
                        textRange.ParagraphFormat.Bullet.UseTextFont = Office.MsoTriState.msoFalse;
                    }
                    
                    System.Diagnostics.Debug.WriteLine("Bullet formatting toggled!");
                }
                else if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    // Handle shape selection - apply bullets to all text shapes
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    int shapesWithText = 0;
                    
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        var shape = shapes[i];
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                            shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            var textRange = shape.TextFrame.TextRange;
                            
                            // Apply bullet points (not numbered list)
                            textRange.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletUnnumbered;
                            textRange.ParagraphFormat.Bullet.Visible = Office.MsoTriState.msoTrue;
                            textRange.ParagraphFormat.Bullet.Character = 8226; // Unicode bullet •
                            textRange.ParagraphFormat.Bullet.Font.Name = "Symbol";
                            textRange.ParagraphFormat.Bullet.UseTextFont = Office.MsoTriState.msoFalse;
                            
                            shapesWithText++;
                        }
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"Bullet points applied to {shapesWithText} shape(s) with text");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("Please select text or shapes with text to add bullets");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error formatting bullets: {ex.Message}");
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

                    // Determine current wrap setting from first text shape
                    Office.MsoTriState? current = null;
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        if (shapes[i].HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            current = shapes[i].TextFrame2.WordWrap;
                            break;
                        }
                    }
                    var newVal = (current == Office.MsoTriState.msoTrue) ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;

                    int count = 0;
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        if (shapes[i].HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            shapes[i].TextFrame2.WordWrap = newVal;
                            count++;
                        }
                    }

                    if (count == 0)
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

        private void BtnNoWrapText_Click(object sender, EventArgs e) // deprecated
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

        #region Advanced Position Section - Extended Functionality

        // Basic Alignment Functions (missing ones)
        private void BtnAlignTop_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                bool useCtrlKey = Control.ModifierKeys.HasFlag(Keys.Control);
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (useCtrlKey || shapes.Count == 1)
                    {
                        // Align to slide top edge
                        shapes.Align(Office.MsoAlignCmd.msoAlignTops, Office.MsoTriState.msoTrue);
                    }
                    else
                    {
                        // Align to master object (last selected)
                        shapes.Align(Office.MsoAlignCmd.msoAlignTops, Office.MsoTriState.msoFalse);
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to align.", "Align Top", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning objects to top: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAlignBottom_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                bool useCtrlKey = Control.ModifierKeys.HasFlag(Keys.Control);
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (useCtrlKey || shapes.Count == 1)
                    {
                        // Align to slide bottom edge
                        shapes.Align(Office.MsoAlignCmd.msoAlignBottoms, Office.MsoTriState.msoTrue);
                    }
                    else
                    {
                        // Align to master object (last selected)
                        shapes.Align(Office.MsoAlignCmd.msoAlignBottoms, Office.MsoTriState.msoFalse);
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to align.", "Align Bottom", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning objects to bottom: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAlignMiddle_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                bool useCtrlKey = Control.ModifierKeys.HasFlag(Keys.Control);
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (useCtrlKey || shapes.Count == 1)
                    {
                        // Align to slide middle
                        shapes.Align(Office.MsoAlignCmd.msoAlignMiddles, Office.MsoTriState.msoTrue);
                    }
                    else
                    {
                        // Align to master object (last selected)
                        shapes.Align(Office.MsoAlignCmd.msoAlignMiddles, Office.MsoTriState.msoFalse);
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to align.", "Align Middle", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning objects to middle: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Docking Functions
        private void BtnDockLeft_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                bool useCtrlKey = Control.ModifierKeys.HasFlag(Keys.Control);
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (useCtrlKey || shapes.Count == 1)
                    {
                        // Move to left edge of slide
                        foreach (PowerPoint.Shape shape in shapes)
                        {
                            shape.Left = 0;
                        }
                    }
                    else
                    {
                        // Move to touch master object (last selected) on left
                        var master = shapes[shapes.Count];
                        float masterLeftEdge = master.Left;
                        
                        for (int i = 1; i < shapes.Count; i++)
                        {
                            shapes[i].Left = masterLeftEdge - shapes[i].Width;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to dock.", "Dock Left", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error docking objects left: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDockRight_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                bool useCtrlKey = Control.ModifierKeys.HasFlag(Keys.Control);
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (useCtrlKey || shapes.Count == 1)
                    {
                        // Move to right edge of slide
                        if (app.ActivePresentation != null)
                        {
                            float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                        
                        foreach (PowerPoint.Shape shape in shapes)
                        {
                            shape.Left = slideWidth - shape.Width;
                        }
                        }
                        else
                        {
                            MessageBox.Show("No active presentation found. Please open or create a presentation first.", "No Presentation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        // Move to touch master object (last selected) on right
                        var master = shapes[shapes.Count];
                        float masterRightEdge = master.Left + master.Width;
                        
                        for (int i = 1; i < shapes.Count; i++)
                        {
                            shapes[i].Left = masterRightEdge;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to dock.", "Dock Right", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error docking objects right: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDockTop_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                bool useCtrlKey = Control.ModifierKeys.HasFlag(Keys.Control);
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (useCtrlKey || shapes.Count == 1)
                    {
                        // Move to top edge of slide
                        foreach (PowerPoint.Shape shape in shapes)
                        {
                            shape.Top = 0;
                        }
                    }
                    else
                    {
                        // Move to touch master object (last selected) on top
                        var master = shapes[shapes.Count];
                        float masterTopEdge = master.Top;
                        
                        for (int i = 1; i < shapes.Count; i++)
                        {
                            shapes[i].Top = masterTopEdge - shapes[i].Height;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to dock.", "Dock Top", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error docking objects top: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDockBottom_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                bool useCtrlKey = Control.ModifierKeys.HasFlag(Keys.Control);
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (useCtrlKey || shapes.Count == 1)
                    {
                        // Move to bottom edge of slide
                        if (app.ActivePresentation != null)
                        {
                            float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
                        
                        foreach (PowerPoint.Shape shape in shapes)
                        {
                            shape.Top = slideHeight - shape.Height;
                        }
                        }
                        else
                        {
                            MessageBox.Show("No active presentation found. Please open or create a presentation first.", "No Presentation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        // Move to touch master object (last selected) on bottom
                        var master = shapes[shapes.Count];
                        float masterBottomEdge = master.Top + master.Height;
                        
                        for (int i = 1; i < shapes.Count; i++)
                        {
                            shapes[i].Top = masterBottomEdge;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to dock.", "Dock Bottom", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error docking objects bottom: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Enhanced Distribution Functions
        private void BtnDistributeHorizontal_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                bool useCtrlKey = Control.ModifierKeys.HasFlag(Keys.Control);
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (shapes.Count >= 3)
                    {
                        if (useCtrlKey)
                        {
                            // Distribute across entire slide width
                            if (app.ActivePresentation != null)
                            {
                                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                            var sortedShapes = shapes.Cast<PowerPoint.Shape>().OrderBy(s => s.Left).ToArray();
                            
                            float spacing = slideWidth / (sortedShapes.Length + 1);
                            for (int i = 0; i < sortedShapes.Length; i++)
                            {
                                sortedShapes[i].Left = spacing * (i + 1) - sortedShapes[i].Width / 2;
                            }
                            }
                            else
                            {
                                MessageBox.Show("No active presentation found. Please open or create a presentation first.", "No Presentation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            // Standard distribution (keeping leftmost and rightmost in place)
                            shapes.Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, Office.MsoTriState.msoFalse);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select at least 3 objects to distribute horizontally.", "Distribute Horizontal", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to distribute.", "Distribute Horizontal", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error distributing objects horizontally: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDistributeVertical_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                bool useCtrlKey = Control.ModifierKeys.HasFlag(Keys.Control);
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (shapes.Count >= 3)
                    {
                        if (useCtrlKey)
                        {
                            // Distribute across entire slide height
                            if (app.ActivePresentation != null)
                            {
                                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
                            var sortedShapes = shapes.Cast<PowerPoint.Shape>().OrderBy(s => s.Top).ToArray();
                            
                            float spacing = slideHeight / (sortedShapes.Length + 1);
                            for (int i = 0; i < sortedShapes.Length; i++)
                            {
                                sortedShapes[i].Top = spacing * (i + 1) - sortedShapes[i].Height / 2;
                            }
                            }
                            else
                            {
                                MessageBox.Show("No active presentation found. Please open or create a presentation first.", "No Presentation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            // Standard distribution (keeping topmost and bottommost in place)
                            shapes.Distribute(Office.MsoDistributeCmd.msoDistributeVertically, Office.MsoTriState.msoFalse);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select at least 3 objects to distribute vertically.", "Distribute Vertical", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to distribute.", "Distribute Vertical", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error distributing objects vertically: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Advanced Positioning Functions
        private void BtnGoldenCanon_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (shapes.Count >= 2)
                    {
                        var master = shapes[shapes.Count]; // Last selected
                        
                        // Golden ratio: margin at bottom is twice the margin at top
                        float masterTop = master.Top;
                        float masterBottom = master.Top + master.Height;
                        float availableHeight = masterBottom - masterTop;
                        
                        for (int i = 1; i < shapes.Count; i++)
                        {
                            var shape = shapes[i];
                            
                            // Calculate golden canon positioning
                            float topMargin = availableHeight / 3;
                            float bottomMargin = topMargin * 2;
                            
                            // Position object in the golden canon ratio
                            shape.Top = masterTop + topMargin;
                            
                            // Ensure it fits within the constraints
                            if (shape.Top + shape.Height > masterBottom - bottomMargin)
                            {
                                shape.Top = masterBottom - bottomMargin - shape.Height;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select at least 2 objects (master should be higher than objects to be aligned).", "Golden Canon", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to align in Golden Canon.", "Golden Canon", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning in Golden Canon: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAlignMatrix_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (shapes.Count >= 2)
                    {
                        // Get user input for matrix dimensions
                        string input = Microsoft.VisualBasic.Interaction.InputBox(
                            "Enter matrix dimensions (rows x columns):\nExample: 3x2 for 3 rows and 2 columns",
                            "Matrix Alignment",
                            "2x3");
                        
                        if (!string.IsNullOrEmpty(input))
                        {
                            var parts = input.ToLower().Split('x');
                            if (parts.Length == 2 && int.TryParse(parts[0], out int rows) && int.TryParse(parts[1], out int columns))
                            {
                                AlignInMatrix(shapes, rows, columns);
                            }
                            else
                            {
                                MessageBox.Show("Invalid format. Please use format like '3x2' for 3 rows and 2 columns.", "Matrix Alignment", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select at least 2 objects to arrange in matrix.", "Matrix Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to align in matrix.", "Matrix Alignment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning in matrix: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AlignInMatrix(PowerPoint.ShapeRange shapes, int rows, int columns)
        {
            // Calculate grid bounds based on selected shapes
            var shapesList = shapes.Cast<PowerPoint.Shape>().ToList();
            
            float minLeft = shapesList.Min(s => s.Left);
            float maxRight = shapesList.Max(s => s.Left + s.Width);
            float minTop = shapesList.Min(s => s.Top);
            float maxBottom = shapesList.Max(s => s.Top + s.Height);
            
            float totalWidth = maxRight - minLeft;
            float totalHeight = maxBottom - minTop;
            
            float cellWidth = totalWidth / columns;
            float cellHeight = totalHeight / rows;
            
            // Place objects in matrix (row-wise, top to bottom)
            for (int i = 0; i < Math.Min(shapes.Count, rows * columns); i++)
            {
                int row = i / columns;
                int col = i % columns;
                
                var shape = shapes[i + 1]; // PowerPoint uses 1-based indexing
                
                // Calculate cell center position
                float cellCenterX = minLeft + (col * cellWidth) + (cellWidth / 2);
                float cellCenterY = minTop + (row * cellHeight) + (cellHeight / 2);
                
                // Position shape at cell center
                shape.Left = cellCenterX - (shape.Width / 2);
                shape.Top = cellCenterY - (shape.Height / 2);
            }
        }

        private void BtnSliceShape_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (shapes.Count == 1)
                    {
                        string input = Microsoft.VisualBasic.Interaction.InputBox(
                            "Enter slice/multiply options:\nFormat: 'rows x columns' for slicing\nExample: '2x3' creates 6 shapes in 2 rows and 3 columns\nOptional spacing: '2x3 10' adds 10pt spacing",
                            "Slice or Multiply Shape",
                            "2x2");
                        
                        if (!string.IsNullOrEmpty(input))
                        {
                            var parts = input.Split(' ');
                            var dimensions = parts[0].Split('x');
                            
                            if (dimensions.Length == 2 && int.TryParse(dimensions[0], out int rows) && int.TryParse(dimensions[1], out int columns))
                            {
                                float spacing = 0;
                                if (parts.Length > 1 && float.TryParse(parts[1], out spacing))
                                {
                                    // Spacing provided
                                }
                                
                                SliceOrMultiplyShape(shapes[1], rows, columns, spacing);
                            }
                            else
                            {
                                MessageBox.Show("Invalid format. Use format like '2x3' or '2x3 10' (with spacing).", "Slice Shape", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select exactly one shape to slice/multiply.", "Slice Shape", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select a shape to slice/multiply.", "Slice Shape", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error slicing shape: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SliceOrMultiplyShape(PowerPoint.Shape originalShape, int rows, int columns, float spacing)
        {
            var slide = originalShape.Parent as PowerPoint.Slide;
            
            // Calculate individual shape dimensions
            float originalWidth = originalShape.Width;
            float originalHeight = originalShape.Height;
            float originalLeft = originalShape.Left;
            float originalTop = originalShape.Top;
            
            float shapeWidth = (originalWidth - spacing * (columns - 1)) / columns;
            float shapeHeight = (originalHeight - spacing * (rows - 1)) / rows;
            
            // Create the grid of shapes
            for (int row = 0; row < rows; row++)
            {
                for (int col = 0; col < columns; col++)
                {
                    if (row == 0 && col == 0)
                    {
                        // Resize the original shape
                        originalShape.Width = shapeWidth;
                        originalShape.Height = shapeHeight;
                        continue;
                    }
                    
                    // Duplicate the original shape
                    var newShape = originalShape.Duplicate()[1];
                    
                    // Position the new shape
                    newShape.Left = originalLeft + col * (shapeWidth + spacing);
                    newShape.Top = originalTop + row * (shapeHeight + spacing);
                    newShape.Width = shapeWidth;
                    newShape.Height = shapeHeight;
                }
            }
        }

        private void BtnDuplicateRight_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    foreach (PowerPoint.Shape shape in shapes)
                    {
                        var duplicate = shape.Duplicate()[1];
                        duplicate.Left = shape.Left + shape.Width + 10; // 10pt spacing
                        duplicate.Top = shape.Top;
                    }

                }
                else
                {
                    MessageBox.Show("Please select objects to duplicate.", "Duplicate Right", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error duplicating objects: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCenterTopLeft_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    if (shapes.Count >= 2)
                    {
                        var master = shapes[shapes.Count]; // Last selected
                        float masterTopLeftX = master.Left;
                        float masterTopLeftY = master.Top;
                        
                        for (int i = 1; i < shapes.Count; i++)
                        {
                            var shape = shapes[i];
                            
                            // Center shape on master's top-left corner
                            shape.Left = masterTopLeftX - (shape.Width / 2);
                            shape.Top = masterTopLeftY - (shape.Height / 2);
                        }

                    }
                    else
                    {
                        MessageBox.Show("Please select at least 2 objects (master is the last selected).", "Center on Top Left", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Please select objects to center.", "Center on Top Left", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error centering objects: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Save/Apply Position and Size Functions
        private struct SavedPosition
        {
            public float Left;
            public float Top;
            public float Width;
            public float Height;
        }

        private List<SavedPosition> savedPositions = new List<SavedPosition>();

        private void BtnSavePosition_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    savedPositions.Clear();
                    
                    foreach (PowerPoint.Shape shape in shapes)
                    {
                        savedPositions.Add(new SavedPosition
                        {
                            Left = shape.Left,
                            Top = shape.Top,
                            Width = shape.Width,
                            Height = shape.Height
                        });
                    }

                }
                else
                {
                    MessageBox.Show("Please select objects whose position and size you want to save.", "Save Position", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving positions: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnApplyPosition_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                
                if (savedPositions.Count == 0)
                {
                    MessageBox.Show("No saved positions found. Please use 'Save Position and Size' first.", "Apply Position", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = app.ActiveWindow.Selection.ShapeRange;
                    
                    for (int i = 0; i < Math.Min(shapes.Count, savedPositions.Count); i++)
                    {
                        var shape = shapes[i + 1]; // PowerPoint uses 1-based indexing
                        var savedPos = savedPositions[i];
                        
                        shape.Left = savedPos.Left;
                        shape.Top = savedPos.Top;
                        shape.Width = savedPos.Width;
                        shape.Height = savedPos.Height;
                    }

                }
                else
                {
                    MessageBox.Show("Please select objects to apply the saved position and size to.", "Apply Position", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying positions: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnRemoveMarginObjects_Click(object sender, EventArgs e)
        {
            try
            {
                var slide = GetActiveSlideOrNull();
                if (slide == null)
                {
                    MessageBox.Show("No active slide found.", "Remove Margin Objects", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var shapes = slide.Shapes;
                var shapesToDelete = new List<PowerPoint.Shape>();
                
                // Get the actual slide dimensions
                float slideWidth, slideHeight;
                try
                {
                    // Try to get from custom layout first
                    slideWidth = slide.CustomLayout.Width;
                    slideHeight = slide.CustomLayout.Height;
                    
                    // If that fails, use presentation page setup
                    if (slideWidth <= 0 || slideHeight <= 0)
                    {
                        slideWidth = slide.Parent.PageSetup.SlideWidth;
                        slideHeight = slide.Parent.PageSetup.SlideHeight;
                    }
                    
                    // If still fails, use master slide as fallback
                    if (slideWidth <= 0 || slideHeight <= 0)
                    {
                        slideWidth = slide.Master.Width;
                        slideHeight = slide.Master.Height;
                    }
                }
                catch
                {
                    // Final fallback to master slide
                    slideWidth = slide.Master.Width;
                    slideHeight = slide.Master.Height;
                }

                // Define margin threshold (objects outside this area will be removed)
                // Use a smaller threshold to be more sensitive to objects outside the main layout
                float marginThreshold = 20; // 20 points margin (reduced from 50)
                float safeLeft = -marginThreshold;
                float safeTop = -marginThreshold;
                float safeRight = slideWidth + marginThreshold;
                float safeBottom = slideHeight + marginThreshold;

                System.Diagnostics.Debug.WriteLine($"Slide dimensions: {slideWidth} x {slideHeight}");
                System.Diagnostics.Debug.WriteLine($"Safe area: Left={safeLeft}, Top={safeTop}, Right={safeRight}, Bottom={safeBottom}");

                // Check each shape
                for (int i = 1; i <= shapes.Count; i++)
                {
                    var shape = shapes[i];
                    
                    // Check if shape is completely outside the safe area
                    bool isOutside = shape.Left + shape.Width < safeLeft || 
                                   shape.Left > safeRight || 
                                   shape.Top + shape.Height < safeTop || 
                                   shape.Top > safeBottom;

                    System.Diagnostics.Debug.WriteLine($"Shape {i}: Left={shape.Left}, Top={shape.Top}, Width={shape.Width}, Height={shape.Height}, IsOutside={isOutside}");

                    if (isOutside)
                    {
                        shapesToDelete.Add(shape);
                        System.Diagnostics.Debug.WriteLine($"Added shape {i} to deletion list");
                    }
                }

                if (shapesToDelete.Count == 0)
                {
                    // No margin objects found - work silently
                    return;
                }

                // Delete shapes (in reverse order to avoid index issues)
                for (int i = shapesToDelete.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        shapesToDelete[i].Delete();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error deleting shape: {ex.Message}");
                    }
                }

                // Log success silently
                System.Diagnostics.Debug.WriteLine($"Successfully removed {shapesToDelete.Count} margin object(s)");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error removing margin objects: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Size Tools Event Handlers

        /// <summary>
        /// Property to get the PowerPoint Application instance
        /// </summary>
        private PowerPoint.Application pptApp => Globals.ThisAddIn.Application;

        /// <summary>
        /// Align Width/Height event handler
        /// </summary>
        private void CmbAlign_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string selected = cmbAlign.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(selected)) return;

                var sel = pptApp.ActiveWindow.Selection;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 1)
                {
                    var shapes = sel.ShapeRange;
                    PowerPoint.Shape master = shapes[shapes.Count]; // last selected = master

                    foreach (PowerPoint.Shape shp in shapes)
                    {
                        if (shp != master)
                        {
                            if (selected.Contains("Width")) shp.Width = master.Width;
                            if (selected.Contains("Height")) shp.Height = master.Height;
                        }
                    }
                    
                    // Reset the selection to avoid triggering again
                    cmbAlign.SelectedIndex = -1;
                }
                else
                {
                    MessageBox.Show("Please select at least 2 shapes to align their dimensions.", "Size Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cmbAlign.SelectedIndex = -1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error aligning shapes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbAlign.SelectedIndex = -1;
            }
        }

        /// <summary>
        /// Stretch Functions event handler
        /// </summary>
        private void CmbStretch_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string selected = cmbStretch.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(selected)) return;

                var sel = pptApp.ActiveWindow.Selection;
                var slide = pptApp.ActiveWindow.View.Slide;

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 1)
                {
                    var shapes = sel.ShapeRange;
                    PowerPoint.Shape master = shapes[shapes.Count];

                    foreach (PowerPoint.Shape shp in shapes)
                    {
                        if (shp != master)
                        {
                            switch (selected)
                            {
                                case "Stretch Left": shp.Left = master.Left; break;
                                case "Stretch Right": shp.Left = master.Left + master.Width - shp.Width; break;
                                case "Stretch Up": shp.Top = master.Top; break;
                                case "Stretch Down": shp.Top = master.Top + master.Height - shp.Height; break;
                            }
                        }
                    }
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count == 1)
                {
                    // Single object - stretch to slide edge
                    var shape = sel.ShapeRange[1];
                    switch (selected)
                    {
                        case "Stretch Left": shape.Left = 0; break;
                        case "Stretch Right": shape.Left = slide.Master.Width - shape.Width; break;
                        case "Stretch Up": shape.Top = 0; break;
                        case "Stretch Down": shape.Top = slide.Master.Height - shape.Height; break;
                    }
                }
                else
                {
                    MessageBox.Show("Please select at least one shape to stretch.", "Size Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                
                // Reset the selection
                cmbStretch.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error stretching shapes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbStretch.SelectedIndex = -1;
            }
        }

        /// <summary>
        /// Fill Functions event handler
        /// </summary>
        private void CmbFill_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string selected = cmbFill.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(selected)) return;

                var sel = pptApp.ActiveWindow.Selection;

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 1)
                {
                    var shapes = sel.ShapeRange;
                    PowerPoint.Shape master = shapes[shapes.Count];

                    foreach (PowerPoint.Shape shp in shapes)
                    {
                        if (shp != master)
                        {
                            switch (selected)
                            {
                                case "Fill Left": 
                                    shp.Width += (master.Left - shp.Left); 
                                    shp.Left = master.Left; 
                                    break;
                                case "Fill Right": 
                                    shp.Width = (master.Left + master.Width) - shp.Left; 
                                    break;
                                case "Fill Up": 
                                    shp.Height += (master.Top - shp.Top); 
                                    shp.Top = master.Top; 
                                    break;
                                case "Fill Down": 
                                    shp.Height = (master.Top + master.Height) - shp.Top; 
                                    break;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select at least 2 shapes to use fill functions.", "Size Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                
                // Reset the selection
                cmbFill.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error filling shapes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbFill.SelectedIndex = -1;
            }
        }

        /// <summary>
        /// Magic Resizer event handler
        /// </summary>
        private void BtnMagicResizer_Click(object sender, EventArgs e)
        {
            try
            {
                var sel = pptApp.ActiveWindow.Selection;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapes = sel.ShapeRange;
                    foreach (PowerPoint.Shape shp in shapes)
                    {
                        // Increase size by 10%
                        shp.Width *= 1.1f;
                        shp.Height *= 1.1f;
                        
                        // Increase line weight if shape has a line
                        if (shp.Line.Visible == Office.MsoTriState.msoTrue)
                        {
                            shp.Line.Weight *= 1.1f;
                        }
                        
                        // Increase font size if shape has text
                        if (shp.HasTextFrame == Office.MsoTriState.msoTrue && 
                            shp.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                        {
                            shp.TextFrame2.TextRange.Font.Size *= 1.1f;
                        }
                    }

                }
                else
                {
                    MessageBox.Show("Please select one or more shapes to resize.", "Magic Resizer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error with Magic Resizer: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #endregion

        /// <summary>
        /// Loads images for presentation buttons
        /// </summary>
        private void LoadPresentationButtonImages()
        {
            try
            {
                // Presentation button images mapping
                var presentationButtons = new Dictionary<Button, string>
                {
                    { btnNew, "icons8-file-50.png" },
                    { btnOpen, "icons8-open-file-48.png" },
                    { btnSave, "icons8-save-50.png" },
                    { btnSaveAs, "icons8-save-as-50.png" },
                    { btnPrint, "icons8-export-50.png" }
                };

                foreach (var button in presentationButtons)
                {
                    if (button.Key != null)
                    {
                        LoadImageForButton(button.Key, button.Value, "file", true);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading presentation button images: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads images for smart elements buttons
        /// </summary>
        private void LoadSmartElementsButtonImages()
        {
            try
            {
                // Smart elements button images mapping
                var smartElementsButtons = new Dictionary<Button, string>
                {
                    { btnChart, "icons8-chart-60.png" },
                    { btnDiagram, "icons8-color-palette-48.png" },
                    { btnTable, "icons8-table-50.png" },
                    { btnMatrixTable, "icons8-matrix-50.png" },
                    { btnStickyNote, "icons8-sticky-notes-50.png" },
                    { btnCitation, "icons8-get-quote-30.png" },
                    { btnStandardObjects, "icons8-object-50.png" }
                };

                foreach (var button in smartElementsButtons)
                {
                    if (button.Key != null)
                    {
                        LoadImageForButton(button.Key, button.Value, "elements", true);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading smart elements button images: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads images for position buttons
        /// </summary>
        private void LoadPositionButtonImages()
        {
            try
            {
                // Position button images mapping
                var positionButtons = new Dictionary<Button, string>
                {
                    { btnAlignLeft, "icons8-align-left-64.png" },
                    { btnAlignCenter, "icons8-align-center-64.png" },
                    { btnAlignRight, "icons8-align-right-64.png" },
                    { btnAlignTop, "icons8-align-top-64.png" },
                    { btnAlignBottom, "icons8-align-bottom-64.png" },
                    { btnAlignMiddle, "icons8-align-center-64.png" },
                    { btnDockLeft, "icons8-align-left-64.png" },
                    { btnDockRight, "icons8-align-right-64.png" },
                    { btnDockTop, "icons8-align-top-64.png" },
                    { btnDockBottom, "icons8-align-bottom-64.png" },
                    { btnDistribute, "icons8-align-justify-64.png" },
                    { btnDistributeHorizontal, "icons8-align-center-64.png" },
                    { btnDistributeVertical, "icons8-align-justify-64.png" },
                    { btnMatchBoth, "icons8-enlarge-50.png" },
                    { btnMatchHeight, "icons8-height-50.png" },
                    { btnMatchWidth, "icons8-width-50.png" },
                    { btnMakeVertical, "icons8-rotate-left-48.png" },
                    { btnMakeHorizontal, "icons8-rotate-right-48.png" },
                    { btnSwapLocations, "icons8-swap-50.png" },
                    { btnGoldenCanon, "icons8-swap-50.png" },
                    { btnAlignMatrix, "icons8-matrix-50.png" },
                    { btnSliceShape, "icons8-slice-50.png" },
                    { btnDuplicateRight, "icons8-duplicate-50.png" },
                    { btnCenterTopLeft, "icons8-snap-to-center-48.png" },
                    { btnSavePosition, "icons8-save-50.png" },
                    { btnApplyPosition, "icons8-apply-64.png" }
                };

                foreach (var button in positionButtons)
                {
                    if (button.Key != null)
                    {
                        LoadImageForButton(button.Key, button.Value, "position", true);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading position button images: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads an image from embedded resources
        /// </summary>
        /// <param name="resourceName">Name of the embedded resource (e.g., "icons8-file-50.png")</param>
        /// <param name="subfolder">Subfolder within icons (e.g., "file", "wizzards", "position")</param>
        /// <returns>Image object or null if not found</returns>
        private Image LoadEmbeddedImage(string resourceName, string subfolder)
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                string resourcePath = $"my_addin.icons.{subfolder}.{resourceName}";
                
                using (var stream = assembly.GetManifestResourceStream(resourcePath))
                {
                    if (stream != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"✅ Loaded embedded image: {resourcePath}");
                        return Image.FromStream(stream);
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"❌ Embedded resource not found: {resourcePath}");
                        
                        // List available resources for debugging
                        var resourceNames = assembly.GetManifestResourceNames();
                        System.Diagnostics.Debug.WriteLine($"Available resources: {string.Join(", ", resourceNames.Where(r => r.Contains("icons")))}");
                        
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading embedded image {resourceName}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Loads an image for a button, trying embedded resources first, then file system
        /// </summary>
        /// <param name="button">Button to set image for</param>
        /// <param name="imageName">Name of the image file</param>
        /// <param name="subfolder">Subfolder (file, wizzards, position, elements)</param>
        /// <param name="useEmbedded">Whether to try embedded resources first</param>
        private void LoadImageForButton(Button button, string imageName, string subfolder = "", bool useEmbedded = false)
        {
            try
            {
                Image loadedImage = null;

                if (useEmbedded)
                {
                    // Try embedded resource first
                    loadedImage = LoadEmbeddedImage(imageName, subfolder);
                }

                if (loadedImage == null)
                {
                    // Fallback to file system
                    string imagePath = FindImagePath(imageName, subfolder);
                    if (!string.IsNullOrEmpty(imagePath))
                    {
                        loadedImage = Image.FromFile(imagePath);
                        System.Diagnostics.Debug.WriteLine($"✅ Loaded image from file: {imagePath}");
                    }
                }

                if (loadedImage != null)
                {
                    button.BackgroundImage = loadedImage;
                    button.BackgroundImageLayout = ImageLayout.Zoom;
                }
                else
                {
                    button.BackgroundImage = null;
                    System.Diagnostics.Debug.WriteLine($"❌ Image not found: {imageName} in {subfolder}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading image for button {button.Name}: {ex.Message}");
                button.BackgroundImage = null;
            }
        }

        private void TaskPaneControl_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control && e.KeyCode == Keys.V)
                {
                    var addin = Globals.ThisAddIn;
                    if (addin != null && addin.TryPasteIntoMatrix())
                    {
                        e.Handled = true;
                    }
                }
            }
            catch { }
        }
    }
}