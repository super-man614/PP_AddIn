using System;
using System.Drawing;
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
            SetupEventHandlers();
            SetInitialValues();
        }

        private void SetupEventHandlers()
        {
            // Size section events
            if (cmbSlideSize != null)
                cmbSlideSize.SelectedIndexChanged += CmbSlideSize_SelectedIndexChanged;
            if (btnApplySize != null)
                btnApplySize.Click += BtnApplySize_Click;
            
            // Button hover events
            if (btnNew != null)
            {
                btnNew.MouseEnter += BtnNew_MouseEnter;
                btnNew.MouseLeave += BtnNew_MouseLeave;
            }
            if (btnOpen != null)
            {
                btnOpen.MouseEnter += BtnOpen_MouseEnter;
                btnOpen.MouseLeave += BtnOpen_MouseLeave;
            }
            if (btnSave != null)
            {
                btnSave.MouseEnter += BtnSave_MouseEnter;
                btnSave.MouseLeave += BtnSave_MouseLeave;
            }
            if (btnSaveAs != null)
            {
                btnSaveAs.MouseEnter += BtnSaveAs_MouseEnter;
                btnSaveAs.MouseLeave += BtnSaveAs_MouseLeave;
            }
            if (btnPrint != null)
            {
                btnPrint.MouseEnter += BtnPrint_MouseEnter;
                btnPrint.MouseLeave += BtnPrint_MouseLeave;
            }
            if (btnShare != null)
            {
                btnShare.MouseEnter += BtnShare_MouseEnter;
                btnShare.MouseLeave += BtnShare_MouseLeave;
            }
            
            // Wizards section button hover events
            if (btnAgenda != null)
            {
                btnAgenda.MouseEnter += BtnAgenda_MouseEnter;
                btnAgenda.MouseLeave += BtnAgenda_MouseLeave;
            }
            if (btnMaster != null)
            {
                btnMaster.MouseEnter += BtnMaster_MouseEnter;
                btnMaster.MouseLeave += BtnMaster_MouseLeave;
            }
            if (btnElement != null)
            {
                btnElement.MouseEnter += BtnElement_MouseEnter;
                btnElement.MouseLeave += BtnElement_MouseLeave;
            }
            if (btnText != null)
            {
                btnText.MouseEnter += BtnText_MouseEnter;
                btnText.MouseLeave += BtnText_MouseLeave;
            }
            if (btnFormat != null)
            {
                btnFormat.MouseEnter += BtnFormat_MouseEnter;
                btnFormat.MouseLeave += BtnFormat_MouseLeave;
            }
            if (btnMap != null)
            {
                btnMap.MouseEnter += BtnMap_MouseEnter;
                btnMap.MouseLeave += BtnMap_MouseLeave;
            }
            
            // Smart Elements section button hover events
            if (btnChart != null)
            {
                btnChart.MouseEnter += BtnChart_MouseEnter;
                btnChart.MouseLeave += BtnChart_MouseLeave;
            }
            if (btnDiagram != null)
            {
                btnDiagram.MouseEnter += BtnDiagram_MouseEnter;
                btnDiagram.MouseLeave += BtnDiagram_MouseLeave;
            }
            if (btnTable != null)
            {
                btnTable.MouseEnter += BtnTable_MouseEnter;
                btnTable.MouseLeave += BtnTable_MouseLeave;
            }
            
            // Position section button hover events
            if (btnAlignLeft != null)
            {
                btnAlignLeft.MouseEnter += BtnAlignLeft_MouseEnter;
                btnAlignLeft.MouseLeave += BtnAlignLeft_MouseLeave;
            }
            if (btnAlignCenter != null)
            {
                btnAlignCenter.MouseEnter += BtnAlignCenter_MouseEnter;
                btnAlignCenter.MouseLeave += BtnAlignCenter_MouseLeave;
            }
            if (btnAlignRight != null)
            {
                btnAlignRight.MouseEnter += BtnAlignRight_MouseEnter;
                btnAlignRight.MouseLeave += BtnAlignRight_MouseLeave;
            }
            if (btnDistribute != null)
            {
                btnDistribute.MouseEnter += BtnDistribute_MouseEnter;
                btnDistribute.MouseLeave += BtnDistribute_MouseLeave;
            }
            
            // Shape section button hover events
            if (btnRectangle != null)
            {
                btnRectangle.MouseEnter += BtnRectangle_MouseEnter;
                btnRectangle.MouseLeave += BtnRectangle_MouseLeave;
            }
            if (btnCircle != null)
            {
                btnCircle.MouseEnter += BtnCircle_MouseEnter;
                btnCircle.MouseLeave += BtnCircle_MouseLeave;
            }
            if (btnArrow != null)
            {
                btnArrow.MouseEnter += BtnArrow_MouseEnter;
                btnArrow.MouseLeave += BtnArrow_MouseLeave;
            }
            if (btnLine != null)
            {
                btnLine.MouseEnter += BtnLine_MouseEnter;
                btnLine.MouseLeave += BtnLine_MouseLeave;
            }
            
            // Load current slides
            this.Load += TaskPaneControl_Load;
        }

        private void SetInitialValues()
        {
            // Set default combo box selection
            if (cmbSlideSize != null && cmbSlideSize.Items.Count > 1)
                cmbSlideSize.SelectedIndex = 1; // Default to 16:9
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

        // Shape section button hover methods
        private void BtnRectangle_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnRectangle_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnCircle_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnCircle_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnArrow_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnArrow_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 0;
            }
        }

        private void BtnLine_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                btn.FlatAppearance.BorderSize = 1;
            }
        }

        private void BtnLine_MouseLeave(object sender, EventArgs e)
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
                    var slide = app.ActiveWindow.Selection.SlideRange[1];
                    slide.Shapes.AddTable(NumRows: 3, NumColumns: 3);
                    MessageBox.Show("Table added successfully!", "Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please select a slide first, then use Insert > Table.", "Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting table: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

        private void BtnRectangle_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    var slide = app.ActiveWindow.Selection.SlideRange[1];
                    slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 100, 100, 200, 100);
                    MessageBox.Show("Rectangle added!", "Shape", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding rectangle: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCircle_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    var slide = app.ActiveWindow.Selection.SlideRange[1];
                    slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, 100, 100, 150, 150);
                    MessageBox.Show("Circle added!", "Shape", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding circle: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnArrow_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    var slide = app.ActiveWindow.Selection.SlideRange[1];
                    slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightArrow, 100, 100, 200, 50);
                    MessageBox.Show("Arrow added!", "Shape", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding arrow: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnLine_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    var slide = app.ActiveWindow.Selection.SlideRange[1];
                    slide.Shapes.AddLine(100, 100, 300, 100);
                    MessageBox.Show("Line added!", "Shape", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding line: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
    }
} 