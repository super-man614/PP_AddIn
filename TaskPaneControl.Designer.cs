using System;
using System.Drawing;
using System.Windows.Forms;

namespace my_addin
{
    partial class TaskPaneControl
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.mainScrollPanel = new System.Windows.Forms.Panel();
            this.sectionsContainer = new System.Windows.Forms.FlowLayoutPanel();
            this.presentationPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblPresentationSection = new System.Windows.Forms.Label();
            this.presentationButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnOpen = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnSaveAs = new System.Windows.Forms.Button();
            this.btnPrint = new System.Windows.Forms.Button();
            //this.btnShare = new System.Windows.Forms.Button();
            this.divider1 = new System.Windows.Forms.Panel();
            this.wizardsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblWizardsSection = new System.Windows.Forms.Label();
            this.wizardButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnAgenda = new System.Windows.Forms.Button();
            this.btnMaster = new System.Windows.Forms.Button();
            this.btnElement = new System.Windows.Forms.Button();
            this.btnText = new System.Windows.Forms.Button();
            this.btnFormat = new System.Windows.Forms.Button();
            this.btnMap = new System.Windows.Forms.Button();
            this.divider2 = new System.Windows.Forms.Panel();
            this.smartElementsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblSmartElementsSection = new System.Windows.Forms.Label();
            this.smartElementsButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnChart = new System.Windows.Forms.Button();
            this.btnDiagram = new System.Windows.Forms.Button();
            this.btnTable = new System.Windows.Forms.Button();
            this.btnMatrixTable = new System.Windows.Forms.Button();
            this.btnStickyNote = new System.Windows.Forms.Button();
            this.btnCitation = new System.Windows.Forms.Button();
            this.btnStandardObjects = new System.Windows.Forms.Button();
            this.divider3 = new System.Windows.Forms.Panel();
            this.positionPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblPositionSection = new System.Windows.Forms.Label();
            this.positionButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnAlignLeft = new System.Windows.Forms.Button();
            this.btnAlignCenter = new System.Windows.Forms.Button();
            this.btnAlignRight = new System.Windows.Forms.Button();
            this.btnAlignTop = new System.Windows.Forms.Button();
            this.btnAlignBottom = new System.Windows.Forms.Button();
            this.btnAlignMiddle = new System.Windows.Forms.Button();
            this.btnDockLeft = new System.Windows.Forms.Button();
            this.btnDockRight = new System.Windows.Forms.Button();
            this.btnDockTop = new System.Windows.Forms.Button();
            this.btnDockBottom = new System.Windows.Forms.Button();
            this.btnDistribute = new System.Windows.Forms.Button();
            this.btnDistributeHorizontal = new System.Windows.Forms.Button();
            this.btnDistributeVertical = new System.Windows.Forms.Button();
            this.btnMatchBoth = new System.Windows.Forms.Button();
            this.btnMatchHeight = new System.Windows.Forms.Button();
            this.btnMatchWidth = new System.Windows.Forms.Button();
            this.btnMakeVertical = new System.Windows.Forms.Button();
            this.btnMakeHorizontal = new System.Windows.Forms.Button();
            this.btnSwapLocations = new System.Windows.Forms.Button();
            this.btnGoldenCanon = new System.Windows.Forms.Button();
            this.btnAlignMatrix = new System.Windows.Forms.Button();
            this.btnSliceShape = new System.Windows.Forms.Button();
            this.btnDuplicateRight = new System.Windows.Forms.Button();
            this.btnCenterTopLeft = new System.Windows.Forms.Button();
            this.btnSavePosition = new System.Windows.Forms.Button();
            this.btnApplyPosition = new System.Windows.Forms.Button();
            this.divider4 = new System.Windows.Forms.Panel();
            this.sizePanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblSizeSection = new System.Windows.Forms.Label();
            this.sizeControlsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.widthPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.lblSlideSize = new System.Windows.Forms.Label();
            this.cmbSlideSize = new System.Windows.Forms.ComboBox();
            this.heightPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.lblWidth = new System.Windows.Forms.Label();
            this.nudWidth = new System.Windows.Forms.NumericUpDown();
            this.flowLayoutPanel3 = new System.Windows.Forms.FlowLayoutPanel();
            this.lblHeight = new System.Windows.Forms.Label();
            this.nudHeight = new System.Windows.Forms.NumericUpDown();
            this.btnApplySize = new System.Windows.Forms.Button();
            this.divider5 = new System.Windows.Forms.Panel();
            this.shapePanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblShapeSection = new System.Windows.Forms.Label();
            this.shapeButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnAlignProcessChain = new System.Windows.Forms.Button();
            this.btnAlignAngles = new System.Windows.Forms.Button();
            this.btnAlignToProcessArrow = new System.Windows.Forms.Button();
            this.btnAdjustPentagonHeader = new System.Windows.Forms.Button();
            this.btnAlignBlockArrows = new System.Windows.Forms.Button();
            this.btnAlignRoundedRectangleArrows = new System.Windows.Forms.Button();
            this.divider6 = new System.Windows.Forms.Panel();
            this.colorPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblColorSection = new System.Windows.Forms.Label();
            this.colorButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnFillColor = new System.Windows.Forms.Button();
            this.btnTextColor = new System.Windows.Forms.Button();
            this.btnOutlineColor = new System.Windows.Forms.Button();
            this.divider7 = new System.Windows.Forms.Panel();
            this.textPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblTextSection = new System.Windows.Forms.Label();
            this.textButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnBold = new System.Windows.Forms.Button();
            this.btnItalic = new System.Windows.Forms.Button();
            this.btnUnderline = new System.Windows.Forms.Button();
            this.btnBullets = new System.Windows.Forms.Button();
            this.btnWrapText = new System.Windows.Forms.Button();
            this.btnNoWrapText = new System.Windows.Forms.Button();
            this.divider8 = new System.Windows.Forms.Panel();
            this.navigationPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblNavigationSection = new System.Windows.Forms.Label();
            this.navigationButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnZoomIn = new System.Windows.Forms.Button();
            this.btnZoomOut = new System.Windows.Forms.Button();
            this.btnFitToWindow = new System.Windows.Forms.Button();
            this.divider9 = new System.Windows.Forms.Panel();
            this.expertToolsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblExpertToolsSection = new System.Windows.Forms.Label();
            this.expertToolsButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnFreeWebinar = new System.Windows.Forms.Button();
            this.mainScrollPanel.SuspendLayout();
            this.sectionsContainer.SuspendLayout();
            this.presentationPanel.SuspendLayout();
            this.presentationButtonsPanel.SuspendLayout();
            this.wizardsPanel.SuspendLayout();
            this.wizardButtonsPanel.SuspendLayout();
            this.smartElementsPanel.SuspendLayout();
            this.smartElementsButtonsPanel.SuspendLayout();
            this.positionPanel.SuspendLayout();
            this.positionButtonsPanel.SuspendLayout();
            this.sizePanel.SuspendLayout();
            this.sizeControlsPanel.SuspendLayout();
            this.widthPanel.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.heightPanel.SuspendLayout();
            this.flowLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudWidth)).BeginInit();
            this.flowLayoutPanel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudHeight)).BeginInit();
            this.shapePanel.SuspendLayout();
            this.shapeButtonsPanel.SuspendLayout();
            this.colorPanel.SuspendLayout();
            this.colorButtonsPanel.SuspendLayout();
            this.textPanel.SuspendLayout();
            this.textButtonsPanel.SuspendLayout();
            this.navigationPanel.SuspendLayout();
            this.navigationButtonsPanel.SuspendLayout();
            this.expertToolsPanel.SuspendLayout();
            this.expertToolsButtonsPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainScrollPanel
            // 
            this.mainScrollPanel.AutoScroll = true;
            this.mainScrollPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.mainScrollPanel.Controls.Add(this.sectionsContainer);
            this.mainScrollPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainScrollPanel.Location = new System.Drawing.Point(0, 0);
            this.mainScrollPanel.Name = "mainScrollPanel";
            this.mainScrollPanel.Padding = new System.Windows.Forms.Padding(2);
            this.mainScrollPanel.Size = new System.Drawing.Size(300, 600);
            this.mainScrollPanel.TabIndex = 0;
            // 
            // sectionsContainer
            // 
            this.sectionsContainer.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.sectionsContainer.AutoSize = true;
            this.sectionsContainer.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.sectionsContainer.Controls.Add(this.presentationPanel);
            this.sectionsContainer.Controls.Add(this.divider1);
            this.sectionsContainer.Controls.Add(this.wizardsPanel);
            this.sectionsContainer.Controls.Add(this.divider2);
            this.sectionsContainer.Controls.Add(this.smartElementsPanel);
            this.sectionsContainer.Controls.Add(this.divider3);
            this.sectionsContainer.Controls.Add(this.positionPanel);
            this.sectionsContainer.Controls.Add(this.divider4);
            this.sectionsContainer.Controls.Add(this.sizePanel);
            this.sectionsContainer.Controls.Add(this.divider5);
            this.sectionsContainer.Controls.Add(this.shapePanel);
            this.sectionsContainer.Controls.Add(this.divider6);
            this.sectionsContainer.Controls.Add(this.colorPanel);
            this.sectionsContainer.Controls.Add(this.divider7);
            this.sectionsContainer.Controls.Add(this.textPanel);
            this.sectionsContainer.Controls.Add(this.divider8);
            this.sectionsContainer.Controls.Add(this.navigationPanel);
            this.sectionsContainer.Controls.Add(this.divider9);
            this.sectionsContainer.Controls.Add(this.expertToolsPanel);
            this.sectionsContainer.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.sectionsContainer.Location = new System.Drawing.Point(0, 0);
            this.sectionsContainer.Name = "sectionsContainer";
            this.sectionsContainer.Size = new System.Drawing.Size(293, 623);
            this.sectionsContainer.TabIndex = 0;
            this.sectionsContainer.WrapContents = false;
            // 
            // presentationPanel
            // 
            this.presentationPanel.AutoSize = true;
            this.presentationPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.presentationPanel.BackColor = System.Drawing.Color.White;
            this.presentationPanel.Controls.Add(this.lblPresentationSection);
            this.presentationPanel.Controls.Add(this.presentationButtonsPanel);
            this.presentationPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.presentationPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.presentationPanel.Location = new System.Drawing.Point(3, 3);
            this.presentationPanel.Name = "presentationPanel";
            this.presentationPanel.Size = new System.Drawing.Size(287, 38);
            this.presentationPanel.TabIndex = 0;
            // 
            // lblPresentationSection
            // 
            this.lblPresentationSection.AutoSize = true;
            this.lblPresentationSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblPresentationSection.ForeColor = System.Drawing.Color.Gray;
            this.lblPresentationSection.Location = new System.Drawing.Point(3, 0);
            this.lblPresentationSection.Name = "lblPresentationSection";
            this.lblPresentationSection.Size = new System.Drawing.Size(72, 13);
            this.lblPresentationSection.TabIndex = 0;
            this.lblPresentationSection.Text = "Presentation";
            // 
            // presentationButtonsPanel
            // 
            this.presentationButtonsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.presentationButtonsPanel.AutoSize = true;
            this.presentationButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.presentationButtonsPanel.Controls.Add(this.btnNew);
            this.presentationButtonsPanel.Controls.Add(this.btnOpen);
            this.presentationButtonsPanel.Controls.Add(this.btnSave);
            this.presentationButtonsPanel.Controls.Add(this.btnSaveAs);
            this.presentationButtonsPanel.Controls.Add(this.btnPrint);
            //this.presentationButtonsPanel.Controls.Add(this.btnShare);
            this.presentationButtonsPanel.Location = new System.Drawing.Point(0, 13);
            this.presentationButtonsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.presentationButtonsPanel.Name = "presentationButtonsPanel";
            this.presentationButtonsPanel.Size = new System.Drawing.Size(150, 25);
            this.presentationButtonsPanel.TabIndex = 1;
            // 
            // btnNew
            // 
            this.btnNew.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnNew.FlatAppearance.BorderSize = 0;
            this.btnNew.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnNew.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnNew.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNew.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnNew.ForeColor = System.Drawing.Color.DarkBlue;
            this.btnNew.Location = new System.Drawing.Point(0, 0);
            this.btnNew.Margin = new System.Windows.Forms.Padding(1);
            this.btnNew.Name = "btnNew";
            this.btnNew.TabIndex = 1;
            //this.btnNew.Text = "📄";
            this.btnNew.Size = new System.Drawing.Size(20, 20);
            this.btnNew.BackColor = Color.Transparent;
            this.btnNew.BackgroundImage = Image.FromFile("icons/file/icons8-file-50.png");
            this.btnNew.BackgroundImageLayout = ImageLayout.Stretch;
            this.btnNew.UseVisualStyleBackColor = false;
            this.btnNew.Click += new System.EventHandler(this.BtnNew_Click);
            // 
            // btnOpen
            // 
            this.btnOpen.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnOpen.FlatAppearance.BorderSize = 0;
            this.btnOpen.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnOpen.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnOpen.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOpen.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnOpen.Location = new System.Drawing.Point(25, 0);
            this.btnOpen.Margin = new System.Windows.Forms.Padding(1);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.TabIndex = 2;
            this.btnOpen.Size = new System.Drawing.Size(20, 20);
            this.btnOpen.BackColor = Color.Transparent;
            this.btnOpen.BackgroundImage = Image.FromFile("icons/file/icons8-open-file-48.png");
            this.btnOpen.BackgroundImageLayout = ImageLayout.Stretch;
            //this.btnOpen.Text = "📂";
            this.btnOpen.UseVisualStyleBackColor = false;
            this.btnOpen.Click += new System.EventHandler(this.BtnOpen_Click);
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnSave.FlatAppearance.BorderSize = 0;
            this.btnSave.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnSave.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSave.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnSave.Location = new System.Drawing.Point(50, 0);
            this.btnSave.Margin = new System.Windows.Forms.Padding(1);
            this.btnSave.Name = "btnSave";
            this.btnSave.TabIndex = 3;
            //this.btnSave.Text = "💾";
            this.btnSave.Size = new System.Drawing.Size(20, 20);
            this.btnSave.BackColor = Color.Transparent;
            this.btnSave.BackgroundImage = Image.FromFile("icons/file/icons8-save-50.png");
            this.btnSave.BackgroundImageLayout = ImageLayout.Stretch;
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // btnSaveAs
            // 
            this.btnSaveAs.BackColor = System.Drawing.Color.Transparent;
            this.btnSaveAs.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnSaveAs.FlatAppearance.BorderSize = 0;
            this.btnSaveAs.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnSaveAs.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnSaveAs.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSaveAs.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnSaveAs.Location = new System.Drawing.Point(75, 0);
            this.btnSaveAs.Margin = new System.Windows.Forms.Padding(1);
            this.btnSaveAs.Name = "btnSaveAs";
            this.btnSaveAs.Size = new System.Drawing.Size(20, 20);
            this.btnSaveAs.TabIndex = 4;
            this.btnSaveAs.BackColor = Color.Transparent;
            this.btnSaveAs.BackgroundImage = Image.FromFile("icons/file/icons8-save-as-50.png");
            this.btnSaveAs.BackgroundImageLayout = ImageLayout.Stretch;
            //this.btnSaveAs.Text = "📋";
            this.btnSaveAs.UseVisualStyleBackColor = false;
            this.btnSaveAs.Click += new System.EventHandler(this.BtnSaveAs_Click);
            // 
            // btnPrint
            // 
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnPrint.FlatAppearance.BorderSize = 0;
            this.btnPrint.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnPrint.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPrint.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnPrint.Location = new System.Drawing.Point(100, 0);
            this.btnPrint.Margin = new System.Windows.Forms.Padding(1);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(20, 20);
            this.btnPrint.TabIndex = 5;
            this.btnPrint.BackColor = Color.Transparent;
            this.btnPrint.BackgroundImage = Image.FromFile("icons/file/icons8-export-50.png");
            this.btnPrint.BackgroundImageLayout = ImageLayout.Stretch;
            //this.btnPrint.Text = "🖨";
            this.btnPrint.UseVisualStyleBackColor = false;
            this.btnPrint.Click += new System.EventHandler(this.BtnPrint_Click);
            // 
            // btnShare
            // 
            //this.btnShare.BackColor = System.Drawing.Color.Transparent;
            //this.btnShare.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            //this.btnShare.FlatAppearance.BorderSize = 0;
            //this.btnShare.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            //this.btnShare.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            //this.btnShare.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            //this.btnShare.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            //this.btnShare.Location = new System.Drawing.Point(125, 0);
            //this.btnShare.Margin = new System.Windows.Forms.Padding(1);
            //this.btnShare.Name = "btnShare";
            //this.btnShare.Size = new System.Drawing.Size(20, 20);
            //this.btnShare.TabIndex = 6;
            //this.btnShare.BackColor = Color.Transparent;
            //this.btnShare.BackgroundImage = Image.FromFile("icons/file/icons8-share-48.png");
            //this.btnShare.BackgroundImageLayout = ImageLayout.Stretch;
            ////this.btnShare.Text = "🤝";
            //this.btnShare.UseVisualStyleBackColor = false;
            //this.btnShare.Click += new System.EventHandler(this.BtnShare_Click);
            // 
            // divider1
            // 
            this.divider1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.divider1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(225)))), ((int)(((byte)(225)))), ((int)(((byte)(225)))));
            this.divider1.Location = new System.Drawing.Point(3, 45);
            this.divider1.Name = "divider1";
            this.divider1.Size = new System.Drawing.Size(287, 1);
            this.divider1.TabIndex = 10;
            // 
            // wizardsPanel
            // 
            this.wizardsPanel.AutoSize = true;
            this.wizardsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.wizardsPanel.BackColor = System.Drawing.Color.White;
            this.wizardsPanel.Controls.Add(this.lblWizardsSection);
            this.wizardsPanel.Controls.Add(this.wizardButtonsPanel);
            this.wizardsPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.wizardsPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.wizardsPanel.Location = new System.Drawing.Point(3, 52);
            this.wizardsPanel.Name = "wizardsPanel";
            this.wizardsPanel.Size = new System.Drawing.Size(287, 38);
            this.wizardsPanel.TabIndex = 1;
            // 
            // lblWizardsSection
            // 
            this.lblWizardsSection.AutoSize = true;
            this.lblWizardsSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblWizardsSection.ForeColor = System.Drawing.Color.Gray;
            this.lblWizardsSection.Location = new System.Drawing.Point(3, 0);
            this.lblWizardsSection.Name = "lblWizardsSection";
            this.lblWizardsSection.Size = new System.Drawing.Size(48, 13);
            this.lblWizardsSection.TabIndex = 0;
            this.lblWizardsSection.Text = "Wizards";
            // 
            // wizardButtonsPanel
            // 
            this.wizardButtonsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.wizardButtonsPanel.AutoSize = true;
            this.wizardButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.wizardButtonsPanel.Controls.Add(this.btnAgenda);
            this.wizardButtonsPanel.Controls.Add(this.btnMaster);
            this.wizardButtonsPanel.Controls.Add(this.btnElement);
            this.wizardButtonsPanel.Controls.Add(this.btnText);
            this.wizardButtonsPanel.Controls.Add(this.btnFormat);
            this.wizardButtonsPanel.Controls.Add(this.btnMap);
            this.wizardButtonsPanel.Location = new System.Drawing.Point(0, 13);
            this.wizardButtonsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.wizardButtonsPanel.Name = "wizardButtonsPanel";
            this.wizardButtonsPanel.Size = new System.Drawing.Size(150, 25);
            this.wizardButtonsPanel.TabIndex = 1;
            // 
            // btnAgenda
            // 
            this.btnAgenda.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAgenda.FlatAppearance.BorderSize = 0;
            this.btnAgenda.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAgenda.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAgenda.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAgenda.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnAgenda.Location = new System.Drawing.Point(0, 0);
            this.btnAgenda.Margin = new System.Windows.Forms.Padding(0);
            this.btnAgenda.Name = "btnAgenda";
            this.btnAgenda.Size = new System.Drawing.Size(65, 20);
            this.btnAgenda.TabIndex = 1;
            // this.btnAgenda.Text = "📋";
            this.btnAgenda.BackgroundImage = Image.FromFile("icons/wizzards/agenda.png");
            this.btnAgenda.BackgroundImageLayout = ImageLayout.Stretch;
            this.btnAgenda.UseVisualStyleBackColor = false;
            this.btnAgenda.Click += new System.EventHandler(this.BtnAgenda_Click);
            // 
            // btnMaster
            // 
            this.btnMaster.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnMaster.FlatAppearance.BorderSize = 0;
            this.btnMaster.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnMaster.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnMaster.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMaster.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnMaster.Location = new System.Drawing.Point(25, 0);
            this.btnMaster.Margin = new System.Windows.Forms.Padding(0);
            this.btnMaster.Name = "btnMaster";
            this.btnMaster.Size = new System.Drawing.Size(65, 20);
            this.btnMaster.TabIndex = 2;
            // this.btnMaster.Text = "🎨";
            this.btnMaster.BackgroundImage = Image.FromFile("icons/wizzards/master.png");
            this.btnMaster.BackgroundImageLayout = ImageLayout.Stretch;
            this.btnMaster.UseVisualStyleBackColor = false;
            this.btnMaster.Click += new System.EventHandler(this.BtnMaster_Click);
            // 
            // btnElement
            // 
            this.btnElement.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnElement.FlatAppearance.BorderSize = 0;
            this.btnElement.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnElement.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnElement.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnElement.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnElement.Location = new System.Drawing.Point(50, 0);
            this.btnElement.Margin = new System.Windows.Forms.Padding(0);
            this.btnElement.Name = "btnElement";
            this.btnElement.Size = new System.Drawing.Size(65, 20);
            this.btnElement.TabIndex = 3;
            // this.btnElement.Text = "🧩";
            this.btnElement.BackgroundImage = Image.FromFile("icons/wizzards/element.png");
            this.btnElement.BackgroundImageLayout = ImageLayout.Stretch;
            this.btnElement.UseVisualStyleBackColor = false;
            this.btnElement.Click += new System.EventHandler(this.BtnElement_Click);
            // 
            // btnText
            // 
            this.btnText.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnText.FlatAppearance.BorderSize = 0;
            this.btnText.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnText.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnText.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnText.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnText.Location = new System.Drawing.Point(75, 0);
            this.btnText.Margin = new System.Windows.Forms.Padding(0);
            this.btnText.Name = "btnText";
            this.btnText.Size = new System.Drawing.Size(65, 20);
            this.btnText.TabIndex = 4;
            // this.btnText.Text = "✏️";
            this.btnText.BackgroundImage = Image.FromFile("icons/wizzards/text.png");
            this.btnText.BackgroundImageLayout = ImageLayout.Stretch;
            this.btnText.UseVisualStyleBackColor = false;
            this.btnText.Click += new System.EventHandler(this.BtnTextWizard_Click);
            // 
            // btnFormat
            // 
            this.btnFormat.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnFormat.FlatAppearance.BorderSize = 0;
            this.btnFormat.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnFormat.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnFormat.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFormat.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnFormat.Location = new System.Drawing.Point(100, 0);
            this.btnFormat.Margin = new System.Windows.Forms.Padding(0);
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Size = new System.Drawing.Size(65, 20);
            this.btnFormat.TabIndex = 5;
            // this.btnFormat.Text = "🎯";
            this.btnFormat.BackgroundImage = Image.FromFile("icons/wizzards/format.png");
            this.btnFormat.BackgroundImageLayout = ImageLayout.Stretch;
            this.btnFormat.UseVisualStyleBackColor = false;
            this.btnFormat.Click += new System.EventHandler(this.BtnFormat_Click);
            // 
            // btnMap
            // 
            this.btnMap.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnMap.FlatAppearance.BorderSize = 0;
            this.btnMap.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnMap.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnMap.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMap.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnMap.Location = new System.Drawing.Point(125, 0);
            this.btnMap.Margin = new System.Windows.Forms.Padding(0);
            this.btnMap.Name = "btnMap";
            this.btnMap.Size = new System.Drawing.Size(65, 20);
            this.btnMap.TabIndex = 6;
            // this.btnMap.Text = "🗺️";
            this.btnMap.BackgroundImage = Image.FromFile("icons/wizzards/map.png");
            this.btnMap.BackgroundImageLayout = ImageLayout.Stretch;
            this.btnMap.UseVisualStyleBackColor = false;
            this.btnMap.Click += new System.EventHandler(this.BtnMap_Click);
            // 
            // divider2
            // 
            this.divider2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.divider2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(225)))), ((int)(((byte)(225)))), ((int)(((byte)(225)))));
            this.divider2.Location = new System.Drawing.Point(3, 98);
            this.divider2.Name = "divider2";
            this.divider2.Size = new System.Drawing.Size(287, 1);
            this.divider2.TabIndex = 11;
            // 
            // smartElementsPanel
            // 
            this.smartElementsPanel.AutoSize = true;
            this.smartElementsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.smartElementsPanel.BackColor = System.Drawing.Color.White;
            this.smartElementsPanel.Controls.Add(this.lblSmartElementsSection);
            this.smartElementsPanel.Controls.Add(this.smartElementsButtonsPanel);
            this.smartElementsPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.smartElementsPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.smartElementsPanel.Location = new System.Drawing.Point(3, 105);
            this.smartElementsPanel.Name = "smartElementsPanel";
            this.smartElementsPanel.Size = new System.Drawing.Size(287, 58);
            this.smartElementsPanel.TabIndex = 2;
            // 
            // lblSmartElementsSection
            // 
            this.lblSmartElementsSection.AutoSize = true;
            this.lblSmartElementsSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblSmartElementsSection.ForeColor = System.Drawing.Color.Gray;
            this.lblSmartElementsSection.Location = new System.Drawing.Point(3, 0);
            this.lblSmartElementsSection.Name = "lblSmartElementsSection";
            this.lblSmartElementsSection.Size = new System.Drawing.Size(85, 13);
            this.lblSmartElementsSection.TabIndex = 0;
            this.lblSmartElementsSection.Text = "Smart Elements";
            // 
            // smartElementsButtonsPanel
            // 
            this.smartElementsButtonsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.smartElementsButtonsPanel.AutoSize = true;
            this.smartElementsButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.smartElementsButtonsPanel.Controls.Add(this.btnChart);
            this.smartElementsButtonsPanel.Controls.Add(this.btnDiagram);
            this.smartElementsButtonsPanel.Controls.Add(this.btnTable);
            this.smartElementsButtonsPanel.Controls.Add(this.btnMatrixTable);
            this.smartElementsButtonsPanel.Controls.Add(this.btnStickyNote);
            this.smartElementsButtonsPanel.Controls.Add(this.btnCitation);
            this.smartElementsButtonsPanel.Controls.Add(this.btnStandardObjects);
            this.smartElementsButtonsPanel.Location = new System.Drawing.Point(0, 13);
            this.smartElementsButtonsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.smartElementsButtonsPanel.Name = "smartElementsButtonsPanel";
            this.smartElementsButtonsPanel.Size = new System.Drawing.Size(287, 45);
            this.smartElementsButtonsPanel.TabIndex = 1;
            // 
            // btnChart
            // 
            this.btnChart.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnChart.FlatAppearance.BorderSize = 0;
            this.btnChart.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnChart.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnChart.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnChart.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnChart.Location = new System.Drawing.Point(5, 5);
            this.btnChart.Margin = new System.Windows.Forms.Padding(2);
            this.btnChart.Name = "btnChart";
            this.btnChart.Size = new System.Drawing.Size(20, 20);
            this.btnChart.TabIndex = 1;
            this.btnChart.BackgroundImage = Image.FromFile("icons/elements/icons8-chart-60.png");
            this.btnChart.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnChart.Text = "📊";
            this.btnChart.UseVisualStyleBackColor = false;
            this.btnChart.Click += new System.EventHandler(this.BtnChart_Click);
            // 
            // btnDiagram
            // 
            this.btnDiagram.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnDiagram.FlatAppearance.BorderSize = 0;
            this.btnDiagram.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnDiagram.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnDiagram.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDiagram.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnDiagram.Location = new System.Drawing.Point(30, 5);
            this.btnDiagram.Margin = new System.Windows.Forms.Padding(2);
            this.btnDiagram.Name = "btnDiagram";
            this.btnDiagram.Size = new System.Drawing.Size(20, 20);
            this.btnDiagram.TabIndex = 2;
            this.btnDiagram.BackgroundImage = Image.FromFile("icons/elements/icons8-color-palette-48.png");
            this.btnDiagram.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnDiagram.Text = "🎨";
            this.btnDiagram.UseVisualStyleBackColor = false;
            this.btnDiagram.Click += new System.EventHandler(this.BtnDiagram_Click);
            // 
            // btnTable
            // 
            this.btnTable.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnTable.FlatAppearance.BorderSize = 0;
            this.btnTable.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnTable.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnTable.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnTable.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnTable.Location = new System.Drawing.Point(55, 5);
            this.btnTable.Margin = new System.Windows.Forms.Padding(2);
            this.btnTable.Name = "btnTable";
            this.btnTable.Size = new System.Drawing.Size(20, 20);
            this.btnTable.TabIndex = 3;
            this.btnTable.BackgroundImage = Image.FromFile("icons/elements/icons8-table-50.png");
            this.btnTable.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnTable.Text = "📋";
            this.btnTable.UseVisualStyleBackColor = false;
            this.btnTable.Click += new System.EventHandler(this.BtnTable_Click);
            // 
            // btnMatrixTable   
            // 
            this.btnMatrixTable.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnMatrixTable.FlatAppearance.BorderSize = 0;
            this.btnMatrixTable.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnMatrixTable.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnMatrixTable.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMatrixTable.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnMatrixTable.Location = new System.Drawing.Point(80, 5);
            this.btnMatrixTable.Margin = new System.Windows.Forms.Padding(2);
            this.btnMatrixTable.Name = "btnMatrixTable";
            this.btnMatrixTable.Size = new System.Drawing.Size(20, 20);
            this.btnMatrixTable.TabIndex = 4;
            this.btnMatrixTable.BackgroundImage = Image.FromFile("icons/elements/icons8-matrix-50.png");
            this.btnMatrixTable.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnMatrixTable.Text = "🏢";
            this.btnMatrixTable.UseVisualStyleBackColor = false;
            this.btnMatrixTable.Click += new System.EventHandler(this.BtnMatrixTable_Click);
            // 
            // btnStickyNote
            // 
            this.btnStickyNote.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnStickyNote.FlatAppearance.BorderSize = 0;
            this.btnStickyNote.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnStickyNote.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnStickyNote.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnStickyNote.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnStickyNote.Location = new System.Drawing.Point(5, 30);
            this.btnStickyNote.Margin = new System.Windows.Forms.Padding(2);
            this.btnStickyNote.Name = "btnStickyNote";
            this.btnStickyNote.Size = new System.Drawing.Size(20, 20);
            this.btnStickyNote.TabIndex = 5;
            this.btnStickyNote.BackgroundImage = Image.FromFile("icons/elements/icons8-sticky-notes-50.png");
            this.btnStickyNote.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnStickyNote.Text = "📝";
            this.btnStickyNote.UseVisualStyleBackColor = false;
            this.btnStickyNote.Click += new System.EventHandler(this.BtnStickyNote_Click);
            // 
            // btnCitation
            // 
            this.btnCitation.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnCitation.FlatAppearance.BorderSize = 0;
            this.btnCitation.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnCitation.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnCitation.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCitation.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnCitation.Location = new System.Drawing.Point(30, 30);
            this.btnCitation.Margin = new System.Windows.Forms.Padding(2);
            this.btnCitation.Name = "btnCitation";
            this.btnCitation.Size = new System.Drawing.Size(20, 20);
            this.btnCitation.TabIndex = 6;
            this.btnCitation.BackgroundImage = Image.FromFile("icons/elements/icons8-get-quote-30.png");
            this.btnCitation.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnCitation.Text = "📑";
            this.btnCitation.UseVisualStyleBackColor = false;
            this.btnCitation.Click += new System.EventHandler(this.BtnCitation_Click);
            // 
            // btnStandardObjects
            // 
            this.btnStandardObjects.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnStandardObjects.FlatAppearance.BorderSize = 0;
            this.btnStandardObjects.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnStandardObjects.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnStandardObjects.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnStandardObjects.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnStandardObjects.Location = new System.Drawing.Point(55, 30);
            this.btnStandardObjects.Margin = new System.Windows.Forms.Padding(2);
            this.btnStandardObjects.Name = "btnStandardObjects";
            this.btnStandardObjects.Size = new System.Drawing.Size(20, 20);
            this.btnStandardObjects.TabIndex = 7;
            this.btnStandardObjects.BackgroundImage = Image.FromFile("icons/elements/icons8-object-50.png");
            this.btnStandardObjects.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnStandardObjects.Text = "🗂️";
            this.btnStandardObjects.UseVisualStyleBackColor = false;
            this.btnStandardObjects.Click += new System.EventHandler(this.BtnStandardObjects_Click);
            // 
            // divider3
            // 
            this.divider3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.divider3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(225)))), ((int)(((byte)(225)))), ((int)(((byte)(225)))));
            this.divider3.Location = new System.Drawing.Point(3, 154);
            this.divider3.Margin = new System.Windows.Forms.Padding(3, 8, 3, 8);
            this.divider3.Name = "divider3";
            this.divider3.Size = new System.Drawing.Size(287, 1);
            this.divider3.TabIndex = 12;
            // 
            // positionPanel
            // 
            this.positionPanel.AutoSize = true;
            this.positionPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.positionPanel.BackColor = System.Drawing.Color.White;
            this.positionPanel.Controls.Add(this.lblPositionSection);
            this.positionPanel.Controls.Add(this.positionButtonsPanel);
            this.positionPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.positionPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.positionPanel.Location = new System.Drawing.Point(3, 166);
            this.positionPanel.Name = "positionPanel";
            this.positionPanel.Size = new System.Drawing.Size(287, 38);
            this.positionPanel.TabIndex = 3;
            // 
            // lblPositionSection
            // 
            this.lblPositionSection.AutoSize = true;
            this.lblPositionSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblPositionSection.ForeColor = System.Drawing.Color.Gray;
            this.lblPositionSection.Location = new System.Drawing.Point(3, 0);
            this.lblPositionSection.Name = "lblPositionSection";
            this.lblPositionSection.Size = new System.Drawing.Size(49, 13);
            this.lblPositionSection.TabIndex = 0;
            this.lblPositionSection.Text = "Position";
            // 
            // positionButtonsPanel
            // 
            this.positionButtonsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.positionButtonsPanel.AutoSize = true;
            this.positionButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.positionButtonsPanel.Controls.Add(this.btnAlignLeft);
            this.positionButtonsPanel.Controls.Add(this.btnAlignCenter);
            this.positionButtonsPanel.Controls.Add(this.btnAlignRight);
            this.positionButtonsPanel.Controls.Add(this.btnAlignTop);
            this.positionButtonsPanel.Controls.Add(this.btnAlignBottom);
            this.positionButtonsPanel.Controls.Add(this.btnAlignMiddle);
            this.positionButtonsPanel.Controls.Add(this.btnDockLeft);
            this.positionButtonsPanel.Controls.Add(this.btnDockRight);
            this.positionButtonsPanel.Controls.Add(this.btnDockTop);
            this.positionButtonsPanel.Controls.Add(this.btnDockBottom);
            this.positionButtonsPanel.Controls.Add(this.btnDistribute);
            this.positionButtonsPanel.Controls.Add(this.btnDistributeHorizontal);
            this.positionButtonsPanel.Controls.Add(this.btnDistributeVertical);
            this.positionButtonsPanel.Controls.Add(this.btnMatchBoth);
            this.positionButtonsPanel.Controls.Add(this.btnMatchHeight);
            this.positionButtonsPanel.Controls.Add(this.btnMatchWidth);
            this.positionButtonsPanel.Controls.Add(this.btnMakeVertical);
            this.positionButtonsPanel.Controls.Add(this.btnMakeHorizontal);
            this.positionButtonsPanel.Controls.Add(this.btnSwapLocations);
            this.positionButtonsPanel.Controls.Add(this.btnGoldenCanon);
            this.positionButtonsPanel.Controls.Add(this.btnAlignMatrix);
            this.positionButtonsPanel.Controls.Add(this.btnSliceShape);
            this.positionButtonsPanel.Controls.Add(this.btnDuplicateRight);
            this.positionButtonsPanel.Controls.Add(this.btnCenterTopLeft);
            this.positionButtonsPanel.Controls.Add(this.btnSavePosition);
            this.positionButtonsPanel.Controls.Add(this.btnApplyPosition);
            this.positionButtonsPanel.Location = new System.Drawing.Point(0, 13);
            this.positionButtonsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.positionButtonsPanel.Name = "positionButtonsPanel";
            this.positionButtonsPanel.Size = new System.Drawing.Size(250, 25);
            this.positionButtonsPanel.TabIndex = 1;
            // 
            // btnAlignLeft
            // 
            this.btnAlignLeft.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignLeft.FlatAppearance.BorderSize = 0;
            this.btnAlignLeft.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignLeft.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignLeft.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignLeft.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnAlignLeft.Location = new System.Drawing.Point(0, 0);
            this.btnAlignLeft.Margin = new System.Windows.Forms.Padding(1);
            this.btnAlignLeft.Name = "btnAlignLeft";
            this.btnAlignLeft.Size = new System.Drawing.Size(20, 20);
            this.btnAlignLeft.TabIndex = 1;
            this.btnAlignLeft.BackgroundImage = Image.FromFile("icons/position/icons8-align-left-64.png");
            this.btnAlignLeft.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnAlignLeft.Text = "←";
            this.btnAlignLeft.UseVisualStyleBackColor = false;
            this.btnAlignLeft.Click += new System.EventHandler(this.BtnAlignLeft_Click);
            // 
            // btnAlignCenter
            // 
            this.btnAlignCenter.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignCenter.FlatAppearance.BorderSize = 0;
            this.btnAlignCenter.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignCenter.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignCenter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignCenter.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnAlignCenter.Location = new System.Drawing.Point(25, 0);
            this.btnAlignCenter.Margin = new System.Windows.Forms.Padding(1);
            this.btnAlignCenter.Name = "btnAlignCenter";
            this.btnAlignCenter.Size = new System.Drawing.Size(20, 20);
            this.btnAlignCenter.TabIndex = 2;
            this.btnAlignCenter.BackgroundImage = Image.FromFile("icons/position/icons8-align-center-64.png");
            this.btnAlignCenter.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnAlignCenter.Text = "●";
            this.btnAlignCenter.UseVisualStyleBackColor = false;
            this.btnAlignCenter.Click += new System.EventHandler(this.BtnAlignCenter_Click);
            // 
            // btnAlignRight
            // 
            this.btnAlignRight.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignRight.FlatAppearance.BorderSize = 0;
            this.btnAlignRight.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignRight.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignRight.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignRight.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnAlignRight.Location = new System.Drawing.Point(50, 0);
            this.btnAlignRight.Margin = new System.Windows.Forms.Padding(1);
            this.btnAlignRight.Name = "btnAlignRight";
            this.btnAlignRight.Size = new System.Drawing.Size(20, 20);
            this.btnAlignRight.TabIndex = 3;
            this.btnAlignRight.BackgroundImage = Image.FromFile("icons/position/icons8-align-right-64.png");
            this.btnAlignRight.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnAlignRight.Text = "→";
            this.btnAlignRight.UseVisualStyleBackColor = false;
            this.btnAlignRight.Click += new System.EventHandler(this.BtnAlignRight_Click);
            // btnAlignTop
            this.btnAlignTop.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignTop.FlatAppearance.BorderSize = 0;
            this.btnAlignTop.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignTop.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignTop.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignTop.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnAlignTop.Location = new System.Drawing.Point(75, 0);
            this.btnAlignTop.Margin = new System.Windows.Forms.Padding(1);
            this.btnAlignTop.Name = "btnAlignTop";
            this.btnAlignTop.Size = new System.Drawing.Size(20, 20);
            this.btnAlignTop.TabIndex = 4;
            this.btnAlignTop.BackgroundImage = Image.FromFile("icons/position/icons8-align-top-64.png");
            this.btnAlignTop.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnAlignTop.Text = "↑";
            this.btnAlignTop.UseVisualStyleBackColor = false;
            this.btnAlignTop.Click += new System.EventHandler(this.BtnAlignTop_Click);
            // btnAlignBottom
            this.btnAlignBottom.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignBottom.FlatAppearance.BorderSize = 0;
            this.btnAlignBottom.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignBottom.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignBottom.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignBottom.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnAlignBottom.Location = new System.Drawing.Point(100, 0);
            this.btnAlignBottom.Margin = new System.Windows.Forms.Padding(1);
            this.btnAlignBottom.Name = "btnAlignBottom";
            this.btnAlignBottom.Size = new System.Drawing.Size(20, 20);
            this.btnAlignBottom.TabIndex = 5;
            this.btnAlignBottom.BackgroundImage = Image.FromFile("icons/position/icons8-align-bottom-64.png");
            this.btnAlignBottom.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnAlignBottom.Text = "↓";
            this.btnAlignBottom.UseVisualStyleBackColor = false;
            this.btnAlignBottom.Click += new System.EventHandler(this.BtnAlignBottom_Click);
            // btnAlignMiddle
            this.btnAlignMiddle.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignMiddle.FlatAppearance.BorderSize = 0;
            this.btnAlignMiddle.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignMiddle.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignMiddle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignMiddle.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnAlignMiddle.Location = new System.Drawing.Point(125, 0);
            this.btnAlignMiddle.Margin = new System.Windows.Forms.Padding(1);
            this.btnAlignMiddle.Name = "btnAlignMiddle";
            this.btnAlignMiddle.Size = new System.Drawing.Size(20, 20);
            this.btnAlignMiddle.TabIndex = 6;
            this.btnAlignMiddle.BackgroundImage = Image.FromFile("icons/position/icons8-align-center-64.png");
            this.btnAlignMiddle.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnAlignMiddle.Text = "∥";
            this.btnAlignMiddle.UseVisualStyleBackColor = false;
            this.btnAlignMiddle.Click += new System.EventHandler(this.BtnAlignMiddle_Click);
                         // btnDockLeft
             this.btnDockLeft.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
             this.btnDockLeft.FlatAppearance.BorderSize = 0;
             this.btnDockLeft.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
             this.btnDockLeft.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
             this.btnDockLeft.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
             this.btnDockLeft.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
             this.btnDockLeft.Location = new System.Drawing.Point(150, 0);
             this.btnDockLeft.Margin = new System.Windows.Forms.Padding(1);
             this.btnDockLeft.Name = "btnDockLeft";
             this.btnDockLeft.Size = new System.Drawing.Size(20, 20);
             this.btnDockLeft.TabIndex = 7;
             this.btnDockLeft.BackgroundImage = Image.FromFile("icons/position/icons8-align-left-64.png");
             this.btnDockLeft.BackgroundImageLayout = ImageLayout.Stretch;
             // this.btnDockLeft.Text = "🗕";
             this.btnDockLeft.UseVisualStyleBackColor = false;
             this.btnDockLeft.Click += new System.EventHandler(this.BtnDockLeft_Click);
                         // btnDockRight
             this.btnDockRight.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
             this.btnDockRight.FlatAppearance.BorderSize = 0;
             this.btnDockRight.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
             this.btnDockRight.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
             this.btnDockRight.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
             this.btnDockRight.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
             this.btnDockRight.Location = new System.Drawing.Point(175, 0);
             this.btnDockRight.Margin = new System.Windows.Forms.Padding(1);
             this.btnDockRight.Name = "btnDockRight";
             this.btnDockRight.Size = new System.Drawing.Size(20, 20);
             this.btnDockRight.TabIndex = 8;
             this.btnDockRight.BackgroundImage = Image.FromFile("icons/position/icons8-align-right-64.png");
             this.btnDockRight.BackgroundImageLayout = ImageLayout.Stretch;
             // this.btnDockRight.Text = "🗖";
             this.btnDockRight.UseVisualStyleBackColor = false;
             this.btnDockRight.Click += new System.EventHandler(this.BtnDockRight_Click);
                         // btnDockTop
             this.btnDockTop.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
             this.btnDockTop.FlatAppearance.BorderSize = 0;
             this.btnDockTop.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
             this.btnDockTop.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
             this.btnDockTop.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
             this.btnDockTop.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
             this.btnDockTop.Location = new System.Drawing.Point(200, 0);
             this.btnDockTop.Margin = new System.Windows.Forms.Padding(1);
             this.btnDockTop.Name = "btnDockTop";
             this.btnDockTop.Size = new System.Drawing.Size(20, 20);
             this.btnDockTop.TabIndex = 9;
             this.btnDockTop.BackgroundImage = Image.FromFile("icons/position/icons8-align-top-64.png");
             this.btnDockTop.BackgroundImageLayout = ImageLayout.Stretch;
             // this.btnDockTop.Text = "🗂";
             this.btnDockTop.UseVisualStyleBackColor = false;
             this.btnDockTop.Click += new System.EventHandler(this.BtnDockTop_Click);
                         // btnDockBottom
             this.btnDockBottom.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
             this.btnDockBottom.FlatAppearance.BorderSize = 0;
             this.btnDockBottom.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
             this.btnDockBottom.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
             this.btnDockBottom.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
             this.btnDockBottom.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
             this.btnDockBottom.Location = new System.Drawing.Point(225, 0);
             this.btnDockBottom.Margin = new System.Windows.Forms.Padding(1);
             this.btnDockBottom.Name = "btnDockBottom";
             this.btnDockBottom.Size = new System.Drawing.Size(20, 20);
             this.btnDockBottom.TabIndex = 10;
             this.btnDockBottom.BackgroundImage = Image.FromFile("icons/position/icons8-align-bottom-64.png");
             this.btnDockBottom.BackgroundImageLayout = ImageLayout.Stretch;
             // this.btnDockBottom.Text = "🗃";
             this.btnDockBottom.UseVisualStyleBackColor = false;
             this.btnDockBottom.Click += new System.EventHandler(this.BtnDockBottom_Click);
            // btnDistribute
            this.btnDistribute.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnDistribute.FlatAppearance.BorderSize = 0;
            this.btnDistribute.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnDistribute.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnDistribute.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDistribute.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnDistribute.Location = new System.Drawing.Point(250, 0);
            this.btnDistribute.Margin = new System.Windows.Forms.Padding(1);
            this.btnDistribute.Name = "btnDistribute";
            this.btnDistribute.Size = new System.Drawing.Size(20, 20);
            this.btnDistribute.TabIndex = 11;
            this.btnDistribute.BackgroundImage = Image.FromFile("icons/position/icons8-align-justify-64.png");
            this.btnDistribute.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnDistribute.Text = "≡";
            this.btnDistribute.UseVisualStyleBackColor = false;
            this.btnDistribute.Click += new System.EventHandler(this.BtnDistribute_Click);
            // btnDistributeHorizontal
            this.btnDistributeHorizontal.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnDistributeHorizontal.FlatAppearance.BorderSize = 0;
            this.btnDistributeHorizontal.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnDistributeHorizontal.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnDistributeHorizontal.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDistributeHorizontal.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnDistributeHorizontal.Location = new System.Drawing.Point(275, 0);
            this.btnDistributeHorizontal.Margin = new System.Windows.Forms.Padding(1);
            this.btnDistributeHorizontal.Name = "btnDistributeHorizontal";
            this.btnDistributeHorizontal.Size = new System.Drawing.Size(20, 20);
            this.btnDistributeHorizontal.TabIndex = 12;
            this.btnDistributeHorizontal.BackgroundImage = Image.FromFile("icons/position/icons8-align-center-64.png");
            this.btnDistributeHorizontal.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnDistributeHorizontal.Text = "⇔";
            this.btnDistributeHorizontal.UseVisualStyleBackColor = false;
            this.btnDistributeHorizontal.Click += new System.EventHandler(this.BtnDistributeHorizontal_Click);
            // btnDistributeVertical
            this.btnDistributeVertical.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnDistributeVertical.FlatAppearance.BorderSize = 0;
            this.btnDistributeVertical.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnDistributeVertical.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnDistributeVertical.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDistributeVertical.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnDistributeVertical.Location = new System.Drawing.Point(300, 0);
            this.btnDistributeVertical.Margin = new System.Windows.Forms.Padding(1);
            this.btnDistributeVertical.Name = "btnDistributeVertical";
            this.btnDistributeVertical.Size = new System.Drawing.Size(20, 20);
            this.btnDistributeVertical.TabIndex = 13;
            this.btnDistributeVertical.BackgroundImage = Image.FromFile("icons/position/icons8-align-justify-64.png");
            this.btnDistributeVertical.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnDistributeVertical.Text = "⇕";
            this.btnDistributeVertical.UseVisualStyleBackColor = false;
            this.btnDistributeVertical.Click += new System.EventHandler(this.BtnDistributeVertical_Click);
            // btnMatchBoth
            this.btnMatchBoth.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnMatchBoth.FlatAppearance.BorderSize = 0;
            this.btnMatchBoth.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnMatchBoth.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnMatchBoth.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMatchBoth.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnMatchBoth.Location = new System.Drawing.Point(250, 0);
            this.btnMatchBoth.Margin = new System.Windows.Forms.Padding(1);
            this.btnMatchBoth.Name = "btnMatchBoth";
            this.btnMatchBoth.Size = new System.Drawing.Size(20, 20);
            this.btnMatchBoth.TabIndex = 14;
            this.btnMatchBoth.BackgroundImage = Image.FromFile("icons/position/icons8-enlarge-50.png");
            this.btnMatchBoth.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnMatchBoth.Text = "📏";
            this.btnMatchBoth.UseVisualStyleBackColor = false;
            this.btnMatchBoth.Click += new System.EventHandler(this.BtnMatchBoth_Click);
            // btnMatchHeight
            this.btnMatchHeight.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnMatchHeight.FlatAppearance.BorderSize = 0;
            this.btnMatchHeight.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnMatchHeight.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnMatchHeight.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMatchHeight.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnMatchHeight.Location = new System.Drawing.Point(275, 0);
            this.btnMatchHeight.Margin = new System.Windows.Forms.Padding(1);
            this.btnMatchHeight.Name = "btnMatchHeight";
            this.btnMatchHeight.Size = new System.Drawing.Size(20, 20);
            this.btnMatchHeight.TabIndex = 15;
            this.btnMatchHeight.BackgroundImage = Image.FromFile("icons/position/icons8-height-50.png");
            this.btnMatchHeight.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnMatchHeight.Text = "↕️";
            this.btnMatchHeight.UseVisualStyleBackColor = false;
            this.btnMatchHeight.Click += new System.EventHandler(this.BtnMatchHeight_Click);
            // 
            // btnMatchWidth
            // 
            this.btnMatchWidth.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnMatchWidth.FlatAppearance.BorderSize = 0;
            this.btnMatchWidth.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnMatchWidth.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnMatchWidth.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMatchWidth.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnMatchWidth.Location = new System.Drawing.Point(300, 0);
            this.btnMatchWidth.Margin = new System.Windows.Forms.Padding(1);
            this.btnMatchWidth.Name = "btnMatchWidth";
            this.btnMatchWidth.Size = new System.Drawing.Size(20, 20);
            this.btnMatchWidth.TabIndex = 16;
            this.btnMatchWidth.BackgroundImage = Image.FromFile("icons/position/icons8-width-50.png");
            this.btnMatchWidth.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnMatchWidth.Text = "↔️";
            this.btnMatchWidth.UseVisualStyleBackColor = false;
            this.btnMatchWidth.Click += new System.EventHandler(this.BtnMatchWidth_Click);
            // 
            // btnMakeVertical
            // 
            this.btnMakeVertical.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnMakeVertical.FlatAppearance.BorderSize = 0;
            this.btnMakeVertical.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnMakeVertical.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnMakeVertical.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMakeVertical.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnMakeVertical.Location = new System.Drawing.Point(325, 0);
            this.btnMakeVertical.Margin = new System.Windows.Forms.Padding(0);
            this.btnMakeVertical.Name = "btnMakeVertical";
            this.btnMakeVertical.Size = new System.Drawing.Size(20, 20);
            this.btnMakeVertical.TabIndex = 17;
            this.btnMakeVertical.BackgroundImage = Image.FromFile("icons/position/icons8-rotate-left-48.png");
            this.btnMakeVertical.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnMakeVertical.Text = "↻";
            this.btnMakeVertical.UseVisualStyleBackColor = false;
            this.btnMakeVertical.Click += new System.EventHandler(this.BtnMakeVertical_Click);
            // 
            // btnMakeHorizontal
            // 
            this.btnMakeHorizontal.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnMakeHorizontal.FlatAppearance.BorderSize = 0;
            this.btnMakeHorizontal.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnMakeHorizontal.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnMakeHorizontal.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMakeHorizontal.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnMakeHorizontal.Location = new System.Drawing.Point(350, 0);
            this.btnMakeHorizontal.Margin = new System.Windows.Forms.Padding(1);
            this.btnMakeHorizontal.Name = "btnMakeHorizontal";
            this.btnMakeHorizontal.Size = new System.Drawing.Size(20, 20);
            this.btnMakeHorizontal.TabIndex = 18;
            this.btnMakeHorizontal.BackgroundImage = Image.FromFile("icons/position/icons8-rotate-right-48.png");
            this.btnMakeHorizontal.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnMakeHorizontal.Text = "↺";
            this.btnMakeHorizontal.UseVisualStyleBackColor = false;
            this.btnMakeHorizontal.Click += new System.EventHandler(this.BtnMakeHorizontal_Click);
            // 
            // btnSwapLocations
            // 
            this.btnSwapLocations.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnSwapLocations.FlatAppearance.BorderSize = 0;
            this.btnSwapLocations.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnSwapLocations.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnSwapLocations.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSwapLocations.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnSwapLocations.Location = new System.Drawing.Point(375, 0);
            this.btnSwapLocations.Margin = new System.Windows.Forms.Padding(1);
            this.btnSwapLocations.Name = "btnSwapLocations";
            this.btnSwapLocations.Size = new System.Drawing.Size(25, 25);
            this.btnSwapLocations.TabIndex = 19;
            this.btnSwapLocations.BackgroundImage = Image.FromFile("icons/position/icons8-swap-50.png");
            this.btnSwapLocations.BackgroundImageLayout = ImageLayout.Stretch;
            // this.btnSwapLocations.Text = "⇄";
            this.btnSwapLocations.UseVisualStyleBackColor = false;
            this.btnSwapLocations.Click += new System.EventHandler(this.BtnSwapLocations_Click);
             // btnGoldenCanon
             this.btnGoldenCanon.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
             this.btnGoldenCanon.FlatAppearance.BorderSize = 0;
             this.btnGoldenCanon.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
             this.btnGoldenCanon.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
             this.btnGoldenCanon.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
             this.btnGoldenCanon.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
             this.btnGoldenCanon.Location = new System.Drawing.Point(400, 0);
             this.btnGoldenCanon.Margin = new System.Windows.Forms.Padding(1);
             this.btnGoldenCanon.Name = "btnGoldenCanon";
             this.btnGoldenCanon.Size = new System.Drawing.Size(20, 20);
             this.btnGoldenCanon.TabIndex = 20;
             this.btnGoldenCanon.BackgroundImage = Image.FromFile("icons/position/icons8-swap-50.png");
             this.btnGoldenCanon.BackgroundImageLayout = ImageLayout.Stretch;
             this.btnGoldenCanon.UseVisualStyleBackColor = false;
             this.btnGoldenCanon.Click += new System.EventHandler(this.BtnGoldenCanon_Click);
             // btnAlignMatrix
             this.btnAlignMatrix.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
             this.btnAlignMatrix.FlatAppearance.BorderSize = 0;
             this.btnAlignMatrix.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
             this.btnAlignMatrix.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
             this.btnAlignMatrix.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
             this.btnAlignMatrix.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
             this.btnAlignMatrix.Location = new System.Drawing.Point(425, 0);
             this.btnAlignMatrix.Margin = new System.Windows.Forms.Padding(1);
             this.btnAlignMatrix.Name = "btnAlignMatrix";
             this.btnAlignMatrix.Size = new System.Drawing.Size(20, 20);
             this.btnAlignMatrix.TabIndex = 21;
             this.btnAlignMatrix.BackgroundImage = Image.FromFile("icons/position/icons8-matrix-50.png");
             this.btnAlignMatrix.BackgroundImageLayout = ImageLayout.Stretch;
             this.btnAlignMatrix.UseVisualStyleBackColor = false;
             this.btnAlignMatrix.Click += new System.EventHandler(this.BtnAlignMatrix_Click);
             // btnSliceShape
             this.btnSliceShape.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
             this.btnSliceShape.FlatAppearance.BorderSize = 0;
             this.btnSliceShape.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
             this.btnSliceShape.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
             this.btnSliceShape.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
             this.btnSliceShape.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
             this.btnSliceShape.Location = new System.Drawing.Point(450, 0);
             this.btnSliceShape.Margin = new System.Windows.Forms.Padding(1);
             this.btnSliceShape.Name = "btnSliceShape";
             this.btnSliceShape.Size = new System.Drawing.Size(20, 20);
             this.btnSliceShape.TabIndex = 22;
             this.btnSliceShape.BackgroundImage = Image.FromFile("icons/position/icons8-slice-50.png");
             this.btnSliceShape.BackgroundImageLayout = ImageLayout.Stretch;
             this.btnSliceShape.UseVisualStyleBackColor = false;
             this.btnSliceShape.Click += new System.EventHandler(this.BtnSliceShape_Click);
             // btnDuplicateRight
             this.btnDuplicateRight.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
             this.btnDuplicateRight.FlatAppearance.BorderSize = 0;
             this.btnDuplicateRight.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
             this.btnDuplicateRight.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
             this.btnDuplicateRight.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
             this.btnDuplicateRight.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
             this.btnDuplicateRight.Location = new System.Drawing.Point(475, 0);
             this.btnDuplicateRight.Margin = new System.Windows.Forms.Padding(1);
             this.btnDuplicateRight.Name = "btnDuplicateRight";
             this.btnDuplicateRight.Size = new System.Drawing.Size(20, 20);
             this.btnDuplicateRight.TabIndex = 23;
             this.btnDuplicateRight.BackgroundImage = Image.FromFile("icons/position/icons8-duplicate-50.png");
             this.btnDuplicateRight.BackgroundImageLayout = ImageLayout.Stretch;
             this.btnDuplicateRight.UseVisualStyleBackColor = false;
             this.btnDuplicateRight.Click += new System.EventHandler(this.BtnDuplicateRight_Click);
             // btnCenterTopLeft
             this.btnCenterTopLeft.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
             this.btnCenterTopLeft.FlatAppearance.BorderSize = 0;
             this.btnCenterTopLeft.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
             this.btnCenterTopLeft.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
             this.btnCenterTopLeft.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
             this.btnCenterTopLeft.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
             this.btnCenterTopLeft.Location = new System.Drawing.Point(500, 0);
             this.btnCenterTopLeft.Margin = new System.Windows.Forms.Padding(1);
             this.btnCenterTopLeft.Name = "btnCenterTopLeft";
             this.btnCenterTopLeft.Size = new System.Drawing.Size(20, 20);
             this.btnCenterTopLeft.TabIndex = 24;
             this.btnCenterTopLeft.BackgroundImage = Image.FromFile("icons/position/icons8-snap-to-center-48.png");
             this.btnCenterTopLeft.BackgroundImageLayout = ImageLayout.Stretch;
             this.btnCenterTopLeft.UseVisualStyleBackColor = false;
             this.btnCenterTopLeft.Click += new System.EventHandler(this.BtnCenterTopLeft_Click);
             // btnSavePosition
             this.btnSavePosition.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
             this.btnSavePosition.FlatAppearance.BorderSize = 0;
             this.btnSavePosition.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
             this.btnSavePosition.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
             this.btnSavePosition.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
             this.btnSavePosition.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
             this.btnSavePosition.Location = new System.Drawing.Point(525, 0);
             this.btnSavePosition.Margin = new System.Windows.Forms.Padding(1);
             this.btnSavePosition.Name = "btnSavePosition";
             this.btnSavePosition.Size = new System.Drawing.Size(20, 20);
             this.btnSavePosition.TabIndex = 25;
             this.btnSavePosition.BackgroundImage = Image.FromFile("icons/position/icons8-save-50.png");
             this.btnSavePosition.BackgroundImageLayout = ImageLayout.Stretch;
             this.btnSavePosition.UseVisualStyleBackColor = false;
             this.btnSavePosition.Click += new System.EventHandler(this.BtnSavePosition_Click);
             // btnApplyPosition
             this.btnApplyPosition.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
             this.btnApplyPosition.FlatAppearance.BorderSize = 0;
             this.btnApplyPosition.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
             this.btnApplyPosition.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
             this.btnApplyPosition.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
             this.btnApplyPosition.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
             this.btnApplyPosition.Location = new System.Drawing.Point(550, 0);
             this.btnApplyPosition.Margin = new System.Windows.Forms.Padding(1);
             this.btnApplyPosition.Name = "btnApplyPosition";
             this.btnApplyPosition.Size = new System.Drawing.Size(20, 20);
             this.btnApplyPosition.TabIndex = 26;
             this.btnApplyPosition.BackgroundImage = Image.FromFile("icons/position/icons8-apply-64.png");
             this.btnApplyPosition.BackgroundImageLayout = ImageLayout.Stretch;
             this.btnApplyPosition.UseVisualStyleBackColor = false;
             this.btnApplyPosition.Click += new System.EventHandler(this.BtnApplyPosition_Click);
            // 
            // divider4
            // 
            this.divider4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.divider4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(225)))), ((int)(((byte)(225)))), ((int)(((byte)(225)))));
            this.divider4.Location = new System.Drawing.Point(3, 210);
            this.divider4.Name = "divider4";
            this.divider4.Size = new System.Drawing.Size(287, 1);
            this.divider4.TabIndex = 13;
            // 
            // sizePanel
            // 
            this.sizePanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.sizePanel.AutoSize = true;
            this.sizePanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.sizePanel.BackColor = System.Drawing.Color.White;
            this.sizePanel.Controls.Add(this.lblSizeSection);
            this.sizePanel.Controls.Add(this.sizeControlsPanel);
            this.sizePanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.sizePanel.Location = new System.Drawing.Point(3, 217);
            this.sizePanel.Name = "sizePanel";
            this.sizePanel.Size = new System.Drawing.Size(287, 141);
            this.sizePanel.TabIndex = 4;
            // 
            // lblSizeSection
            // 
            this.lblSizeSection.AutoSize = true;
            this.lblSizeSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblSizeSection.ForeColor = System.Drawing.Color.Gray;
            this.lblSizeSection.Location = new System.Drawing.Point(3, 0);
            this.lblSizeSection.Name = "lblSizeSection";
            this.lblSizeSection.Size = new System.Drawing.Size(27, 13);
            this.lblSizeSection.TabIndex = 0;
            this.lblSizeSection.Text = "Size";
            this.lblSizeSection.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // sizeControlsPanel
            // 
            this.sizeControlsPanel.AutoSize = true;
            this.sizeControlsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.sizeControlsPanel.Controls.Add(this.widthPanel);
            this.sizeControlsPanel.Controls.Add(this.heightPanel);
            this.sizeControlsPanel.Controls.Add(this.btnApplySize);
            this.sizeControlsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sizeControlsPanel.Location = new System.Drawing.Point(0, 13);
            this.sizeControlsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.sizeControlsPanel.Name = "sizeControlsPanel";
            this.sizeControlsPanel.Size = new System.Drawing.Size(287, 128);
            this.sizeControlsPanel.TabIndex = 1;
            // 
            // widthPanel
            // 
            this.widthPanel.AutoSize = true;
            this.widthPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.widthPanel.Controls.Add(this.flowLayoutPanel1);
            this.widthPanel.Location = new System.Drawing.Point(0, 0);
            this.widthPanel.Margin = new System.Windows.Forms.Padding(0);
            this.widthPanel.Name = "widthPanel";
            this.widthPanel.Size = new System.Drawing.Size(287, 33);
            this.widthPanel.TabIndex = 2;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.lblSlideSize);
            this.flowLayoutPanel1.Controls.Add(this.cmbSlideSize);
            this.flowLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(281, 27);
            this.flowLayoutPanel1.TabIndex = 2;
            // lblSlideSize
            this.lblSlideSize.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblSlideSize.ForeColor = System.Drawing.Color.Black;
            this.lblSlideSize.Location = new System.Drawing.Point(3, 0);
            this.lblSlideSize.Name = "lblSlideSize";
            this.lblSlideSize.Size = new System.Drawing.Size(85, 20);
            this.lblSlideSize.TabIndex = 0;
            this.lblSlideSize.Text = "Ratio: ";
            this.lblSlideSize.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // cmbSlideSize
            this.cmbSlideSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSlideSize.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.cmbSlideSize.FormattingEnabled = true;
            this.cmbSlideSize.Items.AddRange(new object[] {
            "Standard (4:3)",
            "Widescreen (16:9)",
            "Widescreen (16:10)",
            "Custom Size",
            "Standard (4:3)",
            "Widescreen (16:9)",
            "Custom"});
            this.cmbSlideSize.Location = new System.Drawing.Point(94, 3);
            this.cmbSlideSize.Name = "cmbSlideSize";
            this.cmbSlideSize.Size = new System.Drawing.Size(137, 21);
            this.cmbSlideSize.TabIndex = 1;
            this.cmbSlideSize.SelectedIndexChanged += new System.EventHandler(this.CmbSlideSize_SelectedIndexChanged);
            // heightPanel
            this.heightPanel.AutoSize = true;
            this.heightPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.heightPanel.Controls.Add(this.flowLayoutPanel2);
            this.heightPanel.Controls.Add(this.flowLayoutPanel3);
            this.heightPanel.Location = new System.Drawing.Point(0, 33);
            this.heightPanel.Margin = new System.Windows.Forms.Padding(0);
            this.heightPanel.Name = "heightPanel";
            this.heightPanel.Size = new System.Drawing.Size(287, 66);
            this.heightPanel.TabIndex = 3;
            // flowLayoutPanel2
            this.flowLayoutPanel2.Controls.Add(this.lblWidth);
            this.flowLayoutPanel2.Controls.Add(this.nudWidth);
            this.flowLayoutPanel2.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(281, 25);
            this.flowLayoutPanel2.TabIndex = 2;
            // lblWidth
            this.lblWidth.Location = new System.Drawing.Point(3, 0);
            this.lblWidth.Name = "lblWidth";
            this.lblWidth.Size = new System.Drawing.Size(85, 20);
            this.lblWidth.TabIndex = 2;
            this.lblWidth.Text = "Width: ";
            this.lblWidth.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblWidth.Click += new System.EventHandler(this.lblWidth_Click);
            // nudWidth
            this.nudWidth.Location = new System.Drawing.Point(94, 3);
            this.nudWidth.Name = "nudWidth";
            this.nudWidth.Size = new System.Drawing.Size(136, 20);
            this.nudWidth.TabIndex = 3;
            // flowLayoutPanel3
            this.flowLayoutPanel3.Controls.Add(this.lblHeight);
            this.flowLayoutPanel3.Controls.Add(this.nudHeight);
            this.flowLayoutPanel3.Location = new System.Drawing.Point(3, 34);
            this.flowLayoutPanel3.Name = "flowLayoutPanel3";
            this.flowLayoutPanel3.Size = new System.Drawing.Size(281, 29);
            this.flowLayoutPanel3.TabIndex = 6;
            // lblHeight
            this.lblHeight.Location = new System.Drawing.Point(3, 0);
            this.lblHeight.Name = "lblHeight";
            this.lblHeight.Size = new System.Drawing.Size(85, 20);
            this.lblHeight.TabIndex = 0;
            this.lblHeight.Text = "Height: ";
            this.lblHeight.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // nudHeight
            // 
            this.nudHeight.Location = new System.Drawing.Point(94, 3);
            this.nudHeight.Name = "nudHeight";
            this.nudHeight.Size = new System.Drawing.Size(136, 20);
            this.nudHeight.TabIndex = 1;
            // 
            // btnApplySize
            // 
            this.btnApplySize.AutoSize = true;
            this.btnApplySize.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(120)))), ((int)(((byte)(215)))));
            this.btnApplySize.FlatAppearance.BorderSize = 0;
            this.btnApplySize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnApplySize.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.btnApplySize.ForeColor = System.Drawing.Color.White;
            this.btnApplySize.Location = new System.Drawing.Point(3, 102);
            this.btnApplySize.Name = "btnApplySize";
            this.btnApplySize.Size = new System.Drawing.Size(281, 23);
            this.btnApplySize.TabIndex = 5;
            this.btnApplySize.Text = "Apply Size";
            this.btnApplySize.UseVisualStyleBackColor = false;
            this.btnApplySize.Click += new System.EventHandler(this.BtnApplySize_Click);
            // 
            // divider5
            // 
            this.divider5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.divider5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(225)))), ((int)(((byte)(225)))), ((int)(((byte)(225)))));
            this.divider5.Location = new System.Drawing.Point(3, 364);
            this.divider5.Name = "divider5";
            this.divider5.Size = new System.Drawing.Size(287, 1);
            this.divider5.TabIndex = 14;
            // 
            // shapePanel
            // 
            this.shapePanel.AutoSize = true;
            this.shapePanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.shapePanel.BackColor = System.Drawing.Color.White;
            this.shapePanel.Controls.Add(this.lblShapeSection);
            this.shapePanel.Controls.Add(this.shapeButtonsPanel);
            this.shapePanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.shapePanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.shapePanel.Location = new System.Drawing.Point(3, 371);
            this.shapePanel.Name = "shapePanel";
            this.shapePanel.Size = new System.Drawing.Size(287, 38);
            this.shapePanel.TabIndex = 5;
            // 
            // lblShapeSection
            // 
            this.lblShapeSection.AutoSize = true;
            this.lblShapeSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblShapeSection.ForeColor = System.Drawing.Color.Gray;
            this.lblShapeSection.Location = new System.Drawing.Point(3, 0);
            this.lblShapeSection.Name = "lblShapeSection";
            this.lblShapeSection.Size = new System.Drawing.Size(39, 13);
            this.lblShapeSection.TabIndex = 0;
            this.lblShapeSection.Text = "Shape";
            // 
            // shapeButtonsPanel
            // 
            this.shapeButtonsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.shapeButtonsPanel.AutoSize = true;
            this.shapeButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.shapeButtonsPanel.Controls.Add(this.btnAlignProcessChain);
            this.shapeButtonsPanel.Controls.Add(this.btnAlignAngles);
            this.shapeButtonsPanel.Controls.Add(this.btnAlignToProcessArrow);
            this.shapeButtonsPanel.Controls.Add(this.btnAdjustPentagonHeader);
            this.shapeButtonsPanel.Controls.Add(this.btnAlignBlockArrows);
            this.shapeButtonsPanel.Controls.Add(this.btnAlignRoundedRectangleArrows);
            this.shapeButtonsPanel.Location = new System.Drawing.Point(0, 13);
            this.shapeButtonsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.shapeButtonsPanel.Name = "shapeButtonsPanel";
            this.shapeButtonsPanel.Size = new System.Drawing.Size(150, 25);
            this.shapeButtonsPanel.TabIndex = 1;
            // 
            // btnAlignProcessChain
            // 
            this.btnAlignProcessChain.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignProcessChain.FlatAppearance.BorderSize = 0;
            this.btnAlignProcessChain.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignProcessChain.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignProcessChain.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignProcessChain.Font = new System.Drawing.Font("Segoe UI Emoji", 7F);
            this.btnAlignProcessChain.Location = new System.Drawing.Point(0, 0);
            this.btnAlignProcessChain.Margin = new System.Windows.Forms.Padding(1);
            this.btnAlignProcessChain.Name = "btnAlignProcessChain";
            this.btnAlignProcessChain.Size = new System.Drawing.Size(20, 20);
            this.btnAlignProcessChain.TabIndex = 1;
            this.btnAlignProcessChain.Text = "🔗";
            this.btnAlignProcessChain.UseVisualStyleBackColor = false;
            this.btnAlignProcessChain.Click += new System.EventHandler(this.BtnAlignProcessChain_Click);
            // 
            // btnAlignAngles
            // 
            this.btnAlignAngles.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignAngles.FlatAppearance.BorderSize = 0;
            this.btnAlignAngles.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignAngles.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignAngles.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignAngles.Font = new System.Drawing.Font("Segoe UI Emoji", 7F);
            this.btnAlignAngles.Location = new System.Drawing.Point(25, 0);
            this.btnAlignAngles.Margin = new System.Windows.Forms.Padding(1);
            this.btnAlignAngles.Name = "btnAlignAngles";
            this.btnAlignAngles.Size = new System.Drawing.Size(20, 20);
            this.btnAlignAngles.TabIndex = 2;
            this.btnAlignAngles.Text = "📐";
            this.btnAlignAngles.UseVisualStyleBackColor = false;
            this.btnAlignAngles.Click += new System.EventHandler(this.BtnAlignAngles_Click);
            // 
            // btnAlignToProcessArrow
            // 
            this.btnAlignToProcessArrow.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignToProcessArrow.FlatAppearance.BorderSize = 0;
            this.btnAlignToProcessArrow.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignToProcessArrow.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignToProcessArrow.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignToProcessArrow.Font = new System.Drawing.Font("Segoe UI Emoji", 7F);
            this.btnAlignToProcessArrow.Location = new System.Drawing.Point(50, 0);
            this.btnAlignToProcessArrow.Margin = new System.Windows.Forms.Padding(1);
            this.btnAlignToProcessArrow.Name = "btnAlignToProcessArrow";
            this.btnAlignToProcessArrow.Size = new System.Drawing.Size(20, 20);
            this.btnAlignToProcessArrow.TabIndex = 3;
            this.btnAlignToProcessArrow.Text = "➡️";
            this.btnAlignToProcessArrow.UseVisualStyleBackColor = false;
            this.btnAlignToProcessArrow.Click += new System.EventHandler(this.BtnAlignToProcessArrow_Click);
            // 
            // btnAdjustPentagonHeader
            // 
            this.btnAdjustPentagonHeader.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAdjustPentagonHeader.FlatAppearance.BorderSize = 0;
            this.btnAdjustPentagonHeader.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAdjustPentagonHeader.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAdjustPentagonHeader.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAdjustPentagonHeader.Font = new System.Drawing.Font("Segoe UI Emoji", 7F);
            this.btnAdjustPentagonHeader.Location = new System.Drawing.Point(75, 0);
            this.btnAdjustPentagonHeader.Margin = new System.Windows.Forms.Padding(1);
            this.btnAdjustPentagonHeader.Name = "btnAdjustPentagonHeader";
            this.btnAdjustPentagonHeader.Size = new System.Drawing.Size(20, 20);
            this.btnAdjustPentagonHeader.TabIndex = 4;
            this.btnAdjustPentagonHeader.Text = "🔷";
            this.btnAdjustPentagonHeader.UseVisualStyleBackColor = false;
            this.btnAdjustPentagonHeader.Click += new System.EventHandler(this.BtnAdjustPentagonHeader_Click);
            // 
            // btnAlignBlockArrows
            // 
            this.btnAlignBlockArrows.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignBlockArrows.FlatAppearance.BorderSize = 0;
            this.btnAlignBlockArrows.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignBlockArrows.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignBlockArrows.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignBlockArrows.Font = new System.Drawing.Font("Segoe UI Emoji", 7F);
            this.btnAlignBlockArrows.Location = new System.Drawing.Point(100, 0);
            this.btnAlignBlockArrows.Margin = new System.Windows.Forms.Padding(1);
            this.btnAlignBlockArrows.Name = "btnAlignBlockArrows";
            this.btnAlignBlockArrows.Size = new System.Drawing.Size(20, 20);
            this.btnAlignBlockArrows.TabIndex = 5;
            this.btnAlignBlockArrows.Text = "▶️";
            this.btnAlignBlockArrows.UseVisualStyleBackColor = false;
            this.btnAlignBlockArrows.Click += new System.EventHandler(this.BtnAlignBlockArrows_Click);
            // 
            // btnAlignRoundedRectangleArrows
            // 
            this.btnAlignRoundedRectangleArrows.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignRoundedRectangleArrows.FlatAppearance.BorderSize = 0;
            this.btnAlignRoundedRectangleArrows.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignRoundedRectangleArrows.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignRoundedRectangleArrows.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignRoundedRectangleArrows.Font = new System.Drawing.Font("Segoe UI Emoji", 7F);
            this.btnAlignRoundedRectangleArrows.Location = new System.Drawing.Point(125, 0);
            this.btnAlignRoundedRectangleArrows.Margin = new System.Windows.Forms.Padding(1);
            this.btnAlignRoundedRectangleArrows.Name = "btnAlignRoundedRectangleArrows";
            this.btnAlignRoundedRectangleArrows.Size = new System.Drawing.Size(20, 20);
            this.btnAlignRoundedRectangleArrows.TabIndex = 6;
            this.btnAlignRoundedRectangleArrows.Text = "🔲";
            this.btnAlignRoundedRectangleArrows.UseVisualStyleBackColor = false;
            this.btnAlignRoundedRectangleArrows.Click += new System.EventHandler(this.BtnAlignRoundedRectangleArrows_Click);
            // 
            // divider6
            // 
            this.divider6.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.divider6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(225)))), ((int)(((byte)(225)))), ((int)(((byte)(225)))));
            this.divider6.Location = new System.Drawing.Point(3, 415);
            this.divider6.Name = "divider6";
            this.divider6.Size = new System.Drawing.Size(287, 1);
            this.divider6.TabIndex = 15;
            // 
            // colorPanel
            // 
            this.colorPanel.AutoSize = true;
            this.colorPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.colorPanel.BackColor = System.Drawing.Color.White;
            this.colorPanel.Controls.Add(this.lblColorSection);
            this.colorPanel.Controls.Add(this.colorButtonsPanel);
            this.colorPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.colorPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.colorPanel.Location = new System.Drawing.Point(3, 422);
            this.colorPanel.Name = "colorPanel";
            this.colorPanel.Size = new System.Drawing.Size(287, 38);
            this.colorPanel.TabIndex = 6;
            // 
            // lblColorSection
            // 
            this.lblColorSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblColorSection.ForeColor = System.Drawing.Color.Gray;
            this.lblColorSection.Location = new System.Drawing.Point(3, 0);
            this.lblColorSection.Name = "lblColorSection";
            this.lblColorSection.Size = new System.Drawing.Size(100, 13);
            this.lblColorSection.TabIndex = 0;
            this.lblColorSection.Text = "Color";
            // colorButtonsPanel
            this.colorButtonsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.colorButtonsPanel.AutoSize = true;
            this.colorButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.colorButtonsPanel.Controls.Add(this.btnFillColor);
            this.colorButtonsPanel.Controls.Add(this.btnTextColor);
            this.colorButtonsPanel.Controls.Add(this.btnOutlineColor);
            this.colorButtonsPanel.Location = new System.Drawing.Point(0, 13);
            this.colorButtonsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.colorButtonsPanel.Name = "colorButtonsPanel";
            this.colorButtonsPanel.Size = new System.Drawing.Size(106, 25);
            this.colorButtonsPanel.TabIndex = 1;
            // btnFillColor
            this.btnFillColor.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnFillColor.FlatAppearance.BorderSize = 0;
            this.btnFillColor.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnFillColor.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnFillColor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFillColor.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnFillColor.Location = new System.Drawing.Point(0, 0);
            this.btnFillColor.Margin = new System.Windows.Forms.Padding(0);
            this.btnFillColor.Name = "btnFillColor";
            this.btnFillColor.Size = new System.Drawing.Size(20, 20);
            this.btnFillColor.TabIndex = 1;
            this.btnFillColor.Text = "🎨";
            this.btnFillColor.UseVisualStyleBackColor = false;
            this.btnFillColor.Click += new System.EventHandler(this.BtnFillColor_Click);
            // btnTextColor
            this.btnTextColor.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnTextColor.FlatAppearance.BorderSize = 0;
            this.btnTextColor.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnTextColor.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnTextColor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnTextColor.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnTextColor.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnTextColor.Margin = new System.Windows.Forms.Padding(0);
            this.btnTextColor.Name = "btnTextColor";
            this.btnTextColor.Size = new System.Drawing.Size(20, 20);
            this.btnTextColor.TabIndex = 2;
            this.btnTextColor.Text = "A";
            this.btnTextColor.UseVisualStyleBackColor = false;
            this.btnTextColor.Click += new System.EventHandler(this.BtnTextColor_Click);
            // btnOutlineColor
            this.btnOutlineColor.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnOutlineColor.FlatAppearance.BorderSize = 0;
            this.btnOutlineColor.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnOutlineColor.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnOutlineColor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOutlineColor.Font = new System.Drawing.Font("Segoe UI Emoji", 7F);
            this.btnOutlineColor.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnOutlineColor.Location = new System.Drawing.Point(50, 0);
            this.btnOutlineColor.Margin = new System.Windows.Forms.Padding(0);
            this.btnOutlineColor.Name = "btnOutlineColor";
            this.btnOutlineColor.Size = new System.Drawing.Size(20, 20);
            this.btnOutlineColor.TabIndex = 3;
            this.btnOutlineColor.Text = "◯";
            this.btnOutlineColor.UseVisualStyleBackColor = false;
            this.btnOutlineColor.Click += new System.EventHandler(this.BtnOutlineColor_Click);
            // divider7
            this.divider7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.divider7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(225)))), ((int)(((byte)(225)))), ((int)(((byte)(225)))));
            this.divider7.Location = new System.Drawing.Point(3, 466);
            this.divider7.Name = "divider7";
            this.divider7.Size = new System.Drawing.Size(287, 1);
            this.divider7.TabIndex = 16;
            // textPanel
            this.textPanel.AutoSize = true;
            this.textPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.textPanel.BackColor = System.Drawing.Color.White;
            this.textPanel.Controls.Add(this.lblTextSection);
            this.textPanel.Controls.Add(this.textButtonsPanel);
            this.textPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.textPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.textPanel.Location = new System.Drawing.Point(3, 473);
            this.textPanel.Name = "textPanel";
            this.textPanel.Size = new System.Drawing.Size(287, 38);
            this.textPanel.TabIndex = 7;
            // lblTextSection
            this.lblTextSection.AutoSize = true;
            this.lblTextSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblTextSection.ForeColor = System.Drawing.Color.Gray;
            this.lblTextSection.Location = new System.Drawing.Point(3, 0);
            this.lblTextSection.Name = "lblTextSection";
            this.lblTextSection.Size = new System.Drawing.Size(26, 13);
            this.lblTextSection.TabIndex = 0;
            this.lblTextSection.Text = "Text";
            // textButtonsPanel
            this.textButtonsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textButtonsPanel.AutoSize = true;
            this.textButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.textButtonsPanel.Controls.Add(this.btnBold);
            this.textButtonsPanel.Controls.Add(this.btnItalic);
            this.textButtonsPanel.Controls.Add(this.btnUnderline);
            this.textButtonsPanel.Controls.Add(this.btnBullets);
            this.textButtonsPanel.Controls.Add(this.btnWrapText);
            this.textButtonsPanel.Controls.Add(this.btnNoWrapText);
            this.textButtonsPanel.Location = new System.Drawing.Point(0, 13);
            this.textButtonsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.textButtonsPanel.Name = "textButtonsPanel";
            this.textButtonsPanel.Size = new System.Drawing.Size(150, 25);
            this.textButtonsPanel.TabIndex = 1;
            // btnBold
            this.btnBold.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnBold.FlatAppearance.BorderSize = 0;
            this.btnBold.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnBold.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnBold.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBold.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.btnBold.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnBold.Location = new System.Drawing.Point(0, 0);
            this.btnBold.Margin = new System.Windows.Forms.Padding(0);
            this.btnBold.Name = "btnBold";
            this.btnBold.Size = new System.Drawing.Size(20, 20);
            this.btnBold.TabIndex = 1;
            this.btnBold.Text = "B";
            this.btnBold.UseVisualStyleBackColor = false;
            this.btnBold.Click += new System.EventHandler(this.BtnBold_Click);
            // 
            // btnItalic
            // 
            this.btnItalic.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnItalic.FlatAppearance.BorderSize = 0;
            this.btnItalic.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnItalic.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnItalic.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnItalic.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            this.btnItalic.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnItalic.Location = new System.Drawing.Point(25, 0);
            this.btnItalic.Margin = new System.Windows.Forms.Padding(0);
            this.btnItalic.Name = "btnItalic";
            this.btnItalic.Size = new System.Drawing.Size(20, 20);
            this.btnItalic.TabIndex = 2;
            this.btnItalic.Text = "I";
            this.btnItalic.UseVisualStyleBackColor = false;
            this.btnItalic.Click += new System.EventHandler(this.BtnItalic_Click);
            // 
            // btnUnderline
            // 
            this.btnUnderline.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnUnderline.FlatAppearance.BorderSize = 0;
            this.btnUnderline.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnUnderline.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnUnderline.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnUnderline.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Underline);
            this.btnUnderline.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnUnderline.Location = new System.Drawing.Point(50, 0);
            this.btnUnderline.Margin = new System.Windows.Forms.Padding(0);
            this.btnUnderline.Name = "btnUnderline";
            this.btnUnderline.Size = new System.Drawing.Size(20, 20);
            this.btnUnderline.TabIndex = 3;
            this.btnUnderline.Text = "U";
            this.btnUnderline.UseVisualStyleBackColor = false;
            this.btnUnderline.Click += new System.EventHandler(this.BtnUnderline_Click);
            // 
            // btnBullets
            // 
            this.btnBullets.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnBullets.FlatAppearance.BorderSize = 0;
            this.btnBullets.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnBullets.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnBullets.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBullets.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.btnBullets.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnBullets.Location = new System.Drawing.Point(75, 0);
            this.btnBullets.Margin = new System.Windows.Forms.Padding(0);
            this.btnBullets.Name = "btnBullets";
            this.btnBullets.Size = new System.Drawing.Size(20, 20);
            this.btnBullets.TabIndex = 4;
            this.btnBullets.Text = "•";
            this.btnBullets.UseVisualStyleBackColor = false;
            this.btnBullets.Click += new System.EventHandler(this.BtnBullets_Click);
            // 
            // btnWrapText
            // 
            this.btnWrapText.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnWrapText.FlatAppearance.BorderSize = 0;
            this.btnWrapText.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnWrapText.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnWrapText.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnWrapText.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnWrapText.Location = new System.Drawing.Point(100, 0);
            this.btnWrapText.Margin = new System.Windows.Forms.Padding(0);
            this.btnWrapText.Name = "btnWrapText";
            this.btnWrapText.Size = new System.Drawing.Size(20, 20);
            this.btnWrapText.TabIndex = 5;
            this.btnWrapText.Text = "📦";
            this.btnWrapText.UseVisualStyleBackColor = false;
            this.btnWrapText.Click += new System.EventHandler(this.BtnWrapText_Click);
            // 
            // btnNoWrapText
            // 
            this.btnNoWrapText.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnNoWrapText.FlatAppearance.BorderSize = 0;
            this.btnNoWrapText.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnNoWrapText.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnNoWrapText.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNoWrapText.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnNoWrapText.Location = new System.Drawing.Point(125, 0);
            this.btnNoWrapText.Margin = new System.Windows.Forms.Padding(0);
            this.btnNoWrapText.Name = "btnNoWrapText";
            this.btnNoWrapText.Size = new System.Drawing.Size(20, 20);
            this.btnNoWrapText.TabIndex = 6;
            this.btnNoWrapText.Text = "📄";
            this.btnNoWrapText.UseVisualStyleBackColor = false;
            this.btnNoWrapText.Click += new System.EventHandler(this.BtnNoWrapText_Click);
            // 
            // divider8
            // 
            this.divider8.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.divider8.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(225)))), ((int)(((byte)(225)))), ((int)(((byte)(225)))));
            this.divider8.Location = new System.Drawing.Point(3, 517);
            this.divider8.Name = "divider8";
            this.divider8.Size = new System.Drawing.Size(287, 1);
            this.divider8.TabIndex = 17;
            // 
            // navigationPanel
            // 
            this.navigationPanel.AutoSize = true;
            this.navigationPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.navigationPanel.BackColor = System.Drawing.Color.White;
            this.navigationPanel.Controls.Add(this.lblNavigationSection);
            this.navigationPanel.Controls.Add(this.navigationButtonsPanel);
            this.navigationPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.navigationPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.navigationPanel.Location = new System.Drawing.Point(3, 524);
            this.navigationPanel.Name = "navigationPanel";
            this.navigationPanel.Size = new System.Drawing.Size(287, 38);
            this.navigationPanel.TabIndex = 8;
            // 
            // lblNavigationSection
            // 
            this.lblNavigationSection.AutoSize = true;
            this.lblNavigationSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblNavigationSection.ForeColor = System.Drawing.Color.Gray;
            this.lblNavigationSection.Location = new System.Drawing.Point(3, 0);
            this.lblNavigationSection.Name = "lblNavigationSection";
            this.lblNavigationSection.Size = new System.Drawing.Size(91, 13);
            this.lblNavigationSection.TabIndex = 0;
            this.lblNavigationSection.Text = "Navigation View";
            // 
            // navigationButtonsPanel
            // 
            this.navigationButtonsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.navigationButtonsPanel.AutoSize = true;
            this.navigationButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.navigationButtonsPanel.Controls.Add(this.btnZoomIn);
            this.navigationButtonsPanel.Controls.Add(this.btnZoomOut);
            this.navigationButtonsPanel.Controls.Add(this.btnFitToWindow);
            this.navigationButtonsPanel.Location = new System.Drawing.Point(0, 13);
            this.navigationButtonsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.navigationButtonsPanel.Name = "navigationButtonsPanel";
            this.navigationButtonsPanel.Size = new System.Drawing.Size(97, 25);
            this.navigationButtonsPanel.TabIndex = 1;
            // 
            // btnZoomIn
            // 
            this.btnZoomIn.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnZoomIn.FlatAppearance.BorderSize = 0;
            this.btnZoomIn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnZoomIn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnZoomIn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnZoomIn.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnZoomIn.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnZoomIn.Location = new System.Drawing.Point(0, 0);
            this.btnZoomIn.Margin = new System.Windows.Forms.Padding(0);
            this.btnZoomIn.Name = "btnZoomIn";
            this.btnZoomIn.Size = new System.Drawing.Size(20, 20);
            this.btnZoomIn.TabIndex = 1;
            this.btnZoomIn.Text = "➕";
            this.btnZoomIn.UseVisualStyleBackColor = false;
            this.btnZoomIn.Click += new System.EventHandler(this.BtnZoomIn_Click);
            // 
            // btnZoomOut
            // 
            this.btnZoomOut.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnZoomOut.FlatAppearance.BorderSize = 0;
            this.btnZoomOut.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnZoomOut.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnZoomOut.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnZoomOut.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnZoomOut.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnZoomOut.Location = new System.Drawing.Point(25, 0);
            this.btnZoomOut.Margin = new System.Windows.Forms.Padding(0);
            this.btnZoomOut.Name = "btnZoomOut";
            this.btnZoomOut.Size = new System.Drawing.Size(20, 20);
            this.btnZoomOut.TabIndex = 2;
            this.btnZoomOut.Text = "➖";
            this.btnZoomOut.UseVisualStyleBackColor = false;
            this.btnZoomOut.Click += new System.EventHandler(this.BtnZoomOut_Click);
            // 
            // btnFitToWindow
            // 
            this.btnFitToWindow.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnFitToWindow.FlatAppearance.BorderSize = 0;
            this.btnFitToWindow.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnFitToWindow.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnFitToWindow.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFitToWindow.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnFitToWindow.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnFitToWindow.Location = new System.Drawing.Point(50, 0);
            this.btnFitToWindow.Margin = new System.Windows.Forms.Padding(0);
            this.btnFitToWindow.Name = "btnFitToWindow";
            this.btnFitToWindow.Size = new System.Drawing.Size(20, 20);
            this.btnFitToWindow.TabIndex = 3;
            this.btnFitToWindow.Text = "⛶";
            this.btnFitToWindow.UseVisualStyleBackColor = false;
            this.btnFitToWindow.Click += new System.EventHandler(this.BtnFitToWindow_Click);
            // 
            // divider9
            // 
            this.divider9.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.divider9.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(225)))), ((int)(((byte)(225)))), ((int)(((byte)(225)))));
            this.divider9.Location = new System.Drawing.Point(3, 568);
            this.divider9.Name = "divider9";
            this.divider9.Size = new System.Drawing.Size(287, 1);
            this.divider9.TabIndex = 18;
            // 
            // expertToolsPanel
            // 
            this.expertToolsPanel.AutoSize = true;
            this.expertToolsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.expertToolsPanel.BackColor = System.Drawing.Color.White;
            this.expertToolsPanel.Controls.Add(this.lblExpertToolsSection);
            this.expertToolsPanel.Controls.Add(this.expertToolsButtonsPanel);
            this.expertToolsPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.expertToolsPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.expertToolsPanel.Location = new System.Drawing.Point(3, 575);
            this.expertToolsPanel.Name = "expertToolsPanel";
            this.expertToolsPanel.Size = new System.Drawing.Size(287, 45);
            this.expertToolsPanel.TabIndex = 9;
            // 
            // lblExpertToolsSection
            // 
            this.lblExpertToolsSection.AutoSize = true;
            this.lblExpertToolsSection.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblExpertToolsSection.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.lblExpertToolsSection.Location = new System.Drawing.Point(3, 0);
            this.lblExpertToolsSection.Name = "lblExpertToolsSection";
            this.lblExpertToolsSection.Size = new System.Drawing.Size(75, 15);
            this.lblExpertToolsSection.TabIndex = 0;
            this.lblExpertToolsSection.Text = "Expert Tools";
            // 
            // expertToolsButtonsPanel
            // 
            this.expertToolsButtonsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.expertToolsButtonsPanel.AutoSize = true;
            this.expertToolsButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.expertToolsButtonsPanel.Controls.Add(this.btnFreeWebinar);
            this.expertToolsButtonsPanel.Location = new System.Drawing.Point(0, 15);
            this.expertToolsButtonsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.expertToolsButtonsPanel.Name = "expertToolsButtonsPanel";
            this.expertToolsButtonsPanel.Size = new System.Drawing.Size(284, 30);
            this.expertToolsButtonsPanel.TabIndex = 1;
            // 
            // btnFreeWebinar
            // 
            this.btnFreeWebinar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(120)))), ((int)(((byte)(215)))));
            this.btnFreeWebinar.FlatAppearance.BorderSize = 0;
            this.btnFreeWebinar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFreeWebinar.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnFreeWebinar.ForeColor = System.Drawing.Color.White;
            this.btnFreeWebinar.Location = new System.Drawing.Point(0, 0);
            this.btnFreeWebinar.Margin = new System.Windows.Forms.Padding(0);
            this.btnFreeWebinar.Name = "btnFreeWebinar";
            this.btnFreeWebinar.Size = new System.Drawing.Size(284, 30);
            this.btnFreeWebinar.TabIndex = 1;
            this.btnFreeWebinar.Text = "🎓 Free PowerPoint Webinar";
            this.btnFreeWebinar.UseVisualStyleBackColor = false;
            this.btnFreeWebinar.Click += new System.EventHandler(this.BtnFreeWebinar_Click);
            // 
            // TaskPaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.mainScrollPanel);
            this.Name = "TaskPaneControl";
            this.Size = new System.Drawing.Size(300, 600);
            this.mainScrollPanel.ResumeLayout(false);
            this.mainScrollPanel.PerformLayout();
            this.sectionsContainer.ResumeLayout(false);
            this.sectionsContainer.PerformLayout();
            this.presentationPanel.ResumeLayout(false);
            this.presentationPanel.PerformLayout();
            this.presentationButtonsPanel.ResumeLayout(false);
            this.wizardsPanel.ResumeLayout(false);
            this.wizardsPanel.PerformLayout();
            this.wizardButtonsPanel.ResumeLayout(false);
            this.smartElementsPanel.ResumeLayout(false);
            this.smartElementsPanel.PerformLayout();
            this.smartElementsButtonsPanel.ResumeLayout(false);
            this.positionPanel.ResumeLayout(false);
            this.positionPanel.PerformLayout();
            this.positionButtonsPanel.ResumeLayout(false);
            this.sizePanel.ResumeLayout(false);
            this.sizePanel.PerformLayout();
            this.sizeControlsPanel.ResumeLayout(false);
            this.sizeControlsPanel.PerformLayout();
            this.widthPanel.ResumeLayout(false);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.heightPanel.ResumeLayout(false);
            this.flowLayoutPanel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nudWidth)).EndInit();
            this.flowLayoutPanel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nudHeight)).EndInit();
            this.shapePanel.ResumeLayout(false);
            this.shapePanel.PerformLayout();
            this.shapeButtonsPanel.ResumeLayout(false);
            this.colorPanel.ResumeLayout(false);
            this.colorPanel.PerformLayout();
            this.colorButtonsPanel.ResumeLayout(false);
            this.textPanel.ResumeLayout(false);
            this.textPanel.PerformLayout();
            this.textButtonsPanel.ResumeLayout(false);
            this.navigationPanel.ResumeLayout(false);
            this.navigationPanel.PerformLayout();
            this.navigationButtonsPanel.ResumeLayout(false);
            this.expertToolsPanel.ResumeLayout(false);
            this.expertToolsPanel.PerformLayout();
            this.expertToolsButtonsPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        private void BtnText_Click(object sender, EventArgs e)
        {
            // This will be implemented by the actual event handler
        }

        #endregion

        private System.Windows.Forms.Panel mainScrollPanel;
        private System.Windows.Forms.FlowLayoutPanel sectionsContainer;
        
        // Presentation section
        private System.Windows.Forms.FlowLayoutPanel presentationPanel;
        private System.Windows.Forms.FlowLayoutPanel presentationButtonsPanel;
        private System.Windows.Forms.Label lblPresentationSection;
        private System.Windows.Forms.Button btnNew;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnSaveAs;
        private System.Windows.Forms.Button btnPrint;
        //private System.Windows.Forms.Button btnShare;
        
        // Wizards section
        private System.Windows.Forms.FlowLayoutPanel wizardsPanel;
        private System.Windows.Forms.FlowLayoutPanel wizardButtonsPanel;
        private System.Windows.Forms.Label lblWizardsSection;
        private System.Windows.Forms.Button btnAgenda;
        private System.Windows.Forms.Button btnMaster;
        private System.Windows.Forms.Button btnElement;
        private System.Windows.Forms.Button btnText;
        private System.Windows.Forms.Button btnFormat;
        private System.Windows.Forms.Button btnMap;
        
        // Smart Elements section
        private System.Windows.Forms.FlowLayoutPanel smartElementsPanel;
        private System.Windows.Forms.FlowLayoutPanel smartElementsButtonsPanel;
        private System.Windows.Forms.Label lblSmartElementsSection;
        private System.Windows.Forms.Button btnChart;
        private System.Windows.Forms.Button btnDiagram;
        private System.Windows.Forms.Button btnTable;
        private System.Windows.Forms.Button btnMatrixTable;
        private System.Windows.Forms.Button btnStickyNote;
        private System.Windows.Forms.Button btnCitation;
        private System.Windows.Forms.Button btnStandardObjects;
        
        // Position section
        private System.Windows.Forms.FlowLayoutPanel positionPanel;
        private System.Windows.Forms.FlowLayoutPanel positionButtonsPanel;
        private System.Windows.Forms.Label lblPositionSection;
        private System.Windows.Forms.Button btnAlignLeft;
        private System.Windows.Forms.Button btnAlignCenter;
        private System.Windows.Forms.Button btnAlignRight;
        private System.Windows.Forms.Button btnAlignTop;
        private System.Windows.Forms.Button btnAlignBottom;
        private System.Windows.Forms.Button btnAlignMiddle;
                 private System.Windows.Forms.Button btnDockLeft;
         private System.Windows.Forms.Button btnDockRight;
         private System.Windows.Forms.Button btnDockTop;
         private System.Windows.Forms.Button btnDockBottom;
        private System.Windows.Forms.Button btnDistribute;
        private System.Windows.Forms.Button btnDistributeHorizontal;
        private System.Windows.Forms.Button btnDistributeVertical;
        private System.Windows.Forms.Button btnMatchBoth;
        private System.Windows.Forms.Button btnMatchHeight;
        private System.Windows.Forms.Button btnMatchWidth;
        private System.Windows.Forms.Button btnMakeVertical;
        private System.Windows.Forms.Button btnMakeHorizontal;
        private System.Windows.Forms.Button btnSwapLocations;
        private System.Windows.Forms.Button btnGoldenCanon;
        private System.Windows.Forms.Button btnAlignMatrix;
        private System.Windows.Forms.Button btnSliceShape;
        private System.Windows.Forms.Button btnDuplicateRight;
        private System.Windows.Forms.Button btnCenterTopLeft;
        private System.Windows.Forms.Button btnSavePosition;
        private System.Windows.Forms.Button btnApplyPosition;
        
        // Size section
        private System.Windows.Forms.FlowLayoutPanel sizePanel;
        private System.Windows.Forms.Label lblSizeSection;
        private System.Windows.Forms.Label lblSlideSize;
        private System.Windows.Forms.ComboBox cmbSlideSize;
        private System.Windows.Forms.NumericUpDown nudWidth;
        private System.Windows.Forms.Label lblHeight;
        private System.Windows.Forms.NumericUpDown nudHeight;
        private System.Windows.Forms.Button btnApplySize;
        
        // Shape section
        private System.Windows.Forms.FlowLayoutPanel shapePanel;
        private System.Windows.Forms.FlowLayoutPanel shapeButtonsPanel;
        private System.Windows.Forms.Label lblShapeSection;
        private System.Windows.Forms.Button btnAlignProcessChain;
        private System.Windows.Forms.Button btnAlignAngles;
        private System.Windows.Forms.Button btnAlignToProcessArrow;
        private System.Windows.Forms.Button btnAdjustPentagonHeader;
        private System.Windows.Forms.Button btnAlignBlockArrows;
        private System.Windows.Forms.Button btnAlignRoundedRectangleArrows;
        
        // Color section
        private System.Windows.Forms.FlowLayoutPanel colorPanel;
        private System.Windows.Forms.FlowLayoutPanel colorButtonsPanel;
        private System.Windows.Forms.Label lblColorSection;
        private System.Windows.Forms.Button btnFillColor;
        private System.Windows.Forms.Button btnTextColor;
        private System.Windows.Forms.Button btnOutlineColor;
        
        // Text section
        private System.Windows.Forms.FlowLayoutPanel textPanel;
        private System.Windows.Forms.FlowLayoutPanel textButtonsPanel;
        private System.Windows.Forms.Label lblTextSection;
        private System.Windows.Forms.Button btnBold;
        private System.Windows.Forms.Button btnItalic;
        private System.Windows.Forms.Button btnUnderline;
        private System.Windows.Forms.Button btnBullets;
        private System.Windows.Forms.Button btnWrapText;
        private System.Windows.Forms.Button btnNoWrapText;
        
        // Navigation & View section
        private System.Windows.Forms.FlowLayoutPanel navigationPanel;
        private System.Windows.Forms.FlowLayoutPanel navigationButtonsPanel;
        private System.Windows.Forms.Label lblNavigationSection;
        private System.Windows.Forms.Button btnZoomIn;
        private System.Windows.Forms.Button btnZoomOut;
        private System.Windows.Forms.Button btnFitToWindow;
        
        // Expert Tools section
        private System.Windows.Forms.FlowLayoutPanel expertToolsPanel;
        private System.Windows.Forms.FlowLayoutPanel expertToolsButtonsPanel;
        private System.Windows.Forms.Label lblExpertToolsSection;
        private System.Windows.Forms.Button btnFreeWebinar;
        private System.Windows.Forms.FlowLayoutPanel sizeControlsPanel;
        private System.Windows.Forms.Label lblWidth;
        private System.Windows.Forms.FlowLayoutPanel widthPanel;
        private System.Windows.Forms.FlowLayoutPanel heightPanel;
        private System.Windows.Forms.Panel divider1;
        private System.Windows.Forms.Panel divider2;
        private System.Windows.Forms.Panel divider3;
        private System.Windows.Forms.Panel divider4;
        private System.Windows.Forms.Panel divider5;
        private System.Windows.Forms.Panel divider6;
        private System.Windows.Forms.Panel divider7;
        private System.Windows.Forms.Panel divider8;
        private System.Windows.Forms.Panel divider9;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel3;
    }
}