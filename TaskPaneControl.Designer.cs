using System;

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
            this.btnShare = new System.Windows.Forms.Button();
            this.wizardsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblWizardsSection = new System.Windows.Forms.Label();
            this.wizardButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnAgenda = new System.Windows.Forms.Button();
            this.btnMaster = new System.Windows.Forms.Button();
            this.btnElement = new System.Windows.Forms.Button();
            this.btnText = new System.Windows.Forms.Button();
            this.btnFormat = new System.Windows.Forms.Button();
            this.btnMap = new System.Windows.Forms.Button();
            this.smartElementsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblSmartElementsSection = new System.Windows.Forms.Label();
            this.smartElementsButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnChart = new System.Windows.Forms.Button();
            this.btnDiagram = new System.Windows.Forms.Button();
            this.btnTable = new System.Windows.Forms.Button();
            this.positionPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblPositionSection = new System.Windows.Forms.Label();
            this.positionButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnAlignLeft = new System.Windows.Forms.Button();
            this.btnAlignCenter = new System.Windows.Forms.Button();
            this.btnAlignRight = new System.Windows.Forms.Button();
            this.btnDistribute = new System.Windows.Forms.Button();
            this.sizePanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblSizeSection = new System.Windows.Forms.Label();
            this.lblSlideSize = new System.Windows.Forms.Label();
            this.cmbSlideSize = new System.Windows.Forms.ComboBox();
            this.nudWidth = new System.Windows.Forms.NumericUpDown();
            this.lblHeight = new System.Windows.Forms.Label();
            this.nudHeight = new System.Windows.Forms.NumericUpDown();
            this.btnApplySize = new System.Windows.Forms.Button();
            this.shapePanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblShapeSection = new System.Windows.Forms.Label();
            this.shapeButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnRectangle = new System.Windows.Forms.Button();
            this.btnCircle = new System.Windows.Forms.Button();
            this.btnArrow = new System.Windows.Forms.Button();
            this.btnLine = new System.Windows.Forms.Button();
            this.colorPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblColorSection = new System.Windows.Forms.Label();
            this.colorButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnFillColor = new System.Windows.Forms.Button();
            this.btnTextColor = new System.Windows.Forms.Button();
            this.btnOutlineColor = new System.Windows.Forms.Button();
            this.textPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblTextSection = new System.Windows.Forms.Label();
            this.textButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnBold = new System.Windows.Forms.Button();
            this.btnItalic = new System.Windows.Forms.Button();
            this.btnUnderline = new System.Windows.Forms.Button();
            this.btnBullets = new System.Windows.Forms.Button();
            this.navigationPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblNavigationSection = new System.Windows.Forms.Label();
            this.navigationButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnZoomIn = new System.Windows.Forms.Button();
            this.btnZoomOut = new System.Windows.Forms.Button();
            this.btnFitToWindow = new System.Windows.Forms.Button();
            this.expertToolsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblExpertToolsSection = new System.Windows.Forms.Label();
            this.expertToolsButtonsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnFreeWebinar = new System.Windows.Forms.Button();
            this.lblWidth = new System.Windows.Forms.Label();
            this.sizeControlsPanel = new System.Windows.Forms.FlowLayoutPanel();
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
            ((System.ComponentModel.ISupportInitialize)(this.nudWidth)).BeginInit();
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
            this.sizeControlsPanel.SuspendLayout();
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
            this.mainScrollPanel.Padding = new System.Windows.Forms.Padding(5);
            this.mainScrollPanel.Size = new System.Drawing.Size(300, 600);
            this.mainScrollPanel.TabIndex = 0;
            // 
            // sectionsContainer
            // 
            this.sectionsContainer.AutoSize = true;
            this.sectionsContainer.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.sectionsContainer.Controls.Add(this.presentationPanel);
            this.sectionsContainer.Controls.Add(this.wizardsPanel);
            this.sectionsContainer.Controls.Add(this.smartElementsPanel);
            this.sectionsContainer.Controls.Add(this.positionPanel);
            this.sectionsContainer.Controls.Add(this.sizePanel);
            this.sectionsContainer.Controls.Add(this.shapePanel);
            this.sectionsContainer.Controls.Add(this.colorPanel);
            this.sectionsContainer.Controls.Add(this.textPanel);
            this.sectionsContainer.Controls.Add(this.navigationPanel);
            this.sectionsContainer.Controls.Add(this.expertToolsPanel);
            this.sectionsContainer.Dock = System.Windows.Forms.DockStyle.Top;
            this.sectionsContainer.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.sectionsContainer.Location = new System.Drawing.Point(5, 5);
            this.sectionsContainer.Name = "sectionsContainer";
            this.sectionsContainer.Size = new System.Drawing.Size(273, 984);
            this.sectionsContainer.TabIndex = 0;
            this.sectionsContainer.WrapContents = false;
            // 
            // presentationPanel
            // 
            this.presentationPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.presentationPanel.AutoSize = true;
            this.presentationPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.presentationPanel.BackColor = System.Drawing.Color.White;
            this.presentationPanel.Controls.Add(this.lblPresentationSection);
            this.presentationPanel.Controls.Add(this.presentationButtonsPanel);
            this.presentationPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.presentationPanel.Location = new System.Drawing.Point(3, 3);
            this.presentationPanel.MinimumSize = new System.Drawing.Size(0, 50);
            this.presentationPanel.Name = "presentationPanel";
            this.presentationPanel.Size = new System.Drawing.Size(240, 60);
            this.presentationPanel.TabIndex = 0;
            // 
            // lblPresentationSection
            // 
            this.lblPresentationSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblPresentationSection.ForeColor = System.Drawing.Color.Gray;
            this.lblPresentationSection.Location = new System.Drawing.Point(3, 0);
            this.lblPresentationSection.Name = "lblPresentationSection";
            this.lblPresentationSection.Size = new System.Drawing.Size(100, 23);
            this.lblPresentationSection.TabIndex = 0;
            this.lblPresentationSection.Text = "Presentation";
            // 
            // presentationButtonsPanel
            // 
            this.presentationButtonsPanel.AutoSize = true;
            this.presentationButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.presentationButtonsPanel.Controls.Add(this.btnNew);
            this.presentationButtonsPanel.Controls.Add(this.btnOpen);
            this.presentationButtonsPanel.Controls.Add(this.btnSave);
            this.presentationButtonsPanel.Controls.Add(this.btnSaveAs);
            this.presentationButtonsPanel.Controls.Add(this.btnPrint);
            this.presentationButtonsPanel.Controls.Add(this.btnShare);
            this.presentationButtonsPanel.Location = new System.Drawing.Point(3, 26);
            this.presentationButtonsPanel.Name = "presentationButtonsPanel";
            this.presentationButtonsPanel.Size = new System.Drawing.Size(186, 31);
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
            this.btnNew.Location = new System.Drawing.Point(3, 3);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(25, 25);
            this.btnNew.TabIndex = 1;
            this.btnNew.Text = "📄";
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
            this.btnOpen.ForeColor = System.Drawing.Color.Orange;
            this.btnOpen.Location = new System.Drawing.Point(34, 3);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(25, 25);
            this.btnOpen.TabIndex = 2;
            this.btnOpen.Text = "📂";
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
            this.btnSave.ForeColor = System.Drawing.Color.Orange;
            this.btnSave.Location = new System.Drawing.Point(65, 3);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(25, 25);
            this.btnSave.TabIndex = 3;
            this.btnSave.Text = "💾";
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
            this.btnSaveAs.ForeColor = System.Drawing.Color.Orange;
            this.btnSaveAs.Location = new System.Drawing.Point(96, 3);
            this.btnSaveAs.Name = "btnSaveAs";
            this.btnSaveAs.Size = new System.Drawing.Size(25, 25);
            this.btnSaveAs.TabIndex = 4;
            this.btnSaveAs.Text = "📋";
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
            this.btnPrint.ForeColor = System.Drawing.Color.Orange;
            this.btnPrint.Location = new System.Drawing.Point(127, 3);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(25, 25);
            this.btnPrint.TabIndex = 5;
            this.btnPrint.Text = "🖨";
            this.btnPrint.UseVisualStyleBackColor = false;
            this.btnPrint.Click += new System.EventHandler(this.BtnPrint_Click);
            // 
            // btnShare
            // 
            this.btnShare.BackColor = System.Drawing.Color.Transparent;
            this.btnShare.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnShare.FlatAppearance.BorderSize = 0;
            this.btnShare.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnShare.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnShare.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnShare.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnShare.ForeColor = System.Drawing.Color.Orange;
            this.btnShare.Location = new System.Drawing.Point(158, 3);
            this.btnShare.Name = "btnShare";
            this.btnShare.Size = new System.Drawing.Size(25, 25);
            this.btnShare.TabIndex = 6;
            this.btnShare.Text = "🤝";
            this.btnShare.UseVisualStyleBackColor = false;
            this.btnShare.Click += new System.EventHandler(this.BtnShare_Click);
            // 
            // wizardsPanel
            // 
            this.wizardsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.wizardsPanel.AutoSize = true;
            this.wizardsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.wizardsPanel.BackColor = System.Drawing.Color.White;
            this.wizardsPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.wizardsPanel.Controls.Add(this.lblWizardsSection);
            this.wizardsPanel.Controls.Add(this.wizardButtonsPanel);
            this.wizardsPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.wizardsPanel.MinimumSize = new System.Drawing.Size(2, 50);
            this.wizardsPanel.Name = "wizardsPanel";
            this.wizardsPanel.Size = new System.Drawing.Size(240, 50);
            this.wizardsPanel.TabIndex = 1;
            this.wizardsPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.wizardsPanel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            // 
            // lblWizardsSection
            // 
            this.lblWizardsSection.AutoSize = true;
            this.lblWizardsSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblWizardsSection.ForeColor = System.Drawing.Color.Gray;
            this.lblWizardsSection.Name = "lblWizardsSection";
            this.lblWizardsSection.Size = new System.Drawing.Size(51, 15);
            this.lblWizardsSection.TabIndex = 0;
            this.lblWizardsSection.Text = "Wizards";
            // 
            // wizardButtonsPanel
            // 
            this.wizardButtonsPanel.AutoSize = true;
            this.wizardButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.wizardButtonsPanel.Controls.Add(this.btnAgenda);
            this.wizardButtonsPanel.Controls.Add(this.btnMaster);
            this.wizardButtonsPanel.Controls.Add(this.btnElement);
            this.wizardButtonsPanel.Controls.Add(this.btnText);
            this.wizardButtonsPanel.Controls.Add(this.btnFormat);
            this.wizardButtonsPanel.Controls.Add(this.btnMap);
            this.wizardButtonsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wizardButtonsPanel.Name = "wizardButtonsPanel";
            this.wizardButtonsPanel.Size = new System.Drawing.Size(204, 31);
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
            this.btnAgenda.ForeColor = System.Drawing.Color.Orange;
            this.btnAgenda.Location = new System.Drawing.Point(3, 3);
            this.btnAgenda.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnAgenda.Name = "btnAgenda";
            this.btnAgenda.Size = new System.Drawing.Size(25, 25);
            this.btnAgenda.TabIndex = 1;
            this.btnAgenda.Text = "📋";
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
            this.btnMaster.ForeColor = System.Drawing.Color.Orange;
            this.btnMaster.Location = new System.Drawing.Point(37, 3);
            this.btnMaster.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnMaster.Name = "btnMaster";
            this.btnMaster.Size = new System.Drawing.Size(25, 25);
            this.btnMaster.TabIndex = 2;
            this.btnMaster.Text = "🎨";
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
            this.btnElement.ForeColor = System.Drawing.Color.Orange;
            this.btnElement.Location = new System.Drawing.Point(71, 3);
            this.btnElement.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnElement.Name = "btnElement";
            this.btnElement.Size = new System.Drawing.Size(25, 25);
            this.btnElement.TabIndex = 3;
            this.btnElement.Text = "🧩";
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
            this.btnText.ForeColor = System.Drawing.Color.Orange;
            this.btnText.Location = new System.Drawing.Point(105, 3);
            this.btnText.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnText.Name = "btnText";
            this.btnText.Size = new System.Drawing.Size(25, 25);
            this.btnText.TabIndex = 4;
            this.btnText.Text = "✏️";
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
            this.btnFormat.ForeColor = System.Drawing.Color.Orange;
            this.btnFormat.Location = new System.Drawing.Point(139, 3);
            this.btnFormat.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Size = new System.Drawing.Size(25, 25);
            this.btnFormat.TabIndex = 5;
            this.btnFormat.Text = "🎯";
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
            this.btnMap.ForeColor = System.Drawing.Color.Orange;
            this.btnMap.Location = new System.Drawing.Point(173, 3);
            this.btnMap.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnMap.Name = "btnMap";
            this.btnMap.Size = new System.Drawing.Size(25, 25);
            this.btnMap.TabIndex = 6;
            this.btnMap.Text = "🗺️";
            this.btnMap.UseVisualStyleBackColor = false;
            this.btnMap.Click += new System.EventHandler(this.BtnMap_Click);
            // 
            // smartElementsPanel
            // 
            this.smartElementsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.smartElementsPanel.AutoSize = true;
            this.smartElementsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.smartElementsPanel.BackColor = System.Drawing.Color.White;
            this.smartElementsPanel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.smartElementsPanel.Controls.Add(this.lblSmartElementsSection);
            this.smartElementsPanel.Controls.Add(this.smartElementsButtonsPanel);

            this.smartElementsPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.smartElementsPanel.Location = new System.Drawing.Point(3, 160);
            this.smartElementsPanel.MinimumSize = new System.Drawing.Size(0, 50);
            this.smartElementsPanel.Name = "smartElementsPanel";
            this.smartElementsPanel.Size = new System.Drawing.Size(240, 50);
            this.smartElementsPanel.TabIndex = 2;
            // 
            // lblSmartElementsSection
            // 
            this.lblSmartElementsSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblSmartElementsSection.ForeColor = System.Drawing.Color.Gray;
            this.lblSmartElementsSection.Location = new System.Drawing.Point(3, 0);
            this.lblSmartElementsSection.Name = "lblSmartElementsSection";
            this.lblSmartElementsSection.Size = new System.Drawing.Size(100, 23);
            this.lblSmartElementsSection.TabIndex = 0;
            this.lblSmartElementsSection.Text = "Smart Elements";
            // 
            // smartElementsButtonsPanel
            // 
            this.smartElementsButtonsPanel.AutoSize = true;
            this.smartElementsButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.smartElementsButtonsPanel.Controls.Add(this.btnChart);
            this.smartElementsButtonsPanel.Controls.Add(this.btnDiagram);
            this.smartElementsButtonsPanel.Controls.Add(this.btnTable);
            this.smartElementsButtonsPanel.Location = new System.Drawing.Point(3, 26);
            this.smartElementsButtonsPanel.Name = "smartElementsButtonsPanel";
            this.smartElementsButtonsPanel.Size = new System.Drawing.Size(102, 31);
            this.smartElementsButtonsPanel.TabIndex = 1;
            // 
            // btnChart
            // 
            this.btnChart.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnChart.FlatAppearance.BorderSize = 0;
            this.btnChart.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnChart.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnChart.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnChart.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnChart.ForeColor = System.Drawing.Color.Orange;
            this.btnChart.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnChart.Name = "btnChart";
            this.btnChart.Size = new System.Drawing.Size(25, 25);
            this.btnChart.TabIndex = 1;
            this.btnChart.Text = "📊";
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
            this.btnDiagram.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnDiagram.ForeColor = System.Drawing.Color.Orange;
            this.btnDiagram.Name = "btnDiagram";
            this.btnDiagram.Size = new System.Drawing.Size(25, 25);
            this.btnDiagram.TabIndex = 2;
            this.btnDiagram.Text = "🎨";
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
            this.btnTable.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnTable.ForeColor = System.Drawing.Color.Orange;
            this.btnTable.Name = "btnTable";
            this.btnTable.Size = new System.Drawing.Size(25, 25);
            this.btnTable.TabIndex = 3;
            this.btnTable.Text = "📋";
            this.btnTable.UseVisualStyleBackColor = false;
            this.btnTable.Click += new System.EventHandler(this.BtnTable_Click);
            // 
            // positionPanel
            // 
            this.positionPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.positionPanel.AutoSize = true;
            this.positionPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.positionPanel.BackColor = System.Drawing.Color.White;
            this.positionPanel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.positionPanel.Controls.Add(this.lblPositionSection);
            this.positionPanel.Controls.Add(this.positionButtonsPanel);
            this.positionPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.positionPanel.Location = new System.Drawing.Point(3, 251);
            this.positionPanel.MinimumSize = new System.Drawing.Size(0, 50);
            this.positionPanel.Name = "positionPanel";
            this.positionPanel.Size = new System.Drawing.Size(240, 50);
            this.positionPanel.TabIndex = 3;
            // 
            // lblPositionSection
            // 
            this.lblPositionSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblPositionSection.ForeColor = System.Drawing.Color.Gray;
            this.lblPositionSection.Location = new System.Drawing.Point(3, 0);
            this.lblPositionSection.Name = "lblPositionSection";
            this.lblPositionSection.Size = new System.Drawing.Size(100, 23);
            this.lblPositionSection.TabIndex = 0;
            this.lblPositionSection.Text = "Position";
            // 
            // positionButtonsPanel
            // 
            this.positionButtonsPanel.AutoSize = true;
            this.positionButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.positionButtonsPanel.Controls.Add(this.btnAlignLeft);
            this.positionButtonsPanel.Controls.Add(this.btnAlignCenter);
            this.positionButtonsPanel.Controls.Add(this.btnAlignRight);
            this.positionButtonsPanel.Controls.Add(this.btnDistribute);
            this.positionButtonsPanel.Location = new System.Drawing.Point(3, 26);
            this.positionButtonsPanel.Name = "positionButtonsPanel";
            this.positionButtonsPanel.Size = new System.Drawing.Size(136, 31);
            this.positionButtonsPanel.TabIndex = 1;
            // 
            // btnAlignLeft
            // 
            this.btnAlignLeft.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnAlignLeft.FlatAppearance.BorderSize = 0;
            this.btnAlignLeft.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnAlignLeft.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnAlignLeft.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlignLeft.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnAlignLeft.ForeColor = System.Drawing.Color.Orange;
            this.btnAlignLeft.Location = new System.Drawing.Point(3, 3);
            this.btnAlignLeft.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnAlignLeft.Name = "btnAlignLeft";
            this.btnAlignLeft.Size = new System.Drawing.Size(25, 25);
            this.btnAlignLeft.TabIndex = 1;
            this.btnAlignLeft.Text = "⬅️";
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
            this.btnAlignCenter.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnAlignCenter.ForeColor = System.Drawing.Color.Orange;
            this.btnAlignCenter.Location = new System.Drawing.Point(37, 3);
            this.btnAlignCenter.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnAlignCenter.Name = "btnAlignCenter";
            this.btnAlignCenter.Size = new System.Drawing.Size(25, 25);
            this.btnAlignCenter.TabIndex = 2;
            this.btnAlignCenter.Text = "⚖️";
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
            this.btnAlignRight.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnAlignRight.ForeColor = System.Drawing.Color.Orange;
            this.btnAlignRight.Location = new System.Drawing.Point(71, 3);
            this.btnAlignRight.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnAlignRight.Name = "btnAlignRight";
            this.btnAlignRight.Size = new System.Drawing.Size(25, 25);
            this.btnAlignRight.TabIndex = 3;
            this.btnAlignRight.Text = "➡️";
            this.btnAlignRight.UseVisualStyleBackColor = false;
            this.btnAlignRight.Click += new System.EventHandler(this.BtnAlignRight_Click);
            // 
            // btnDistribute
            // 
            this.btnDistribute.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnDistribute.FlatAppearance.BorderSize = 0;
            this.btnDistribute.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnDistribute.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnDistribute.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDistribute.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnDistribute.ForeColor = System.Drawing.Color.Orange;
            this.btnDistribute.Location = new System.Drawing.Point(105, 3);
            this.btnDistribute.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnDistribute.Name = "btnDistribute";
            this.btnDistribute.Size = new System.Drawing.Size(25, 25);
            this.btnDistribute.TabIndex = 4;
            this.btnDistribute.Text = "🔄";
            this.btnDistribute.UseVisualStyleBackColor = false;
            this.btnDistribute.Click += new System.EventHandler(this.BtnDistribute_Click);
            // 
            // sizePanel
            // 
            this.sizePanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.sizePanel.AutoSize = true;
            this.sizePanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.sizePanel.BackColor = System.Drawing.Color.White;
            this.sizePanel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.sizePanel.Controls.Add(this.lblSizeSection);
            this.sizePanel.Controls.Add(this.sizeControlsPanel);
            this.sizePanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.sizePanel.Location = new System.Drawing.Point(3, 342);
            this.sizePanel.MinimumSize = new System.Drawing.Size(0, 100);
            this.sizePanel.Name = "sizePanel";
            this.sizePanel.Size = new System.Drawing.Size(240, 130);
            this.sizePanel.TabIndex = 4;
            // 
            // lblSizeSection
            // 
            this.lblSizeSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblSizeSection.ForeColor = System.Drawing.Color.Gray;
            this.lblSizeSection.Location = new System.Drawing.Point(3, 0);
            this.lblSizeSection.Name = "lblSizeSection";
            this.lblSizeSection.Size = new System.Drawing.Size(100, 23);
            this.lblSizeSection.TabIndex = 0;
            this.lblSizeSection.Text = "Size";
            // 
            // lblSlideSize
            // 
            this.lblSlideSize.AutoSize = true;
            this.lblSlideSize.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblSlideSize.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.lblSlideSize.Location = new System.Drawing.Point(3, 0);
            this.lblSlideSize.Name = "lblSlideSize";
            this.lblSlideSize.Size = new System.Drawing.Size(58, 13);
            this.lblSlideSize.TabIndex = 0;
            this.lblSlideSize.Text = "Slide Size:";
            // 
            // cmbSlideSize
            // 
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
            this.cmbSlideSize.Location = new System.Drawing.Point(67, 3);
            this.cmbSlideSize.Name = "cmbSlideSize";
            this.cmbSlideSize.Size = new System.Drawing.Size(103, 21);
            this.cmbSlideSize.TabIndex = 1;
            this.cmbSlideSize.SelectedIndexChanged += new System.EventHandler(this.CmbSlideSize_SelectedIndexChanged);
            // 
            // nudWidth
            // 
            this.nudWidth.DecimalPlaces = 1;
            this.nudWidth.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.nudWidth.Location = new System.Drawing.Point(3, 58);
            this.nudWidth.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.nudWidth.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudWidth.Name = "nudWidth";
            this.nudWidth.Size = new System.Drawing.Size(43, 22);
            this.nudWidth.TabIndex = 2;
            this.nudWidth.Value = new decimal(new int[] {
            133,
            0,
            0,
            65536});
            // 
            // lblHeight
            // 
            this.lblHeight.AutoSize = true;
            this.lblHeight.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblHeight.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.lblHeight.Location = new System.Drawing.Point(109, 27);
            this.lblHeight.Name = "lblHeight";
            this.lblHeight.Size = new System.Drawing.Size(45, 13);
            this.lblHeight.TabIndex = 3;
            this.lblHeight.Text = "Height:";
            // 
            // nudHeight
            // 
            this.nudHeight.DecimalPlaces = 1;
            this.nudHeight.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.nudHeight.Location = new System.Drawing.Point(160, 30);
            this.nudHeight.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.nudHeight.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudHeight.Name = "nudHeight";
            this.nudHeight.Size = new System.Drawing.Size(43, 22);
            this.nudHeight.TabIndex = 4;
            this.nudHeight.Value = new decimal(new int[] {
            75,
            0,
            0,
            65536});
            // 
            // btnApplySize
            // 
            this.btnApplySize.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(120)))), ((int)(((byte)(215)))));
            this.btnApplySize.FlatAppearance.BorderSize = 0;
            this.btnApplySize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnApplySize.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.btnApplySize.ForeColor = System.Drawing.Color.White;
            this.btnApplySize.Location = new System.Drawing.Point(3, 86);
            this.btnApplySize.Name = "btnApplySize";
            this.btnApplySize.Size = new System.Drawing.Size(206, 22);
            this.btnApplySize.TabIndex = 5;
            this.btnApplySize.Text = "Apply Size";
            this.btnApplySize.UseVisualStyleBackColor = false;
            this.btnApplySize.Click += new System.EventHandler(this.BtnApplySize_Click);
            // 
            // shapePanel
            // 
            this.shapePanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.shapePanel.AutoSize = true;
            this.shapePanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.shapePanel.BackColor = System.Drawing.Color.White;
            this.shapePanel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.shapePanel.Controls.Add(this.lblShapeSection);
            this.shapePanel.Controls.Add(this.shapeButtonsPanel);
            this.shapePanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.shapePanel.Location = new System.Drawing.Point(3, 512);
            this.shapePanel.MinimumSize = new System.Drawing.Size(0, 50);
            this.shapePanel.Name = "shapePanel";
            this.shapePanel.Size = new System.Drawing.Size(240, 50);
            this.shapePanel.TabIndex = 5;
            // 
            // lblShapeSection
            // 
            this.lblShapeSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblShapeSection.ForeColor = System.Drawing.Color.Gray;
            this.lblShapeSection.Location = new System.Drawing.Point(3, 0);
            this.lblShapeSection.Name = "lblShapeSection";
            this.lblShapeSection.Size = new System.Drawing.Size(100, 23);
            this.lblShapeSection.TabIndex = 0;
            this.lblShapeSection.Text = "Shape";
            // 
            // shapeButtonsPanel
            // 
            this.shapeButtonsPanel.AutoSize = true;
            this.shapeButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.shapeButtonsPanel.Controls.Add(this.btnRectangle);
            this.shapeButtonsPanel.Controls.Add(this.btnCircle);
            this.shapeButtonsPanel.Controls.Add(this.btnArrow);
            this.shapeButtonsPanel.Controls.Add(this.btnLine);
            this.shapeButtonsPanel.Location = new System.Drawing.Point(3, 26);
            this.shapeButtonsPanel.Name = "shapeButtonsPanel";
            this.shapeButtonsPanel.Size = new System.Drawing.Size(172, 36);
            this.shapeButtonsPanel.TabIndex = 1;
            // 
            // btnRectangle
            // 
            this.btnRectangle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRectangle.FlatAppearance.BorderSize = 0;
            this.btnRectangle.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnRectangle.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnRectangle.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnRectangle.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnRectangle.ForeColor = System.Drawing.Color.Orange;
            this.btnRectangle.Location = new System.Drawing.Point(3, 3);
            this.btnRectangle.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnRectangle.Name = "btnRectangle";
            this.btnRectangle.Size = new System.Drawing.Size(25, 25);
            this.btnRectangle.TabIndex = 1;
            this.btnRectangle.Text = "▭";
            this.btnRectangle.UseVisualStyleBackColor = false;
            this.btnRectangle.Click += new System.EventHandler(this.BtnRectangle_Click);
            // 
            // btnCircle
            // 
            this.btnCircle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCircle.FlatAppearance.BorderSize = 0;
            this.btnCircle.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnCircle.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnCircle.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnCircle.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnCircle.ForeColor = System.Drawing.Color.Orange;
            this.btnCircle.Location = new System.Drawing.Point(34, 3);
            this.btnCircle.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnCircle.Name = "btnCircle";
            this.btnCircle.Size = new System.Drawing.Size(25, 25);
            this.btnCircle.TabIndex = 2;
            this.btnCircle.Text = "●";
            this.btnCircle.UseVisualStyleBackColor = false;
            this.btnCircle.Click += new System.EventHandler(this.BtnCircle_Click);
            // 
            // btnArrow
            // 
            this.btnArrow.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnArrow.FlatAppearance.BorderSize = 0;
            this.btnArrow.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnArrow.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnArrow.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnArrow.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnArrow.ForeColor = System.Drawing.Color.Orange;
            this.btnArrow.Location = new System.Drawing.Point(65, 3);
            this.btnArrow.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnArrow.Name = "btnArrow";
            this.btnArrow.Size = new System.Drawing.Size(25, 25);
            this.btnArrow.TabIndex = 3;
            this.btnArrow.Text = "→";
            this.btnArrow.UseVisualStyleBackColor = false;
            this.btnArrow.Click += new System.EventHandler(this.BtnArrow_Click);
            // 
            // btnLine
            // 
            this.btnLine.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnLine.FlatAppearance.BorderSize = 0;
            this.btnLine.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnLine.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.btnLine.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.btnLine.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnLine.ForeColor = System.Drawing.Color.Orange;
            this.btnLine.Location = new System.Drawing.Point(96, 3);
            this.btnLine.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnLine.Name = "btnLine";
            this.btnLine.Size = new System.Drawing.Size(25, 25);
            this.btnLine.TabIndex = 4;
            this.btnLine.Text = "─";
            this.btnLine.UseVisualStyleBackColor = false;
            this.btnLine.Click += new System.EventHandler(this.BtnLine_Click);
            // 
            // colorPanel
            // 
            this.colorPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.colorPanel.AutoSize = true;
            this.colorPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.colorPanel.BackColor = System.Drawing.Color.White;
            this.colorPanel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.colorPanel.Controls.Add(this.lblColorSection);
            this.colorPanel.Controls.Add(this.colorButtonsPanel);
            this.colorPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.colorPanel.Location = new System.Drawing.Point(3, 607);
            this.colorPanel.MinimumSize = new System.Drawing.Size(0, 50);
            this.colorPanel.Name = "colorPanel";
            this.colorPanel.Size = new System.Drawing.Size(240, 50);
            this.colorPanel.TabIndex = 6;
            // 
            // lblColorSection
            // 
            this.lblColorSection.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblColorSection.ForeColor = System.Drawing.Color.Gray;
            this.lblColorSection.Location = new System.Drawing.Point(3, 0);
            this.lblColorSection.Name = "lblColorSection";
            this.lblColorSection.Size = new System.Drawing.Size(100, 23);
            this.lblColorSection.TabIndex = 0;
            this.lblColorSection.Text = "Color";
            // 
            // colorButtonsPanel
            // 
            this.colorButtonsPanel.AutoSize = true;
            this.colorButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.colorButtonsPanel.Controls.Add(this.btnFillColor);
            this.colorButtonsPanel.Controls.Add(this.btnTextColor);
            this.colorButtonsPanel.Controls.Add(this.btnOutlineColor);
            this.colorButtonsPanel.Location = new System.Drawing.Point(3, 26);
            this.colorButtonsPanel.Name = "colorButtonsPanel";
            this.colorButtonsPanel.Size = new System.Drawing.Size(129, 36);
            this.colorButtonsPanel.TabIndex = 1;
            // 
            // btnFillColor
            // 
            this.btnFillColor.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnFillColor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFillColor.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnFillColor.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnFillColor.Location = new System.Drawing.Point(3, 3);
            this.btnFillColor.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnFillColor.Name = "btnFillColor";
            this.btnFillColor.Size = new System.Drawing.Size(34, 30);
            this.btnFillColor.TabIndex = 1;
            this.btnFillColor.Text = "🎨\nFill";
            this.btnFillColor.UseVisualStyleBackColor = false;
            this.btnFillColor.Click += new System.EventHandler(this.BtnFillColor_Click);
            // 
            // btnTextColor
            // 
            this.btnTextColor.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnTextColor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnTextColor.Font = new System.Drawing.Font("Segoe UI Emoji", 8F);
            this.btnTextColor.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnTextColor.Location = new System.Drawing.Point(46, 3);
            this.btnTextColor.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnTextColor.Name = "btnTextColor";
            this.btnTextColor.Size = new System.Drawing.Size(34, 30);
            this.btnTextColor.TabIndex = 2;
            this.btnTextColor.Text = "A\nText";
            this.btnTextColor.UseVisualStyleBackColor = false;
            this.btnTextColor.Click += new System.EventHandler(this.BtnTextColor_Click);
            // 
            // btnOutlineColor
            // 
            this.btnOutlineColor.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnOutlineColor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOutlineColor.Font = new System.Drawing.Font("Segoe UI Emoji", 7F);
            this.btnOutlineColor.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnOutlineColor.Location = new System.Drawing.Point(89, 3);
            this.btnOutlineColor.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnOutlineColor.Name = "btnOutlineColor";
            this.btnOutlineColor.Size = new System.Drawing.Size(34, 30);
            this.btnOutlineColor.TabIndex = 3;
            this.btnOutlineColor.Text = "◯\nOutline";
            this.btnOutlineColor.UseVisualStyleBackColor = false;
            this.btnOutlineColor.Click += new System.EventHandler(this.BtnOutlineColor_Click);
            // 
            // textPanel
            // 
            this.textPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.textPanel.AutoSize = true;
            this.textPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.textPanel.BackColor = System.Drawing.Color.White;
            this.textPanel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textPanel.Controls.Add(this.lblTextSection);
            this.textPanel.Controls.Add(this.textButtonsPanel);
            this.textPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.textPanel.Location = new System.Drawing.Point(3, 702);
            this.textPanel.Margin = new System.Windows.Forms.Padding(3, 3, 3, 8);
            this.textPanel.MinimumSize = new System.Drawing.Size(2, 80);
            this.textPanel.Name = "textPanel";
            this.textPanel.Padding = new System.Windows.Forms.Padding(10);
            this.textPanel.Size = new System.Drawing.Size(240, 84);
            this.textPanel.TabIndex = 7;
            // 
            // lblTextSection
            // 
            this.lblTextSection.AutoSize = true;
            this.lblTextSection.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblTextSection.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.lblTextSection.Location = new System.Drawing.Point(13, 10);
            this.lblTextSection.Margin = new System.Windows.Forms.Padding(3, 0, 3, 5);
            this.lblTextSection.Name = "lblTextSection";
            this.lblTextSection.Size = new System.Drawing.Size(32, 15);
            this.lblTextSection.TabIndex = 0;
            this.lblTextSection.Text = "Text";
            // 
            // textButtonsPanel
            // 
            this.textButtonsPanel.AutoSize = true;
            this.textButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.textButtonsPanel.Controls.Add(this.btnBold);
            this.textButtonsPanel.Controls.Add(this.btnItalic);
            this.textButtonsPanel.Controls.Add(this.btnUnderline);
            this.textButtonsPanel.Controls.Add(this.btnBullets);
            this.textButtonsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textButtonsPanel.Location = new System.Drawing.Point(13, 33);
            this.textButtonsPanel.Name = "textButtonsPanel";
            this.textButtonsPanel.Size = new System.Drawing.Size(172, 36);
            this.textButtonsPanel.TabIndex = 1;
            // 
            // btnBold
            // 
            this.btnBold.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnBold.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBold.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.btnBold.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnBold.Location = new System.Drawing.Point(3, 3);
            this.btnBold.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnBold.Name = "btnBold";
            this.btnBold.Size = new System.Drawing.Size(34, 30);
            this.btnBold.TabIndex = 1;
            this.btnBold.Text = "B";
            this.btnBold.UseVisualStyleBackColor = false;
            this.btnBold.Click += new System.EventHandler(this.BtnBold_Click);
            // 
            // btnItalic
            // 
            this.btnItalic.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnItalic.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnItalic.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            this.btnItalic.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnItalic.Location = new System.Drawing.Point(46, 3);
            this.btnItalic.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnItalic.Name = "btnItalic";
            this.btnItalic.Size = new System.Drawing.Size(34, 30);
            this.btnItalic.TabIndex = 2;
            this.btnItalic.Text = "I";
            this.btnItalic.UseVisualStyleBackColor = false;
            this.btnItalic.Click += new System.EventHandler(this.BtnItalic_Click);
            // 
            // btnUnderline
            // 
            this.btnUnderline.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnUnderline.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnUnderline.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Underline);
            this.btnUnderline.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnUnderline.Location = new System.Drawing.Point(89, 3);
            this.btnUnderline.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnUnderline.Name = "btnUnderline";
            this.btnUnderline.Size = new System.Drawing.Size(34, 30);
            this.btnUnderline.TabIndex = 3;
            this.btnUnderline.Text = "U";
            this.btnUnderline.UseVisualStyleBackColor = false;
            this.btnUnderline.Click += new System.EventHandler(this.BtnUnderline_Click);
            // 
            // btnBullets
            // 
            this.btnBullets.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnBullets.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBullets.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.btnBullets.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnBullets.Location = new System.Drawing.Point(132, 3);
            this.btnBullets.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnBullets.Name = "btnBullets";
            this.btnBullets.Size = new System.Drawing.Size(34, 30);
            this.btnBullets.TabIndex = 4;
            this.btnBullets.Text = "•";
            this.btnBullets.UseVisualStyleBackColor = false;
            this.btnBullets.Click += new System.EventHandler(this.BtnBullets_Click);
            // 
            // navigationPanel
            // 
            this.navigationPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.navigationPanel.AutoSize = true;
            this.navigationPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.navigationPanel.BackColor = System.Drawing.Color.White;
            this.navigationPanel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.navigationPanel.Controls.Add(this.lblNavigationSection);
            this.navigationPanel.Controls.Add(this.navigationButtonsPanel);
            this.navigationPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.navigationPanel.Location = new System.Drawing.Point(3, 797);
            this.navigationPanel.Margin = new System.Windows.Forms.Padding(3, 3, 3, 8);
            this.navigationPanel.MinimumSize = new System.Drawing.Size(2, 80);
            this.navigationPanel.Name = "navigationPanel";
            this.navigationPanel.Padding = new System.Windows.Forms.Padding(10);
            this.navigationPanel.Size = new System.Drawing.Size(240, 84);
            this.navigationPanel.TabIndex = 8;
            // 
            // lblNavigationSection
            // 
            this.lblNavigationSection.AutoSize = true;
            this.lblNavigationSection.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblNavigationSection.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.lblNavigationSection.Location = new System.Drawing.Point(13, 10);
            this.lblNavigationSection.Margin = new System.Windows.Forms.Padding(3, 0, 3, 5);
            this.lblNavigationSection.Name = "lblNavigationSection";
            this.lblNavigationSection.Size = new System.Drawing.Size(101, 15);
            this.lblNavigationSection.TabIndex = 0;
            this.lblNavigationSection.Text = "Navigation & View";
            // 
            // navigationButtonsPanel
            // 
            this.navigationButtonsPanel.AutoSize = true;
            this.navigationButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.navigationButtonsPanel.Controls.Add(this.btnZoomIn);
            this.navigationButtonsPanel.Controls.Add(this.btnZoomOut);
            this.navigationButtonsPanel.Controls.Add(this.btnFitToWindow);
            this.navigationButtonsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.navigationButtonsPanel.Location = new System.Drawing.Point(13, 33);
            this.navigationButtonsPanel.Name = "navigationButtonsPanel";
            this.navigationButtonsPanel.Size = new System.Drawing.Size(138, 36);
            this.navigationButtonsPanel.TabIndex = 1;
            // 
            // btnZoomIn
            // 
            this.btnZoomIn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnZoomIn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnZoomIn.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.btnZoomIn.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnZoomIn.Location = new System.Drawing.Point(3, 3);
            this.btnZoomIn.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnZoomIn.Name = "btnZoomIn";
            this.btnZoomIn.Size = new System.Drawing.Size(34, 30);
            this.btnZoomIn.TabIndex = 1;
            this.btnZoomIn.Text = "+";
            this.btnZoomIn.UseVisualStyleBackColor = false;
            this.btnZoomIn.Click += new System.EventHandler(this.BtnZoomIn_Click);
            // 
            // btnZoomOut
            // 
            this.btnZoomOut.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnZoomOut.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnZoomOut.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.btnZoomOut.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnZoomOut.Location = new System.Drawing.Point(46, 3);
            this.btnZoomOut.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnZoomOut.Name = "btnZoomOut";
            this.btnZoomOut.Size = new System.Drawing.Size(34, 30);
            this.btnZoomOut.TabIndex = 2;
            this.btnZoomOut.Text = "-";
            this.btnZoomOut.UseVisualStyleBackColor = false;
            this.btnZoomOut.Click += new System.EventHandler(this.BtnZoomOut_Click);
            // 
            // btnFitToWindow
            // 
            this.btnFitToWindow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnFitToWindow.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFitToWindow.Font = new System.Drawing.Font("Segoe UI", 7F);
            this.btnFitToWindow.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.btnFitToWindow.Location = new System.Drawing.Point(89, 3);
            this.btnFitToWindow.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnFitToWindow.Name = "btnFitToWindow";
            this.btnFitToWindow.Size = new System.Drawing.Size(43, 30);
            this.btnFitToWindow.TabIndex = 3;
            this.btnFitToWindow.Text = "🔍\nFit";
            this.btnFitToWindow.UseVisualStyleBackColor = false;
            this.btnFitToWindow.Click += new System.EventHandler(this.BtnFitToWindow_Click);
            // 
            // expertToolsPanel
            // 
            this.expertToolsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.expertToolsPanel.AutoSize = true;
            this.expertToolsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.expertToolsPanel.BackColor = System.Drawing.Color.White;
            this.expertToolsPanel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.expertToolsPanel.Controls.Add(this.lblExpertToolsSection);
            this.expertToolsPanel.Controls.Add(this.expertToolsButtonsPanel);
            this.expertToolsPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.expertToolsPanel.Location = new System.Drawing.Point(3, 892);
            this.expertToolsPanel.Margin = new System.Windows.Forms.Padding(3, 3, 3, 8);
            this.expertToolsPanel.MinimumSize = new System.Drawing.Size(2, 80);
            this.expertToolsPanel.Name = "expertToolsPanel";
            this.expertToolsPanel.Padding = new System.Windows.Forms.Padding(10);
            this.expertToolsPanel.Size = new System.Drawing.Size(240, 84);
            this.expertToolsPanel.TabIndex = 9;
            // 
            // lblExpertToolsSection
            // 
            this.lblExpertToolsSection.AutoSize = true;
            this.lblExpertToolsSection.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblExpertToolsSection.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(68)))), ((int)(((byte)(68)))));
            this.lblExpertToolsSection.Location = new System.Drawing.Point(13, 10);
            this.lblExpertToolsSection.Margin = new System.Windows.Forms.Padding(3, 0, 3, 5);
            this.lblExpertToolsSection.Name = "lblExpertToolsSection";
            this.lblExpertToolsSection.Size = new System.Drawing.Size(75, 15);
            this.lblExpertToolsSection.TabIndex = 0;
            this.lblExpertToolsSection.Text = "Expert Tools";
            // 
            // expertToolsButtonsPanel
            // 
            this.expertToolsButtonsPanel.AutoSize = true;
            this.expertToolsButtonsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.expertToolsButtonsPanel.Controls.Add(this.btnFreeWebinar);
            this.expertToolsButtonsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.expertToolsButtonsPanel.Location = new System.Drawing.Point(13, 33);
            this.expertToolsButtonsPanel.Name = "expertToolsButtonsPanel";
            this.expertToolsButtonsPanel.Size = new System.Drawing.Size(180, 36);
            this.expertToolsButtonsPanel.TabIndex = 1;
            // 
            // btnFreeWebinar
            // 
            this.btnFreeWebinar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(120)))), ((int)(((byte)(215)))));
            this.btnFreeWebinar.FlatAppearance.BorderSize = 0;
            this.btnFreeWebinar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFreeWebinar.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnFreeWebinar.ForeColor = System.Drawing.Color.White;
            this.btnFreeWebinar.Location = new System.Drawing.Point(3, 3);
            this.btnFreeWebinar.Margin = new System.Windows.Forms.Padding(3, 3, 6, 3);
            this.btnFreeWebinar.Name = "btnFreeWebinar";
            this.btnFreeWebinar.Size = new System.Drawing.Size(171, 30);
            this.btnFreeWebinar.TabIndex = 1;
            this.btnFreeWebinar.Text = "🎓 Free PowerPoint Webinar";
            this.btnFreeWebinar.UseVisualStyleBackColor = false;
            this.btnFreeWebinar.Click += new System.EventHandler(this.BtnFreeWebinar_Click);
            // 
            // lblWidth
            // 
            this.lblWidth.Location = new System.Drawing.Point(3, 27);
            this.lblWidth.Name = "lblWidth";
            this.lblWidth.Size = new System.Drawing.Size(100, 23);
            this.lblWidth.TabIndex = 2;
            this.lblWidth.Click += new System.EventHandler(this.lblWidth_Click);
            // 
            // sizeControlsPanel
            // 
            this.sizeControlsPanel.AutoSize = true;
            this.sizeControlsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.sizeControlsPanel.Controls.Add(this.lblSlideSize);
            this.sizeControlsPanel.Controls.Add(this.cmbSlideSize);
            this.sizeControlsPanel.Controls.Add(this.lblWidth);
            this.sizeControlsPanel.Controls.Add(this.lblHeight);
            this.sizeControlsPanel.Controls.Add(this.nudHeight);
            this.sizeControlsPanel.Controls.Add(this.nudWidth);
            this.sizeControlsPanel.Controls.Add(this.btnApplySize);
            this.sizeControlsPanel.Location = new System.Drawing.Point(3, 26);
            this.sizeControlsPanel.Name = "sizeControlsPanel";
            this.sizeControlsPanel.Size = new System.Drawing.Size(212, 111);
            this.sizeControlsPanel.TabIndex = 1;
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
            ((System.ComponentModel.ISupportInitialize)(this.nudWidth)).EndInit();
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
            this.sizeControlsPanel.ResumeLayout(false);
            this.sizeControlsPanel.PerformLayout();
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
        private System.Windows.Forms.Button btnShare;
        
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
        
        // Position section
        private System.Windows.Forms.FlowLayoutPanel positionPanel;
        private System.Windows.Forms.FlowLayoutPanel positionButtonsPanel;
        private System.Windows.Forms.Label lblPositionSection;
        private System.Windows.Forms.Button btnAlignLeft;
        private System.Windows.Forms.Button btnAlignCenter;
        private System.Windows.Forms.Button btnAlignRight;
        private System.Windows.Forms.Button btnDistribute;
        
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
        private System.Windows.Forms.Button btnRectangle;
        private System.Windows.Forms.Button btnCircle;
        private System.Windows.Forms.Button btnArrow;
        private System.Windows.Forms.Button btnLine;
        
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
    }
} 