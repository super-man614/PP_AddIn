using System;
using System.Drawing;
using System.Windows.Forms;

namespace my_addin
{
    public partial class MatrixTableDialog : Form
    {
        public int Rows { get; private set; } = 3;
        public int Columns { get; private set; } = 3;
        public bool HasHeader { get; private set; } = true;

        private NumericUpDown nudRows;
        private NumericUpDown nudColumns;
        private CheckBox chkHeader;
        private Button btnOK;
        private Button btnCancel;
        private Label lblRows;
        private Label lblColumns;
        private Label lblTitle;
        private Panel separatorPanel;

        public MatrixTableDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.nudRows = new NumericUpDown();
            this.nudColumns = new NumericUpDown();
            this.chkHeader = new CheckBox();
            this.btnOK = new Button();
            this.btnCancel = new Button();
            this.lblRows = new Label();
            this.lblColumns = new Label();
            this.lblTitle = new Label();
            this.separatorPanel = new Panel();
            
            ((System.ComponentModel.ISupportInitialize)(this.nudRows)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudColumns)).BeginInit();
            this.SuspendLayout();
            
            // Form properties
            this.Text = "Create Professional Table";
            this.Size = new Size(320, 260); // Increased height from 220 to 260
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.BackColor = Color.White;
            
            // lblTitle
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            this.lblTitle.ForeColor = Color.FromArgb(68, 114, 196);
            this.lblTitle.Location = new Point(12, 15);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new Size(150, 21);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Table Configuration";
            
            // separatorPanel
            this.separatorPanel.BackColor = Color.FromArgb(230, 230, 230);
            this.separatorPanel.Location = new Point(12, 45);
            this.separatorPanel.Name = "separatorPanel";
            this.separatorPanel.Size = new Size(285, 1);
            this.separatorPanel.TabIndex = 1;
            
            // lblRows
            this.lblRows.AutoSize = true;
            this.lblRows.Font = new Font("Segoe UI", 9F);
            this.lblRows.Location = new Point(20, 65);
            this.lblRows.Name = "lblRows";
            this.lblRows.Size = new Size(37, 15);
            this.lblRows.TabIndex = 2;
            this.lblRows.Text = "Rows:";
            
            // nudRows
            this.nudRows.Font = new Font("Segoe UI", 9F);
            this.nudRows.Location = new Point(100, 63);
            this.nudRows.Minimum = 1;
            this.nudRows.Maximum = 20;
            this.nudRows.Value = 3;
            this.nudRows.Name = "nudRows";
            this.nudRows.Size = new Size(80, 23);
            this.nudRows.TabIndex = 3;
            
            // lblColumns
            this.lblColumns.AutoSize = true;
            this.lblColumns.Font = new Font("Segoe UI", 9F);
            this.lblColumns.Location = new Point(20, 95);
            this.lblColumns.Name = "lblColumns";
            this.lblColumns.Size = new Size(58, 15);
            this.lblColumns.TabIndex = 4;
            this.lblColumns.Text = "Columns:";
            
            // nudColumns
            this.nudColumns.Font = new Font("Segoe UI", 9F);
            this.nudColumns.Location = new Point(100, 93);
            this.nudColumns.Minimum = 1;
            this.nudColumns.Maximum = 20;
            this.nudColumns.Value = 3;
            this.nudColumns.Name = "nudColumns";
            this.nudColumns.Size = new Size(80, 23);
            this.nudColumns.TabIndex = 5;
            
            // chkHeader
            this.chkHeader.AutoSize = true;
            this.chkHeader.Checked = true;
            this.chkHeader.CheckState = CheckState.Checked;
            this.chkHeader.Font = new Font("Segoe UI", 9F);
            this.chkHeader.Location = new Point(20, 125);
            this.chkHeader.Name = "chkHeader";
            this.chkHeader.Size = new Size(162, 19);
            this.chkHeader.TabIndex = 6;
            this.chkHeader.Text = "Include header row (styled)";
            this.chkHeader.UseVisualStyleBackColor = true;
            
            // btnOK
            this.btnOK.BackColor = Color.FromArgb(68, 114, 196);
            this.btnOK.DialogResult = DialogResult.OK;
            this.btnOK.FlatStyle = FlatStyle.Flat;
            this.btnOK.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            this.btnOK.ForeColor = Color.White;
            this.btnOK.Location = new Point(70, 190); // Moved down from 160 to 190
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new Size(80, 30);
            this.btnOK.TabIndex = 7;
            this.btnOK.Text = "Create";
            this.btnOK.UseVisualStyleBackColor = false;
            this.btnOK.Click += new EventHandler(this.BtnOK_Click);
            
            // btnCancel
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.FlatStyle = FlatStyle.Flat;
            this.btnCancel.Font = new Font("Segoe UI", 9F);
            this.btnCancel.Location = new Point(160, 190); // Moved down from 160 to 190
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new Size(75, 30);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            
            // MatrixTableDialog
            this.AcceptButton = this.btnOK;
            this.CancelButton = this.btnCancel;
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.separatorPanel);
            this.Controls.Add(this.lblRows);
            this.Controls.Add(this.nudRows);
            this.Controls.Add(this.lblColumns);
            this.Controls.Add(this.nudColumns);
            this.Controls.Add(this.chkHeader);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            
            ((System.ComponentModel.ISupportInitialize)(this.nudRows)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudColumns)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            this.Rows = (int)nudRows.Value;
            this.Columns = (int)nudColumns.Value;
            this.HasHeader = chkHeader.Checked;
        }
    }
} 