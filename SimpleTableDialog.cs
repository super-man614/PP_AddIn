using System;
using System.Drawing;
using System.Windows.Forms;

namespace my_addin
{
    public partial class SimpleTableDialog : Form
    {
        public int Rows { get; private set; } = 2;
        public int Columns { get; private set; } = 5;

        private Label lblColumns;
        private Label lblRows;
        private NumericUpDown nudColumns;
        private NumericUpDown nudRows;
        private Button btnOK;
        private Button btnCancel;

        public SimpleTableDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form properties
            this.Text = "Insert Table";
            this.Size = new Size(240, 160);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.BackColor = SystemColors.Control;
            
            // lblColumns
            this.lblColumns = new Label();
            this.lblColumns.Location = new Point(15, 20);
            this.lblColumns.Size = new Size(100, 15);
            this.lblColumns.Text = "Number of columns:";
            this.lblColumns.Font = new Font("Segoe UI", 9F);
            this.Controls.Add(this.lblColumns);
            
            // nudColumns
            this.nudColumns = new NumericUpDown();
            this.nudColumns.Location = new Point(140, 18);
            this.nudColumns.Size = new Size(60, 23);
            this.nudColumns.Minimum = 1;
            this.nudColumns.Maximum = 63;
            this.nudColumns.Value = 5;
            this.nudColumns.Font = new Font("Segoe UI", 9F);
            this.Controls.Add(this.nudColumns);
            
            // lblRows
            this.lblRows = new Label();
            this.lblRows.Location = new Point(15, 50);
            this.lblRows.Size = new Size(100, 15);
            this.lblRows.Text = "Number of rows:";
            this.lblRows.Font = new Font("Segoe UI", 9F);
            this.Controls.Add(this.lblRows);
            
            // nudRows
            this.nudRows = new NumericUpDown();
            this.nudRows.Location = new Point(140, 48);
            this.nudRows.Size = new Size(60, 23);
            this.nudRows.Minimum = 1;
            this.nudRows.Maximum = 100;
            this.nudRows.Value = 2;
            this.nudRows.Font = new Font("Segoe UI", 9F);
            this.Controls.Add(this.nudRows);
            
            // btnOK
            this.btnOK = new Button();
            this.btnOK.Location = new Point(60, 90);
            this.btnOK.Size = new Size(60, 25);
            this.btnOK.Text = "OK";
            this.btnOK.Font = new Font("Segoe UI", 9F);
            this.btnOK.DialogResult = DialogResult.OK;
            this.btnOK.Click += BtnOK_Click;
            this.Controls.Add(this.btnOK);
            
            // btnCancel
            this.btnCancel = new Button();
            this.btnCancel.Location = new Point(130, 90);
            this.btnCancel.Size = new Size(60, 25);
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Font = new Font("Segoe UI", 9F);
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.Controls.Add(this.btnCancel);
            
            // Set default button
            this.AcceptButton = this.btnOK;
            this.CancelButton = this.btnCancel;
            
            this.ResumeLayout(false);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            this.Columns = (int)nudColumns.Value;
            this.Rows = (int)nudRows.Value;
        }
    }
} 