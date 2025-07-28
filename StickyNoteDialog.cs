using System;
using System.Drawing;
using System.Windows.Forms;

namespace my_addin
{
    public partial class StickyNoteDialog : Form
    {
        public string NoteText { get; private set; } = "";
        public Color NoteColor { get; private set; } = Color.FromArgb(255, 255, 102); // Default yellow

        private TextBox txtNote;
        private Button btnOK;
        private Button btnCancel;
        private Label lblNote;
        private Label lblColor;
        private Panel pnlYellow;
        private Panel pnlPink;
        private Panel pnlBlue;
        private Panel pnlGreen;
        private Panel selectedColorPanel;

        public StickyNoteDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.txtNote = new TextBox();
            this.btnOK = new Button();
            this.btnCancel = new Button();
            this.lblNote = new Label();
            this.lblColor = new Label();
            this.pnlYellow = new Panel();
            this.pnlPink = new Panel();
            this.pnlBlue = new Panel();
            this.pnlGreen = new Panel();
            
            this.SuspendLayout();
            
            // Form properties
            this.Text = "Add Sticky Note";
            this.Size = new Size(350, 320); // Increased height from 280 to 320
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.BackColor = Color.White;
            
            // lblNote
            this.lblNote.AutoSize = true;
            this.lblNote.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            this.lblNote.Location = new Point(15, 15);
            this.lblNote.Name = "lblNote";
            this.lblNote.Size = new Size(35, 15);
            this.lblNote.TabIndex = 0;
            this.lblNote.Text = "Note:";
            
            // txtNote
            this.txtNote.Font = new Font("Segoe UI", 9F);
            this.txtNote.Location = new Point(15, 35);
            this.txtNote.Multiline = true;
            this.txtNote.Name = "txtNote";
            this.txtNote.ScrollBars = ScrollBars.Vertical;
            this.txtNote.Size = new Size(305, 100);
            this.txtNote.TabIndex = 1;
            this.txtNote.Text = "Add your note here...";
            this.txtNote.Enter += new EventHandler(this.TxtNote_Enter);
            
            // lblColor
            this.lblColor.AutoSize = true;
            this.lblColor.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            this.lblColor.Location = new Point(15, 150);
            this.lblColor.Name = "lblColor";
            this.lblColor.Size = new Size(39, 15);
            this.lblColor.TabIndex = 2;
            this.lblColor.Text = "Color:";
            
            // Color panels
            this.pnlYellow.BackColor = Color.FromArgb(255, 255, 102);
            this.pnlYellow.BorderStyle = BorderStyle.FixedSingle;
            this.pnlYellow.Cursor = Cursors.Hand;
            this.pnlYellow.Location = new Point(15, 175);
            this.pnlYellow.Name = "pnlYellow";
            this.pnlYellow.Size = new Size(30, 30);
            this.pnlYellow.TabIndex = 3;
            this.pnlYellow.Click += new EventHandler(this.ColorPanel_Click);
            
            this.pnlPink.BackColor = Color.FromArgb(255, 182, 193);
            this.pnlPink.BorderStyle = BorderStyle.FixedSingle;
            this.pnlPink.Cursor = Cursors.Hand;
            this.pnlPink.Location = new Point(55, 175);
            this.pnlPink.Name = "pnlPink";
            this.pnlPink.Size = new Size(30, 30);
            this.pnlPink.TabIndex = 4;
            this.pnlPink.Click += new EventHandler(this.ColorPanel_Click);
            
            this.pnlBlue.BackColor = Color.FromArgb(173, 216, 230);
            this.pnlBlue.BorderStyle = BorderStyle.FixedSingle;
            this.pnlBlue.Cursor = Cursors.Hand;
            this.pnlBlue.Location = new Point(95, 175);
            this.pnlBlue.Name = "pnlBlue";
            this.pnlBlue.Size = new Size(30, 30);
            this.pnlBlue.TabIndex = 5;
            this.pnlBlue.Click += new EventHandler(this.ColorPanel_Click);
            
            this.pnlGreen.BackColor = Color.FromArgb(144, 238, 144);
            this.pnlGreen.BorderStyle = BorderStyle.FixedSingle;
            this.pnlGreen.Cursor = Cursors.Hand;
            this.pnlGreen.Location = new Point(135, 175);
            this.pnlGreen.Name = "pnlGreen";
            this.pnlGreen.Size = new Size(30, 30);
            this.pnlGreen.TabIndex = 6;
            this.pnlGreen.Click += new EventHandler(this.ColorPanel_Click);
            
            // Set default selection (yellow)
            this.selectedColorPanel = this.pnlYellow;
            this.pnlYellow.BorderStyle = BorderStyle.Fixed3D;
            
            // btnOK
            this.btnOK.BackColor = Color.FromArgb(68, 114, 196);
            this.btnOK.DialogResult = DialogResult.OK;
            this.btnOK.FlatStyle = FlatStyle.Flat;
            this.btnOK.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            this.btnOK.ForeColor = Color.White;
            this.btnOK.Location = new Point(170, 250); // Moved down from 220 to 250
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new Size(75, 30);
            this.btnOK.TabIndex = 7;
            this.btnOK.Text = "Add Note";
            this.btnOK.UseVisualStyleBackColor = false;
            this.btnOK.Click += new EventHandler(this.BtnOK_Click);
            
            // btnCancel
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.FlatStyle = FlatStyle.Flat;
            this.btnCancel.Font = new Font("Segoe UI", 9F);
            this.btnCancel.Location = new Point(255, 250); // Moved down from 220 to 250
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new Size(75, 30);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            
            // StickyNoteDialog
            this.AcceptButton = this.btnOK;
            this.CancelButton = this.btnCancel;
            this.Controls.Add(this.lblNote);
            this.Controls.Add(this.txtNote);
            this.Controls.Add(this.lblColor);
            this.Controls.Add(this.pnlYellow);
            this.Controls.Add(this.pnlPink);
            this.Controls.Add(this.pnlBlue);
            this.Controls.Add(this.pnlGreen);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void TxtNote_Enter(object sender, EventArgs e)
        {
            if (txtNote.Text == "Add your note here...")
            {
                txtNote.Text = "";
                txtNote.ForeColor = Color.Black;
            }
        }

        private void ColorPanel_Click(object sender, EventArgs e)
        {
            // Reset previous selection
            if (selectedColorPanel != null)
            {
                selectedColorPanel.BorderStyle = BorderStyle.FixedSingle;
            }
            
            // Set new selection
            var panel = sender as Panel;
            panel.BorderStyle = BorderStyle.Fixed3D;
            selectedColorPanel = panel;
            
            // Update note color
            this.NoteColor = panel.BackColor;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            this.NoteText = txtNote.Text.Trim();
            if (this.NoteText == "Add your note here..." || string.IsNullOrEmpty(this.NoteText))
            {
                this.NoteText = "Sticky Note";
            }
        }
    }
} 