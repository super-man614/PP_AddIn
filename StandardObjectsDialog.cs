using System;
using System.Drawing;
using System.Windows.Forms;

namespace my_addin
{
    public partial class StandardObjectsDialog : Form
    {
        public string SelectedObject { get; private set; } = "";

        private ListBox lstObjects;
        private Button btnOK;
        private Button btnCancel;
        private Label lblTitle;
        private Label lblDescription;

        public StandardObjectsDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.lstObjects = new ListBox();
            this.btnOK = new Button();
            this.btnCancel = new Button();
            this.lblTitle = new Label();
            this.lblDescription = new Label();
            
            this.SuspendLayout();
            
            // Form properties
            this.Text = "Standard Objects Library";
            this.Size = new Size(400, 420); // Increased height from 380 to 420
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
            this.lblTitle.Location = new Point(15, 15);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new Size(200, 21);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Frequently Used Objects";
            
            // lblDescription
            this.lblDescription.AutoSize = true;
            this.lblDescription.Font = new Font("Segoe UI", 9F);
            this.lblDescription.ForeColor = Color.FromArgb(100, 100, 100);
            this.lblDescription.Location = new Point(15, 45);
            this.lblDescription.Name = "lblDescription";
            this.lblDescription.Size = new Size(300, 15);
            this.lblDescription.TabIndex = 1;
            this.lblDescription.Text = "Select a standard object to quickly add to your slide:";
            
            // lstObjects
            this.lstObjects.Font = new Font("Segoe UI", 9F);
            this.lstObjects.ItemHeight = 20;
            this.lstObjects.Location = new Point(15, 75);
            this.lstObjects.Name = "lstObjects";
            this.lstObjects.Size = new Size(355, 260); // Increased height from 240 to 260
            this.lstObjects.TabIndex = 2;
            this.lstObjects.DoubleClick += new EventHandler(this.LstObjects_DoubleClick);
            
            // Populate with standard objects
            PopulateStandardObjects();
            
            // btnOK
            this.btnOK.BackColor = Color.FromArgb(68, 114, 196);
            this.btnOK.DialogResult = DialogResult.OK;
            this.btnOK.FlatStyle = FlatStyle.Flat;
            this.btnOK.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            this.btnOK.ForeColor = Color.White;
            this.btnOK.Location = new Point(210, 350); // Moved down from 330 to 350
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new Size(80, 30);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "Add Object";
            this.btnOK.UseVisualStyleBackColor = false;
            this.btnOK.Click += new EventHandler(this.BtnOK_Click);
            
            // btnCancel
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.FlatStyle = FlatStyle.Flat;
            this.btnCancel.Font = new Font("Segoe UI", 9F);
            this.btnCancel.Location = new Point(300, 350); // Moved down from 330 to 350
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new Size(75, 30);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            
            // StandardObjectsDialog
            this.AcceptButton = this.btnOK;
            this.CancelButton = this.btnCancel;
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.lblDescription);
            this.Controls.Add(this.lstObjects);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void PopulateStandardObjects()
        {
            // Add frequently used objects
            lstObjects.Items.Add("📋 Title & Subtitle Layout");
            lstObjects.Items.Add("🏷️ Header Text Box");
            lstObjects.Items.Add("📝 Content Text Box");
            lstObjects.Items.Add("💡 Callout Box");
            lstObjects.Items.Add("⚠️ Warning Box");
            lstObjects.Items.Add("✅ Success Box");
            lstObjects.Items.Add("ℹ️ Information Box");
            lstObjects.Items.Add("🎯 Objective Box");
            lstObjects.Items.Add("📊 Data Box");
            lstObjects.Items.Add("🔗 Link Button");
            lstObjects.Items.Add("🏢 Company Logo Placeholder");
            lstObjects.Items.Add("👤 Contact Info Box");
            lstObjects.Items.Add("📅 Date Stamp");
            lstObjects.Items.Add("🔢 Page Number");
            lstObjects.Items.Add("➡️ Navigation Arrow");
            lstObjects.Items.Add("🎨 Decorative Line");
            lstObjects.Items.Add("📱 Device Frame (Phone)");
            lstObjects.Items.Add("💻 Device Frame (Laptop)");
            lstObjects.Items.Add("🖼️ Image Placeholder");
            lstObjects.Items.Add("📈 Growth Arrow");
            
            // Select first item by default
            if (lstObjects.Items.Count > 0)
            {
                lstObjects.SelectedIndex = 0;
            }
        }

        private void LstObjects_DoubleClick(object sender, EventArgs e)
        {
            if (lstObjects.SelectedIndex >= 0)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (lstObjects.SelectedIndex >= 0)
            {
                this.SelectedObject = lstObjects.SelectedItem.ToString();
            }
        }
    }
} 