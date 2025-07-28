using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace my_addin
{
    public partial class TableDropdownControl : Form
    {
        public int SelectedRows { get; private set; } = 0;
        public int SelectedColumns { get; private set; } = 0;
        public string ActionType { get; private set; } = ""; // "GridSelect", "InsertTable", "ExcelSpreadsheet"

        private const int MaxRows = 8;
        private const int MaxCols = 10;
        private const int CellSize = 18;
        private const int CellMargin = 1;

        private Button[,] gridButtons;
        private Label lblPreview;
        private Button btnInsertTable;
        private Button btnExcelSpreadsheet;
        private int hoverRows = 0;
        private int hoverCols = 0;

        public TableDropdownControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form properties
            this.Text = "";
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
            this.TopMost = true;
            this.Size = new Size(220, 260);
            this.BackColor = Color.White;
            this.ControlBox = false;
            
            // Create grid of buttons for table selection
            CreateTableGrid();
            
            // Create preview label
            this.lblPreview = new Label();
            this.lblPreview.Location = new Point(10, 155);
            this.lblPreview.Size = new Size(200, 20);
            this.lblPreview.Text = "1 x 1 Table";
            this.lblPreview.Font = new Font("Segoe UI", 9F);
            this.lblPreview.ForeColor = Color.FromArgb(68, 114, 196);
            this.lblPreview.TextAlign = ContentAlignment.MiddleCenter;
            this.Controls.Add(this.lblPreview);
            
            // Create separator line
            var separator = new Panel();
            separator.Location = new Point(10, 180);
            separator.Size = new Size(190, 1);
            separator.BackColor = Color.FromArgb(200, 200, 200);
            this.Controls.Add(separator);
            
            // Create Insert Table button
            this.btnInsertTable = new Button();
            this.btnInsertTable.Location = new Point(10, 190);
            this.btnInsertTable.Size = new Size(190, 25);
            this.btnInsertTable.Text = "Insert Table...";
            this.btnInsertTable.Font = new Font("Segoe UI", 9F);
            this.btnInsertTable.FlatStyle = FlatStyle.Flat;
            this.btnInsertTable.FlatAppearance.BorderSize = 0;
            this.btnInsertTable.BackColor = Color.White;
            this.btnInsertTable.TextAlign = ContentAlignment.MiddleLeft;
            this.btnInsertTable.Image = CreateTableIcon();
            this.btnInsertTable.ImageAlign = ContentAlignment.MiddleLeft;
            this.btnInsertTable.TextImageRelation = TextImageRelation.ImageBeforeText;
            this.btnInsertTable.Click += BtnInsertTable_Click;
            this.btnInsertTable.MouseEnter += (s, e) => btnInsertTable.BackColor = Color.FromArgb(240, 240, 240);
            this.btnInsertTable.MouseLeave += (s, e) => btnInsertTable.BackColor = Color.White;
            this.Controls.Add(this.btnInsertTable);
            
            // Create Excel Spreadsheet button
            this.btnExcelSpreadsheet = new Button();
            this.btnExcelSpreadsheet.Location = new Point(10, 220);
            this.btnExcelSpreadsheet.Size = new Size(190, 25);
            this.btnExcelSpreadsheet.Text = "Excel Spreadsheet";
            this.btnExcelSpreadsheet.Font = new Font("Segoe UI", 9F);
            this.btnExcelSpreadsheet.FlatStyle = FlatStyle.Flat;
            this.btnExcelSpreadsheet.FlatAppearance.BorderSize = 0;
            this.btnExcelSpreadsheet.BackColor = Color.White;
            this.btnExcelSpreadsheet.TextAlign = ContentAlignment.MiddleLeft;
            this.btnExcelSpreadsheet.Image = CreateExcelIcon();
            this.btnExcelSpreadsheet.ImageAlign = ContentAlignment.MiddleLeft;
            this.btnExcelSpreadsheet.TextImageRelation = TextImageRelation.ImageBeforeText;
            this.btnExcelSpreadsheet.Click += BtnExcelSpreadsheet_Click;
            this.btnExcelSpreadsheet.MouseEnter += (s, e) => btnExcelSpreadsheet.BackColor = Color.FromArgb(240, 240, 240);
            this.btnExcelSpreadsheet.MouseLeave += (s, e) => btnExcelSpreadsheet.BackColor = Color.White;
            this.Controls.Add(this.btnExcelSpreadsheet);
            
            // Handle form deactivation to close dropdown
            this.Deactivate += (s, e) => this.Hide();
            
            this.ResumeLayout(false);
        }

        private void CreateTableGrid()
        {
            gridButtons = new Button[MaxRows, MaxCols];
            int startX = 10;
            int startY = 10;
            
            for (int row = 0; row < MaxRows; row++)
            {
                for (int col = 0; col < MaxCols; col++)
                {
                    var btn = new Button();
                    btn.Size = new Size(CellSize, CellSize);
                    btn.Location = new Point(
                        startX + col * (CellSize + CellMargin),
                        startY + row * (CellSize + CellMargin));
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.FlatAppearance.BorderSize = 1;
                    btn.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
                    btn.BackColor = Color.White;
                    btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(68, 114, 196);
                    btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(68, 114, 196);
                    
                    // Store row and col in Tag
                    btn.Tag = new Point(col, row);
                    
                    // Event handlers
                    btn.MouseEnter += GridButton_MouseEnter;
                    btn.MouseLeave += GridButton_MouseLeave;
                    btn.Click += GridButton_Click;
                    
                    gridButtons[row, col] = btn;
                    this.Controls.Add(btn);
                }
            }
        }

        private void GridButton_MouseEnter(object sender, EventArgs e)
        {
            var btn = sender as Button;
            var pos = (Point)btn.Tag;
            hoverCols = pos.X + 1;
            hoverRows = pos.Y + 1;
            
            UpdateGridHighlight();
            UpdatePreviewText();
        }

        private void GridButton_MouseLeave(object sender, EventArgs e)
        {
            // Only clear if mouse leaves the entire grid area
        }

        private void GridButton_Click(object sender, EventArgs e)
        {
            var btn = sender as Button;
            var pos = (Point)btn.Tag;
            
            SelectedColumns = pos.X + 1;
            SelectedRows = pos.Y + 1;
            ActionType = "GridSelect";
            
            this.DialogResult = DialogResult.OK;
            this.Hide();
        }

        private void UpdateGridHighlight()
        {
            for (int row = 0; row < MaxRows; row++)
            {
                for (int col = 0; col < MaxCols; col++)
                {
                    var btn = gridButtons[row, col];
                    if (row < hoverRows && col < hoverCols)
                    {
                        btn.BackColor = Color.FromArgb(68, 114, 196);
                    }
                    else
                    {
                        btn.BackColor = Color.White;
                    }
                }
            }
        }

        private void UpdatePreviewText()
        {
            lblPreview.Text = $"{hoverCols} x {hoverRows} Table";
        }

        private void BtnInsertTable_Click(object sender, EventArgs e)
        {
            ActionType = "InsertTable";
            this.DialogResult = DialogResult.OK;
            this.Hide();
        }

        private void BtnExcelSpreadsheet_Click(object sender, EventArgs e)
        {
            ActionType = "ExcelSpreadsheet";
            this.DialogResult = DialogResult.OK;
            this.Hide();
        }

        private Image CreateTableIcon()
        {
            var bitmap = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Transparent);
                using (var pen = new Pen(Color.Black, 1))
                {
                    // Draw simple table icon
                    g.DrawRectangle(pen, 2, 2, 12, 10);
                    g.DrawLine(pen, 2, 6, 14, 6);
                    g.DrawLine(pen, 6, 2, 6, 12);
                    g.DrawLine(pen, 10, 2, 10, 12);
                }
            }
            return bitmap;
        }

        private Image CreateExcelIcon()
        {
            var bitmap = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Transparent);
                using (var brush = new SolidBrush(Color.Green))
                {
                    g.FillRectangle(brush, 2, 2, 12, 12);
                }
                using (var pen = new Pen(Color.White, 1))
                {
                    g.DrawString("X", new Font("Arial", 8, FontStyle.Bold), Brushes.White, 4, 2);
                }
            }
            return bitmap;
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            // Reset grid when mouse leaves the form
            hoverRows = 0;
            hoverCols = 0;
            UpdateGridHighlight();
            lblPreview.Text = "1 x 1 Table";
        }
    }
} 