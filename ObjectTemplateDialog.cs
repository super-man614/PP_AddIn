using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace my_addin
{
    public partial class ObjectTemplateDialog : Form
    {
        public ObjectTemplateItem SelectedItem { get; private set; }
        
        private FlowLayoutPanel templatePanel;
        private TextBox searchBox;
        private ComboBox categoryCombo;
        private Button okButton;
        private Button cancelButton;
        
        private List<ObjectTemplateItem> allTemplates;
        private List<ObjectTemplateItem> filteredTemplates;

        public ObjectTemplateDialog()
        {
            InitializeComponent();
            LoadTemplates();
            FilterTemplates();
        }

        private void InitializeComponent()
        {
            this.Text = "オブジェクトテンプレート";
            this.Size = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Search box
            var searchLabel = new Label { Text = "検索", Location = new Point(12, 15), Size = new Size(50, 23) };
            searchBox = new TextBox { Location = new Point(65, 12), Size = new Size(600, 23) };
            searchBox.TextChanged += SearchBox_TextChanged;

            // Category dropdown
            var categoryLabel = new Label { Text = "カテゴリ", Location = new Point(680, 15), Size = new Size(60, 23) };
            categoryCombo = new ComboBox 
            { 
                Location = new Point(745, 12), 
                Size = new Size(120, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            categoryCombo.Items.AddRange(new[] { "すべて", "基本図形", "矢印", "フローチャート", "アイコン", "デバイス", "その他" });
            categoryCombo.SelectedIndex = 0;
            categoryCombo.SelectedIndexChanged += CategoryCombo_SelectedIndexChanged;

            // Template panel with scroll
            var templateContainer = new Panel 
            { 
                Location = new Point(12, 45), 
                Size = new Size(860, 480),
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle
            };
            
            templatePanel = new FlowLayoutPanel
            {
                Location = new Point(0, 0),
                Size = new Size(840, 460),
                AutoScroll = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(5)
            };
            templateContainer.Controls.Add(templatePanel);

            // Buttons
            okButton = new Button 
            { 
                Text = "OK", 
                Location = new Point(715, 535), 
                Size = new Size(75, 23),
                DialogResult = DialogResult.OK,
                Enabled = false
            };
            okButton.Click += OkButton_Click;

            cancelButton = new Button 
            { 
                Text = "キャンセル", 
                Location = new Point(796, 535), 
                Size = new Size(75, 23),
                DialogResult = DialogResult.Cancel
            };

            // Add controls to form
            this.Controls.AddRange(new Control[] 
            { 
                searchLabel, searchBox, categoryLabel, categoryCombo,
                templateContainer, okButton, cancelButton 
            });
        }

        private void LoadTemplates()
        {
            allTemplates = new List<ObjectTemplateItem>();

            // Basic shapes - Updated with more appropriate icon mappings
            allTemplates.Add(new ObjectTemplateItem("frame14", "基本図形", @"icons\icon-images\257_12.png", CreateFrame14Template));
            allTemplates.Add(new ObjectTemplateItem("heading_textbox_column", "基本図形", @"icons\icon-images\257_15.png", CreateHeadingTextboxColumn));
            allTemplates.Add(new ObjectTemplateItem("column_textbox_long", "基本図形", @"icons\icon-images\257_16.png", CreateColumnTextboxLong));
            allTemplates.Add(new ObjectTemplateItem("row_textbox_long", "基本図形", @"icons\icon-images\257_19.png", CreateRowTextboxLong));
            allTemplates.Add(new ObjectTemplateItem("column_textbox_short", "基本図形", @"icons\icon-images\257_20.png", CreateColumnTextboxShort));
            allTemplates.Add(new ObjectTemplateItem("row_textbox_short", "基本図形", @"icons\icon-images\257_23.png", CreateRowTextboxShort));
            allTemplates.Add(new ObjectTemplateItem("matrix_textbox", "基本図形", @"icons\icon-images\257_26.png", CreateMatrixTextbox));
            allTemplates.Add(new ObjectTemplateItem("round_text", "基本図形", @"icons\icon-images\257_29.png", CreateRoundText));

            // Process flow - Using flowchart-like icons
            allTemplates.Add(new ObjectTemplateItem("box_3sides", "フローチャート", @"icons\icon-images\257_32.png", CreateBox3Sides));
            allTemplates.Add(new ObjectTemplateItem("horizontal_process_sequence", "フローチャート", @"icons\icon-images\257_35.png", CreateHorizontalProcessSequence));
            allTemplates.Add(new ObjectTemplateItem("3row_3textbox", "フローチャート", @"icons\icon-images\257_36.png", Create3Row3Textbox));
            allTemplates.Add(new ObjectTemplateItem("vertical_process_sequence", "フローチャート", @"icons\icon-images\257_37.png", CreateVerticalProcessSequence));

            // Devices - Using device-related icons
            allTemplates.Add(new ObjectTemplateItem("iPhone_status_bar", "デバイス", @"icons\icon-images\258_10.png", CreateIPhoneStatusBar));
            allTemplates.Add(new ObjectTemplateItem("iPhone", "デバイス", @"icons\icon-images\258_16.png", CreateIPhone));
            allTemplates.Add(new ObjectTemplateItem("iPad", "デバイス", @"icons\icon-images\258_22.png", CreateIPad));
            allTemplates.Add(new ObjectTemplateItem("MacBook_Pro", "デバイス", @"icons\icon-images\258_30.png", CreateMacBookPro));

            // Icons and others - Using various icon designs
            allTemplates.Add(new ObjectTemplateItem("gear", "アイコン", @"icons\icon-images\258_33.png", CreateGear));
            allTemplates.Add(new ObjectTemplateItem("evolution", "アイコン", @"icons\icon-images\258_36.png", CreateEvolution));
            
            // Add more templates with existing icons
            allTemplates.Add(new ObjectTemplateItem("textbox_grid", "基本図形", @"icons\icon-images\258_40.png", CreateTextboxGrid));
            allTemplates.Add(new ObjectTemplateItem("process_chain", "フローチャート", @"icons\icon-images\258_47.png", CreateProcessChain));
            allTemplates.Add(new ObjectTemplateItem("data_flow", "フローチャート", @"icons\icon-images\258_54.png", CreateDataFlow));
            allTemplates.Add(new ObjectTemplateItem("timeline", "フローチャート", @"icons\icon-images\258_61.png", CreateTimeline));
        }

        private void FilterTemplates()
        {
            string searchText = searchBox?.Text?.ToLower() ?? "";
            string selectedCategory = categoryCombo?.SelectedItem?.ToString() ?? "すべて";

            filteredTemplates = allTemplates.Where(t => 
                (selectedCategory == "すべて" || t.Category == selectedCategory) &&
                (string.IsNullOrEmpty(searchText) || 
                 t.Name.ToLower().Contains(searchText) || 
                 t.Category.ToLower().Contains(searchText))
            ).ToList();

            UpdateTemplatePanel();
        }

        private void UpdateTemplatePanel()
        {
            templatePanel.Controls.Clear();
            
            foreach (var template in filteredTemplates)
            {
                var templateControl = CreateTemplateControl(template);
                templatePanel.Controls.Add(templateControl);
            }
        }

        private Control CreateTemplateControl(ObjectTemplateItem template)
        {
            var panel = new Panel
            {
                Size = new Size(160, 140),
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.Hand,
                Tag = template,
                Margin = new Padding(5)
            };

            // Icon
            var pictureBox = new PictureBox
            {
                Size = new Size(120, 80),
                Location = new Point(20, 10),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.None
            };

            try
            {
                // Prioritize loading from the icons/icon-images folder
                string[] possiblePaths = {
                    Path.Combine(Application.StartupPath, template.IconPath),
                    Path.Combine(Directory.GetCurrentDirectory(), template.IconPath),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, template.IconPath),
                    // Also try with forward slashes for cross-platform compatibility
                    Path.Combine(Application.StartupPath, template.IconPath.Replace('\\', Path.DirectorySeparatorChar)),
                    Path.Combine(Directory.GetCurrentDirectory(), template.IconPath.Replace('\\', Path.DirectorySeparatorChar)),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, template.IconPath.Replace('\\', Path.DirectorySeparatorChar)),
                    template.IconPath // Try as absolute path
                };

                bool imageLoaded = false;
                foreach (string iconPath in possiblePaths)
                {
                    if (File.Exists(iconPath))
                    {
                        try
                        {
                            pictureBox.Image = Image.FromFile(iconPath);
                            imageLoaded = true;
                            System.Diagnostics.Debug.WriteLine($"Successfully loaded image from: {iconPath}");
                            break;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to load image from {iconPath}: {ex.Message}");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"Icon not found at: {iconPath}");
                    }
                }

                if (!imageLoaded)
                {
                    // Create a representative thumbnail as fallback
                    pictureBox.Image = CreateTemplateThumbnail(template.Name);
                    System.Diagnostics.Debug.WriteLine($"Using generated thumbnail for: {template.Name}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading template icon: {ex.Message}");
                pictureBox.Image = CreateTemplateThumbnail(template.Name);
            }

            // Label
            var label = new Label
            {
                Text = template.Name,
                Location = new Point(5, 95),
                Size = new Size(150, 40),
                TextAlign = ContentAlignment.TopCenter,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 8)
            };

            panel.Controls.Add(pictureBox);
            panel.Controls.Add(label);

            // Click events
            panel.Click += (s, e) => SelectTemplate(template, panel);
            pictureBox.Click += (s, e) => SelectTemplate(template, panel);
            label.Click += (s, e) => SelectTemplate(template, panel);

            return panel;
        }

        private void SelectTemplate(ObjectTemplateItem template, Panel panel)
        {
            // Clear previous selection
            foreach (Control control in templatePanel.Controls)
            {
                if (control is Panel p)
                    p.BackColor = SystemColors.Control;
            }

            // Highlight selected
            panel.BackColor = Color.LightBlue;
            SelectedItem = template;
            okButton.Enabled = true;
        }

        private void SearchBox_TextChanged(object sender, EventArgs e)
        {
            FilterTemplates();
        }

        private void CategoryCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            FilterTemplates();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private Bitmap CreateTemplateThumbnail(string templateName)
        {
            var bitmap = new Bitmap(120, 80);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                // Create different visual representations based on template name
                switch (templateName.ToLower())
                {
                    case "frame14":
                    case "heading_textbox_column":
                        // Blue header rectangle
                        g.FillRectangle(new SolidBrush(Color.FromArgb(70, 130, 180)), 10, 10, 100, 25);
                        g.FillRectangle(new SolidBrush(Color.LightGray), 10, 35, 100, 35);
                        g.DrawString("XXX", new Font("Arial", 8, FontStyle.Bold), Brushes.White, 12, 15);
                        break;

                    case "matrix_textbox":
                        // 3x3 grid
                        for (int row = 0; row < 3; row++)
                        {
                            for (int col = 0; col < 3; col++)
                            {
                                var rect = new Rectangle(15 + col * 30, 15 + row * 20, 25, 15);
                                g.FillRectangle(Brushes.LightBlue, rect);
                                g.DrawRectangle(Pens.Black, rect);
                                g.DrawString("X", SystemFonts.DefaultFont, Brushes.Black, rect.X + 8, rect.Y + 2);
                            }
                        }
                        break;

                    case "round_text":
                        // Oval shape
                        g.FillEllipse(new SolidBrush(Color.LightBlue), 20, 20, 80, 40);
                        g.DrawEllipse(Pens.Blue, 20, 20, 80, 40);
                        g.DrawString("XXX", SystemFonts.DefaultFont, Brushes.Black, 50, 35);
                        break;

                    case "horizontal_process_sequence":
                        // Horizontal flow
                        for (int i = 0; i < 3; i++)
                        {
                            var rect = new Rectangle(10 + i * 35, 25, 30, 20);
                            g.FillRectangle(new SolidBrush(Color.FromArgb(70, 130, 180)), rect);
                            g.DrawString("X", new Font("Arial", 8), Brushes.White, rect.X + 12, rect.Y + 5);
                            if (i < 2)
                            {
                                g.DrawLine(new Pen(Color.Black, 2), rect.Right + 2, rect.Y + 10, rect.Right + 8, rect.Y + 10);
                                g.DrawLine(new Pen(Color.Black, 2), rect.Right + 5, rect.Y + 7, rect.Right + 8, rect.Y + 10);
                                g.DrawLine(new Pen(Color.Black, 2), rect.Right + 5, rect.Y + 13, rect.Right + 8, rect.Y + 10);
                            }
                        }
                        break;

                    case "vertical_process_sequence":
                        // Vertical flow
                        for (int i = 0; i < 3; i++)
                        {
                            var rect = new Rectangle(35, 10 + i * 25, 50, 20);
                            g.FillRectangle(new SolidBrush(Color.FromArgb(70, 130, 180)), rect);
                            g.DrawString("XXX", new Font("Arial", 8), Brushes.White, rect.X + 15, rect.Y + 5);
                        }
                        break;

                    case "iphone":
                    case "iphone_status_bar":
                        // Phone shape
                        FillRoundedRectangle(g, Brushes.Black, 35, 10, 50, 60, 8);
                        FillRoundedRectangle(g, Brushes.White, 40, 15, 40, 50, 3);
                        break;

                    case "ipad":
                        // Tablet shape
                        FillRoundedRectangle(g, Brushes.Black, 25, 20, 70, 50, 5);
                        FillRoundedRectangle(g, Brushes.White, 30, 25, 60, 40, 3);
                        break;

                    case "macbook_pro":
                        // Laptop shape
                        g.FillRectangle(Brushes.Silver, 20, 30, 80, 40);
                        g.DrawRectangle(Pens.Black, 20, 30, 80, 40);
                        g.FillRectangle(Brushes.Black, 25, 35, 70, 30);
                        break;

                    case "gear":
                        // Gear shape
                        g.FillEllipse(Brushes.Gray, 40, 25, 40, 30);
                        g.FillEllipse(Brushes.White, 50, 32, 20, 16);
                        for (int i = 0; i < 8; i++)
                        {
                            double angle = i * Math.PI / 4;
                            int x = (int)(60 + 25 * Math.Cos(angle));
                            int y = (int)(40 + 18 * Math.Sin(angle));
                            g.FillRectangle(Brushes.Gray, x - 2, y - 2, 4, 4);
                        }
                        break;

                    default:
                        // Generic template
                        g.FillRectangle(Brushes.LightGray, 10, 10, 100, 60);
                        g.DrawRectangle(Pens.DarkGray, 10, 10, 100, 60);
                        g.DrawString("Template", new Font("Arial", 8), Brushes.Black, 30, 35);
                        break;
                }

                // Add border
                g.DrawRectangle(Pens.LightGray, 0, 0, 119, 79);
            }
            return bitmap;
        }

        // Extension method for rounded rectangles
        private void FillRoundedRectangle(Graphics g, Brush brush, int x, int y, int width, int height, int radius)
        {
            var path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddArc(x, y, radius * 2, radius * 2, 180, 90);
            path.AddArc(x + width - radius * 2, y, radius * 2, radius * 2, 270, 90);
            path.AddArc(x + width - radius * 2, y + height - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(x, y + height - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseAllFigures();
            g.FillPath(brush, path);
        }

        #region Template Creation Methods

        private void CreateFrame14Template(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, 200, 150);
            shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightBlue);
            shape.TextFrame.TextRange.Text = "Frame14";
            shape.Name = "frame14_template";
        }

        private void CreateHeadingTextboxColumn(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 300, 60);
            shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            shape.TextFrame.TextRange.Text = "見出し";
            shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
            shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            shape.Name = "heading_textbox_column";
        }

        private void CreateColumnTextboxLong(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, top, 200, 100);
            shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
            shape.TextFrame.TextRange.Text = "XXX";
            shape.Name = "column_textbox_long";
        }

        private void CreateRowTextboxLong(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var rect = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 60, 150);
            rect.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            rect.TextFrame.TextRange.Text = "XXX";
            rect.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
            
            var textBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left + 70, top, 200, 150);
            textBox.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
            textBox.TextFrame.TextRange.Text = "説明テキスト";
            
            var group = slide.Shapes.Range(new object[] { rect.Name, textBox.Name }).Group();
            group.Name = "row_textbox_long";
        }

        private void CreateColumnTextboxShort(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 150, 100);
            shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            shape.TextFrame.TextRange.Text = "XXX";
            shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
            shape.Name = "column_textbox_short";
        }

        private void CreateRowTextboxShort(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 150, 80);
            shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
            shape.TextFrame.TextRange.Text = "XXX";
            shape.Name = "row_textbox_short";
        }

        private void CreateMatrixTextbox(PowerPoint.Slide slide)
        {
            float left = 100, top = 100, cellSize = 40;
            var shapes = new List<string>();
            
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    var cell = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                        left + col * cellSize, top + row * cellSize, cellSize, cellSize);
                    cell.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
                    cell.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    cell.TextFrame.TextRange.Text = "XXX";
                    shapes.Add(cell.Name);
                }
            }
            
            var group = slide.Shapes.Range(shapes.ToArray()).Group();
            group.Name = "matrix_textbox";
        }

        private void CreateRoundText(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, top, 120, 60);
            shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightBlue);
            shape.TextFrame.TextRange.Text = "XXX";
            shape.Name = "round_text";
        }

        private void CreateBox3Sides(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 200, 80);
            shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
            shape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
            shape.TextFrame.TextRange.Text = "片側線なし\nボックス";
            shape.Name = "box_3sides";
        }

        private void CreateHorizontalProcessSequence(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shapes = new List<string>();
            
            for (int i = 0; i < 4; i++)
            {
                var box = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                    left + i * 80, top, 70, 40);
                box.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
                box.TextFrame.TextRange.Text = "XXX";
                box.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
                shapes.Add(box.Name);
                
                if (i < 3)
                {
                    var arrow = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightArrow, 
                        left + i * 80 + 70, top + 15, 10, 10);
                    arrow.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    shapes.Add(arrow.Name);
                }
            }
            
            var group = slide.Shapes.Range(shapes.ToArray()).Group();
            group.Name = "horizontal_process_sequence";
        }

        private void Create3Row3Textbox(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shapes = new List<string>();
            
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    var box = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                        left + col * 80, top + row * 50, 70, 40);
                    box.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
                    box.TextFrame.TextRange.Text = "XXX";
                    box.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
                    shapes.Add(box.Name);
                }
            }
            
            var group = slide.Shapes.Range(shapes.ToArray()).Group();
            group.Name = "3row_3textbox";
        }

        private void CreateVerticalProcessSequence(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shapes = new List<string>();
            
            for (int i = 0; i < 4; i++)
            {
                var box = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                    left, top + i * 60, 120, 50);
                box.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
                box.TextFrame.TextRange.Text = "XXX";
                box.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
                shapes.Add(box.Name);
            }
            
            var group = slide.Shapes.Range(shapes.ToArray()).Group();
            group.Name = "vertical_process_sequence";
        }

        private void CreateIPhoneStatusBar(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var phone = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, 80, 160);
            phone.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
            phone.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
            phone.Name = "iPhone_status_bar";
        }

        private void CreateIPhone(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var phone = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, 80, 160);
            phone.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
            phone.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
            phone.Line.Weight = 2;
            phone.Name = "iPhone";
        }

        private void CreateIPad(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var tablet = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, 120, 90);
            tablet.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
            tablet.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
            tablet.Line.Weight = 2;
            tablet.Name = "iPad";
        }

        private void CreateMacBookPro(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var laptop = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 150, 100);
            laptop.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Silver);
            laptop.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
            laptop.Line.Weight = 2;
            laptop.Name = "MacBook_Pro";
        }

        private void CreateGear(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var gear = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShape32pointStar, left, top, 80, 80);
            gear.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Gray);
            gear.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
            gear.Name = "gear";
        }

        private void CreateEvolution(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shapes = new List<string>();
            
            for (int i = 0; i < 3; i++)
            {
                var circle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, 
                    left + i * 60, top, 40, 40);
                circle.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightBlue);
                circle.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                shapes.Add(circle.Name);
                
                if (i < 2)
                {
                    var arrow = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightArrow, 
                        left + i * 60 + 40, top + 15, 20, 10);
                    arrow.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    shapes.Add(arrow.Name);
                }
            }
            
            var group = slide.Shapes.Range(shapes.ToArray()).Group();
            group.Name = "evolution";
        }

        private void CreateTextboxGrid(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shapes = new List<string>();
            
            for (int row = 0; row < 2; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    var box = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                        left + col * 60, top + row * 40, 50, 35);
                    box.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
                    box.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    box.TextFrame.TextRange.Text = "XXX";
                    shapes.Add(box.Name);
                }
            }
            
            var group = slide.Shapes.Range(shapes.ToArray()).Group();
            group.Name = "textbox_grid";
        }

        private void CreateProcessChain(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shapes = new List<string>();
            
            // Create connected process boxes
            for (int i = 0; i < 4; i++)
            {
                var box = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                    left + i * 80, top, 70, 50);
                box.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(100, 150, 200));
                box.TextFrame.TextRange.Text = $"Step {i + 1}";
                box.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
                shapes.Add(box.Name);
                
                if (i < 3)
                {
                    var connector = slide.Shapes.AddConnector(Office.MsoConnectorType.msoConnectorStraight,
                        left + (i + 1) * 80 - 10, top + 25, left + (i + 1) * 80, top + 25);
                    connector.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    connector.Line.Weight = 2;
                    shapes.Add(connector.Name);
                }
            }
            
            var group = slide.Shapes.Range(shapes.ToArray()).Group();
            group.Name = "process_chain";
        }

        private void CreateDataFlow(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shapes = new List<string>();
            
            // Input
            var input = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeParallelogram, left, top, 80, 40);
            input.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightBlue);
            input.TextFrame.TextRange.Text = "Input";
            shapes.Add(input.Name);
            
            // Process
            var process = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left + 100, top, 80, 40);
            process.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            process.TextFrame.TextRange.Text = "Process";
            process.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
            shapes.Add(process.Name);
            
            // Output
            var output = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeParallelogram, left + 200, top, 80, 40);
            output.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightGreen);
            output.TextFrame.TextRange.Text = "Output";
            shapes.Add(output.Name);
            
            var group = slide.Shapes.Range(shapes.ToArray()).Group();
            group.Name = "data_flow";
        }

        private void CreateTimeline(PowerPoint.Slide slide)
        {
            float left = 100, top = 100;
            var shapes = new List<string>();
            
            // Timeline base line
            var baseLine = slide.Shapes.AddLine(left, top + 50, left + 300, top + 50);
            baseLine.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
            baseLine.Line.Weight = 3;
            shapes.Add(baseLine.Name);
            
            // Timeline markers
            for (int i = 0; i < 4; i++)
            {
                float x = left + i * 75;
                
                // Marker line
                var marker = slide.Shapes.AddLine(x, top + 40, x, top + 60);
                marker.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Black);
                marker.Line.Weight = 2;
                shapes.Add(marker.Name);
                
                // Event box
                var eventBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, x - 25, top + 10, 50, 25);
                eventBox.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightYellow);
                eventBox.TextFrame.TextRange.Text = $"Event {i + 1}";
                eventBox.TextFrame.TextRange.Font.Size = 8;
                shapes.Add(eventBox.Name);
            }
            
            var group = slide.Shapes.Range(shapes.ToArray()).Group();
            group.Name = "timeline";
        }

        #endregion
    }

    public class ObjectTemplateItem
    {
        public string Name { get; set; }
        public string Category { get; set; }
        public string IconPath { get; set; }
        public Action<PowerPoint.Slide> InsertAction { get; set; }

        public ObjectTemplateItem(string name, string category, string iconPath, Action<PowerPoint.Slide> insertAction)
        {
            Name = name;
            Category = category;
            IconPath = iconPath;
            InsertAction = insertAction;
        }
    }
}
