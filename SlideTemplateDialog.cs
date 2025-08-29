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
    public partial class SlideTemplateDialog : Form
    {
        public SlideTemplateItem SelectedItem { get; private set; }
        
        private FlowLayoutPanel templatePanel;
        private TextBox searchBox;
        private ComboBox categoryCombo;
        private Button okButton;
        private Button cancelButton;
        
        private List<SlideTemplateItem> allTemplates;
        private List<SlideTemplateItem> filteredTemplates;

        public SlideTemplateDialog()
        {
            InitializeComponent();
            LoadTemplates();
            FilterTemplates();
        }

        private void InitializeComponent()
        {
            this.Text = "スライドテンプレート";
            this.Size = new Size(950, 650);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Search box
            var searchLabel = new Label { Text = "検索", Location = new Point(12, 15), Size = new Size(50, 23) };
            searchBox = new TextBox { Location = new Point(65, 12), Size = new Size(650, 23) };
            searchBox.TextChanged += SearchBox_TextChanged;

            // Category dropdown
            var categoryLabel = new Label { Text = "カテゴリ", Location = new Point(730, 15), Size = new Size(60, 23) };
            categoryCombo = new ComboBox 
            { 
                Location = new Point(795, 12), 
                Size = new Size(130, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            categoryCombo.Items.AddRange(new[] { "すべて", "プレゼンテーション", "レポート", "マーケティング", "教育", "ビジネス", "その他" });
            categoryCombo.SelectedIndex = 0;
            categoryCombo.SelectedIndexChanged += CategoryCombo_SelectedIndexChanged;

            // Template panel with scroll
            var templateContainer = new Panel 
            { 
                Location = new Point(12, 45), 
                Size = new Size(910, 520),
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle
            };
            
            templatePanel = new FlowLayoutPanel
            {
                Location = new Point(0, 0),
                Size = new Size(890, 500),
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
                Location = new Point(765, 580), 
                Size = new Size(75, 23),
                DialogResult = DialogResult.OK,
                Enabled = false
            };
            okButton.Click += OkButton_Click;

            cancelButton = new Button 
            { 
                Text = "キャンセル", 
                Location = new Point(846, 580), 
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
            allTemplates = new List<SlideTemplateItem>();

            // Presentation templates - using actual image files from icons/slide-images
            allTemplates.Add(new SlideTemplateItem("title_slide", "プレゼンテーション", @"icons\slide-images\256.png", CreateTitleSlide));
            allTemplates.Add(new SlideTemplateItem("agenda_slide", "プレゼンテーション", @"icons\slide-images\257.png", CreateAgendaSlide));
            allTemplates.Add(new SlideTemplateItem("section_header", "プレゼンテーション", @"icons\slide-images\258.png", CreateSectionHeaderSlide));
            allTemplates.Add(new SlideTemplateItem("two_column_layout", "プレゼンテーション", @"icons\slide-images\259.png", CreateTwoColumnSlide));
            allTemplates.Add(new SlideTemplateItem("three_column_layout", "プレゼンテーション", @"icons\slide-images\260.png", CreateThreeColumnSlide));
            allTemplates.Add(new SlideTemplateItem("bullet_points", "プレゼンテーション", @"icons\slide-images\262.png", CreateBulletPointsSlide));
            allTemplates.Add(new SlideTemplateItem("image_with_caption", "プレゼンテーション", @"icons\slide-images\263.png", CreateImageCaptionSlide));
            allTemplates.Add(new SlideTemplateItem("thank_you_slide", "プレゼンテーション", @"icons\slide-images\264.png", CreateThankYouSlide));

            // Business templates
            allTemplates.Add(new SlideTemplateItem("comparison_table", "ビジネス", @"icons\slide-images\265.png", CreateComparisonTableSlide));
            allTemplates.Add(new SlideTemplateItem("timeline_slide", "ビジネス", @"icons\slide-images\266.png", CreateTimelineSlide));
            allTemplates.Add(new SlideTemplateItem("process_flow", "ビジネス", @"icons\slide-images\268.png", CreateProcessFlowSlide));
            allTemplates.Add(new SlideTemplateItem("chart_placeholder", "ビジネス", @"icons\slide-images\271.png", CreateChartPlaceholderSlide));
            allTemplates.Add(new SlideTemplateItem("team_introduction", "ビジネス", @"icons\slide-images\276.png", CreateTeamIntroSlide));
            allTemplates.Add(new SlideTemplateItem("contact_info", "ビジネス", @"icons\slide-images\277.png", CreateContactInfoSlide));

            // Marketing templates
            allTemplates.Add(new SlideTemplateItem("product_showcase", "マーケティング", @"icons\slide-images\278.png", CreateProductShowcaseSlide));
            allTemplates.Add(new SlideTemplateItem("features_benefits", "マーケティング", @"icons\slide-images\279.png", CreateFeaturesBenefitsSlide));
            allTemplates.Add(new SlideTemplateItem("testimonial_slide", "マーケティング", @"icons\slide-images\280.png", CreateTestimonialSlide));
            allTemplates.Add(new SlideTemplateItem("pricing_table", "マーケティング", @"icons\slide-images\281.png", CreatePricingTableSlide));

            // Report templates
            allTemplates.Add(new SlideTemplateItem("executive_summary", "レポート", @"icons\slide-images\283.png", CreateExecutiveSummarySlide));
            allTemplates.Add(new SlideTemplateItem("data_visualization", "レポート", @"icons\slide-images\284.png", CreateDataVisualizationSlide));
            allTemplates.Add(new SlideTemplateItem("key_insights", "レポート", @"icons\slide-images\285.png", CreateKeyInsightsSlide));
            allTemplates.Add(new SlideTemplateItem("recommendations", "レポート", @"icons\slide-images\286.png", CreateRecommendationsSlide));

            // Education templates
            allTemplates.Add(new SlideTemplateItem("lesson_title", "教育", @"icons\slide-images\290.png", CreateLessonTitleSlide));
            allTemplates.Add(new SlideTemplateItem("learning_objectives", "教育", @"icons\slide-images\291.png", CreateLearningObjectivesSlide));
            allTemplates.Add(new SlideTemplateItem("quiz_question", "教育", @"icons\slide-images\293.png", CreateQuizQuestionSlide));
            allTemplates.Add(new SlideTemplateItem("summary_review", "教育", @"icons\slide-images\295.png", CreateSummaryReviewSlide));
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

        private Control CreateTemplateControl(SlideTemplateItem template)
        {
            var panel = new Panel
            {
                Size = new Size(180, 160),
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.Hand,
                Tag = template,
                Margin = new Padding(5)
            };

            // Icon
            var pictureBox = new PictureBox
            {
                Size = new Size(140, 100),
                Location = new Point(20, 10),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.None
            };

            try
            {
                // Try to load from multiple possible paths
                string[] possiblePaths = {
                    Path.Combine(Application.StartupPath, template.IconPath),
                    Path.Combine(Directory.GetCurrentDirectory(), template.IconPath),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, template.IconPath),
                    Path.Combine(Application.StartupPath, template.IconPath.Replace('\\', Path.DirectorySeparatorChar)),
                    Path.Combine(Directory.GetCurrentDirectory(), template.IconPath.Replace('\\', Path.DirectorySeparatorChar)),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, template.IconPath.Replace('\\', Path.DirectorySeparatorChar)),
                    template.IconPath
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
                            System.Diagnostics.Debug.WriteLine($"Successfully loaded slide image from: {iconPath}");
                            break;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to load slide image from {iconPath}: {ex.Message}");
                        }
                    }
                }

                if (!imageLoaded)
                {
                    // Create a representative thumbnail as fallback
                    pictureBox.Image = CreateSlideThumbnail(template.Name);
                    System.Diagnostics.Debug.WriteLine($"Using generated slide thumbnail for: {template.Name}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading slide template icon: {ex.Message}");
                pictureBox.Image = CreateSlideThumbnail(template.Name);
            }

            // Label
            var label = new Label
            {
                Text = template.DisplayName,
                Location = new Point(5, 115),
                Size = new Size(170, 40),
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

        private void SelectTemplate(SlideTemplateItem template, Panel panel)
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

        private Bitmap CreateSlideThumbnail(string templateName)
        {
            var bitmap = new Bitmap(140, 100);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                // Create different visual representations based on template name
                switch (templateName.ToLower())
                {
                    case "title_slide":
                        // Title slide layout
                        g.FillRectangle(new SolidBrush(Color.FromArgb(70, 130, 180)), 10, 20, 120, 20);
                        g.FillRectangle(new SolidBrush(Color.LightGray), 20, 50, 100, 10);
                        g.DrawString("TITLE", new Font("Arial", 8, FontStyle.Bold), Brushes.White, 15, 25);
                        break;

                    case "two_column_layout":
                        // Two column layout
                        g.FillRectangle(Brushes.LightBlue, 10, 10, 60, 15);
                        g.FillRectangle(Brushes.LightGray, 10, 30, 60, 60);
                        g.FillRectangle(Brushes.LightGray, 75, 30, 60, 60);
                        g.DrawLine(Pens.Gray, 72, 30, 72, 90);
                        break;

                    case "three_column_layout":
                        // Three column layout
                        g.FillRectangle(Brushes.LightBlue, 10, 10, 120, 15);
                        for (int i = 0; i < 3; i++)
                        {
                            g.FillRectangle(Brushes.LightGray, 10 + i * 40, 30, 35, 60);
                        }
                        break;

                    case "bullet_points":
                        // Bullet points
                        g.FillRectangle(Brushes.LightBlue, 10, 10, 120, 15);
                        for (int i = 0; i < 4; i++)
                        {
                            g.FillEllipse(Brushes.Black, 15, 35 + i * 15, 3, 3);
                            g.FillRectangle(Brushes.LightGray, 25, 33 + i * 15, 100, 8);
                        }
                        break;

                    case "comparison_table":
                        // Table layout
                        g.FillRectangle(Brushes.LightBlue, 10, 10, 120, 15);
                        for (int row = 0; row < 3; row++)
                        {
                            for (int col = 0; col < 3; col++)
                            {
                                var rect = new Rectangle(15 + col * 35, 30 + row * 20, 30, 15);
                                g.FillRectangle(row == 0 ? Brushes.LightBlue : Brushes.LightGray, rect);
                                g.DrawRectangle(Pens.Gray, rect);
                            }
                        }
                        break;

                    case "timeline_slide":
                        // Timeline
                        g.FillRectangle(Brushes.LightBlue, 10, 10, 120, 15);
                        g.DrawLine(new Pen(Color.Black, 2), 20, 50, 120, 50);
                        for (int i = 0; i < 4; i++)
                        {
                            int x = 30 + i * 25;
                            g.FillEllipse(Brushes.Blue, x - 3, 47, 6, 6);
                            g.FillRectangle(Brushes.LightGray, x - 10, 60, 20, 8);
                        }
                        break;

                    case "chart_placeholder":
                        // Chart placeholder
                        g.FillRectangle(Brushes.LightBlue, 10, 10, 120, 15);
                        g.DrawRectangle(Pens.Gray, 20, 30, 100, 50);
                        g.DrawString("Chart", SystemFonts.DefaultFont, Brushes.Gray, 60, 50);
                        break;

                    default:
                        // Generic slide
                        g.FillRectangle(Brushes.LightBlue, 10, 10, 120, 15);
                        g.FillRectangle(Brushes.LightGray, 20, 35, 100, 50);
                        g.DrawString("Slide", new Font("Arial", 8), Brushes.Black, 60, 55);
                        break;
                }

                // Add slide border
                g.DrawRectangle(Pens.DarkGray, 0, 0, 139, 99);
            }
            return bitmap;
        }

        #region Slide Template Creation Methods

        private void CreateTitleSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Title
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 150, 600, 100);
            title.TextFrame.TextRange.Text = "プレゼンテーションタイトル";
            title.TextFrame.TextRange.Font.Size = 36;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            title.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            
            // Subtitle
            var subtitle = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 150, 280, 500, 50);
            subtitle.TextFrame.TextRange.Text = "サブタイトル・発表者名・日付";
            subtitle.TextFrame.TextRange.Font.Size = 18;
            subtitle.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            
            slide.Name = "Title Slide";
        }

        private void CreateAgendaSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Title
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 600, 60);
            title.TextFrame.TextRange.Text = "アジェンダ";
            title.TextFrame.TextRange.Font.Size = 28;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            
            // Agenda items
            var agenda = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 80, 130, 550, 350);
            agenda.TextFrame.TextRange.Text = "1. イントロダクション\n\n2. 現状分析\n\n3. 提案内容\n\n4. 実施計画\n\n5. まとめ・質疑応答";
            agenda.TextFrame.TextRange.Font.Size = 20;
            agenda.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 12;
            
            slide.Name = "Agenda Slide";
        }

        private void CreateSectionHeaderSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Background shape
            var bg = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, 720, 540);
            bg.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            bg.Line.Visible = Office.MsoTriState.msoFalse;
            
            // Section title
            var sectionTitle = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 200, 520, 140);
            sectionTitle.TextFrame.TextRange.Text = "セクション タイトル";
            sectionTitle.TextFrame.TextRange.Font.Size = 48;
            sectionTitle.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            sectionTitle.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
            sectionTitle.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            sectionTitle.Fill.Visible = Office.MsoTriState.msoFalse;
            sectionTitle.Line.Visible = Office.MsoTriState.msoFalse;
            
            slide.Name = "Section Header";
        }

        private void CreateTwoColumnSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Title
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 620, 60);
            title.TextFrame.TextRange.Text = "2カラムレイアウト";
            title.TextFrame.TextRange.Font.Size = 28;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            
            // Left column
            var leftColumn = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 130, 300, 350);
            leftColumn.TextFrame.TextRange.Text = "左カラム\n\n• ポイント1\n• ポイント2\n• ポイント3\n\nここに詳細な説明を記載します。";
            leftColumn.TextFrame.TextRange.Font.Size = 16;
            leftColumn.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(240, 248, 255));
            leftColumn.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
            
            // Right column
            var rightColumn = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 370, 130, 300, 350);
            rightColumn.TextFrame.TextRange.Text = "右カラム\n\n• ポイント1\n• ポイント2\n• ポイント3\n\nここに詳細な説明を記載します。";
            rightColumn.TextFrame.TextRange.Font.Size = 16;
            rightColumn.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(240, 248, 255));
            rightColumn.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
            
            slide.Name = "Two Column Layout";
        }

        private void CreateThreeColumnSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Title
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 620, 60);
            title.TextFrame.TextRange.Text = "3カラムレイアウト";
            title.TextFrame.TextRange.Font.Size = 28;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            
            // Three columns
            for (int i = 0; i < 3; i++)
            {
                var column = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50 + i * 200, 130, 180, 350);
                column.TextFrame.TextRange.Text = $"カラム {i + 1}\n\n• ポイント1\n• ポイント2\n• ポイント3";
                column.TextFrame.TextRange.Font.Size = 14;
                column.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(240, 248, 255));
                column.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
            }
            
            slide.Name = "Three Column Layout";
        }

        private void CreateBulletPointsSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Title
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 620, 60);
            title.TextFrame.TextRange.Text = "主要ポイント";
            title.TextFrame.TextRange.Font.Size = 28;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            
            // Bullet points
            var bullets = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 80, 130, 560, 350);
            bullets.TextFrame.TextRange.Text = "• 重要なポイント1についての説明\n\n• 重要なポイント2についての説明\n\n• 重要なポイント3についての説明\n\n• 重要なポイント4についての説明";
            bullets.TextFrame.TextRange.Font.Size = 18;
            bullets.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 12;
            
            slide.Name = "Bullet Points";
        }

        private void CreateImageCaptionSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Title
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 620, 60);
            title.TextFrame.TextRange.Text = "画像とキャプション";
            title.TextFrame.TextRange.Font.Size = 28;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            
            // Image placeholder
            var imagePlaceholder = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 100, 130, 400, 250);
            imagePlaceholder.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
            imagePlaceholder.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Gray);
            
            var imageText = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 250, 240, 100, 30);
            imageText.TextFrame.TextRange.Text = "画像";
            imageText.TextFrame.TextRange.Font.Size = 16;
            imageText.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            imageText.Fill.Visible = Office.MsoTriState.msoFalse;
            imageText.Line.Visible = Office.MsoTriState.msoFalse;
            
            // Caption
            var caption = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 400, 400, 80);
            caption.TextFrame.TextRange.Text = "ここに画像の説明やキャプションを記載します。";
            caption.TextFrame.TextRange.Font.Size = 14;
            caption.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            
            slide.Name = "Image with Caption";
        }

        private void CreateThankYouSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Background
            var bg = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, 720, 540);
            bg.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            bg.Line.Visible = Office.MsoTriState.msoFalse;
            
            // Thank you message
            var thankYou = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 180, 520, 100);
            thankYou.TextFrame.TextRange.Text = "ありがとうございました";
            thankYou.TextFrame.TextRange.Font.Size = 42;
            thankYou.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            thankYou.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
            thankYou.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            thankYou.Fill.Visible = Office.MsoTriState.msoFalse;
            thankYou.Line.Visible = Office.MsoTriState.msoFalse;
            
            // Q&A text
            var qna = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 200, 320, 320, 50);
            qna.TextFrame.TextRange.Text = "質疑応答";
            qna.TextFrame.TextRange.Font.Size = 24;
            qna.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
            qna.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            qna.Fill.Visible = Office.MsoTriState.msoFalse;
            qna.Line.Visible = Office.MsoTriState.msoFalse;
            
            slide.Name = "Thank You";
        }

        // Additional template methods would continue here following the same pattern...
        private void CreateComparisonTableSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Title
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 620, 60);
            title.TextFrame.TextRange.Text = "比較表";
            title.TextFrame.TextRange.Font.Size = 28;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            
            // Table
            var table = slide.Shapes.AddTable(4, 3, 100, 130, 520, 300);
            table.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "項目";
            table.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "オプションA";
            table.Table.Cell(1, 3).Shape.TextFrame.TextRange.Text = "オプションB";
            
            // Style header row
            for (int col = 1; col <= 3; col++)
            {
                table.Table.Cell(1, col).Shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
                table.Table.Cell(1, col).Shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
                table.Table.Cell(1, col).Shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            }
            
            slide.Name = "Comparison Table";
        }

        private void CreateTimelineSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Title
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 620, 60);
            title.TextFrame.TextRange.Text = "タイムライン";
            title.TextFrame.TextRange.Font.Size = 28;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            
            // Timeline base line
            var baseLine = slide.Shapes.AddLine(100, 300, 600, 300);
            baseLine.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            baseLine.Line.Weight = 4;
            
            // Timeline events
            for (int i = 0; i < 4; i++)
            {
                float x = 150 + i * 125;
                
                // Event marker
                var marker = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, x - 8, 292, 16, 16);
                marker.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
                marker.Line.Visible = Office.MsoTriState.msoFalse;
                
                // Event text
                var eventText = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, x - 50, 320, 100, 80);
                eventText.TextFrame.TextRange.Text = $"イベント {i + 1}\n2024年{i + 3}月";
                eventText.TextFrame.TextRange.Font.Size = 12;
                eventText.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            }
            
            slide.Name = "Timeline";
        }

        private void CreateProcessFlowSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Title
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 620, 60);
            title.TextFrame.TextRange.Text = "プロセスフロー";
            title.TextFrame.TextRange.Font.Size = 28;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
            
            // Process boxes
            for (int i = 0; i < 4; i++)
            {
                var box = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 80 + i * 140, 200, 120, 80);
                box.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(70, 130, 180));
                box.TextFrame.TextRange.Text = $"ステップ {i + 1}";
                box.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
                box.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                box.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                
                // Arrow (except for last box)
                if (i < 3)
                {
                    var arrow = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightArrow, 205 + i * 140, 230, 30, 20);
                    arrow.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Gray);
                }
            }
            
            slide.Name = "Process Flow";
        }

        // Placeholder methods for remaining templates
        private void CreateChartPlaceholderSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Chart Placeholder";
        }

        private void CreateTeamIntroSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Team Introduction";
        }

        private void CreateContactInfoSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Contact Info";
        }

        private void CreateProductShowcaseSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Product Showcase";
        }

        private void CreateFeaturesBenefitsSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Features Benefits";
        }

        private void CreateTestimonialSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Testimonial";
        }

        private void CreatePricingTableSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Pricing Table";
        }

        private void CreateExecutiveSummarySlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Executive Summary";
        }

        private void CreateDataVisualizationSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Data Visualization";
        }

        private void CreateKeyInsightsSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Key Insights";
        }

        private void CreateRecommendationsSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Recommendations";
        }

        private void CreateLessonTitleSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Lesson Title";
        }

        private void CreateLearningObjectivesSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Learning Objectives";
        }

        private void CreateQuizQuestionSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Quiz Question";
        }

        private void CreateSummaryReviewSlide(PowerPoint.Presentation presentation)
        {
            var slide = presentation.Slides.Add(presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            // Implementation similar to other methods...
            slide.Name = "Summary Review";
        }

        #endregion
    }

    public class SlideTemplateItem
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Category { get; set; }
        public string IconPath { get; set; }
        public Action<PowerPoint.Presentation> InsertAction { get; set; }

        public SlideTemplateItem(string name, string category, string iconPath, Action<PowerPoint.Presentation> insertAction)
        {
            Name = name;
            DisplayName = GetDisplayName(name);
            Category = category;
            IconPath = iconPath;
            InsertAction = insertAction;
        }

        private string GetDisplayName(string name)
        {
            var displayNames = new Dictionary<string, string>
            {
                {"title_slide", "タイトルスライド"},
                {"agenda_slide", "アジェンダ"},
                {"section_header", "セクションヘッダー"},
                {"two_column_layout", "2カラムレイアウト"},
                {"three_column_layout", "3カラムレイアウト"},
                {"bullet_points", "箇条書き"},
                {"image_with_caption", "画像とキャプション"},
                {"thank_you_slide", "ありがとうスライド"},
                {"comparison_table", "比較表"},
                {"timeline_slide", "タイムライン"},
                {"process_flow", "プロセスフロー"},
                {"chart_placeholder", "グラフプレースホルダー"},
                {"team_introduction", "チーム紹介"},
                {"contact_info", "連絡先情報"},
                {"product_showcase", "商品紹介"},
                {"features_benefits", "機能とメリット"},
                {"testimonial_slide", "お客様の声"},
                {"pricing_table", "価格表"},
                {"executive_summary", "エグゼクティブサマリー"},
                {"data_visualization", "データ可視化"},
                {"key_insights", "重要な洞察"},
                {"recommendations", "推奨事項"},
                {"lesson_title", "レッスンタイトル"},
                {"learning_objectives", "学習目標"},
                {"quiz_question", "クイズ"},
                {"summary_review", "まとめ・復習"}
            };
            
            return displayNames.ContainsKey(name) ? displayNames[name] : name;
        }
    }
}
