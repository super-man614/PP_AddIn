using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace my_addin
{
	public class ObjectTemplatesDialog : Form
	{
		private TextBox txtSearch;
		private ListBox lstCategories;
		private FlowLayoutPanel flowTemplates;
		private Button btnInsert;
		private Button btnCancel;
		private Label lblTitle;
		private Panel separator;

		private readonly List<TemplateItem> allTemplates = new List<TemplateItem>();
		private TemplateItem selectedTemplate;

		public string SelectedTemplateName => selectedTemplate?.Name ?? string.Empty;

		public ObjectTemplatesDialog()
		{
			InitializeComponent();
			LoadTemplates();
			ApplyFilters();
		}

		private void InitializeComponent()
		{
			this.Text = "Object Templates";
			this.StartPosition = FormStartPosition.CenterParent;
			this.FormBorderStyle = FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.BackColor = Color.White;
			this.Size = new Size(920, 620);

			lblTitle = new Label
			{
				Text = "Object Template Library",
				Font = new Font("Segoe UI", 12F, FontStyle.Bold),
				ForeColor = Color.FromArgb(68, 114, 196),
				Location = new Point(16, 14),
				AutoSize = true
			};

			separator = new Panel
			{
				BackColor = Color.FromArgb(230, 230, 230),
				Location = new Point(16, 44),
				Size = new Size(880, 1)
			};

			txtSearch = new TextBox
			{
				Font = new Font("Segoe UI", 9F),
				Location = new Point(620, 12),
				Size = new Size(276, 23)
			};
			// Simple placeholder for .NET Framework
			txtSearch.ForeColor = Color.Gray;
			txtSearch.Text = "Search...";
			txtSearch.GotFocus += (s, e) => { if (txtSearch.ForeColor == Color.Gray) { txtSearch.Text = string.Empty; txtSearch.ForeColor = Color.Black; } };
			txtSearch.LostFocus += (s, e) => { if (string.IsNullOrWhiteSpace(txtSearch.Text)) { txtSearch.Text = "Search..."; txtSearch.ForeColor = Color.Gray; } };
			txtSearch.TextChanged += (s, e) => { if (txtSearch.ForeColor != Color.Gray) ApplyFilters(); };

			lstCategories = new ListBox
			{
				Font = new Font("Segoe UI", 9F),
				Location = new Point(16, 56),
				Size = new Size(170, 480)
			};
			lstCategories.SelectedIndexChanged += (s, e) => ApplyFilters();

			flowTemplates = new FlowLayoutPanel
			{
				Location = new Point(196, 56),
				Size = new Size(700, 480),
				AutoScroll = true,
				WrapContents = true,
				FlowDirection = FlowDirection.LeftToRight,
				Margin = new Padding(0)
			};

			btnInsert = new Button
			{
				Text = "Insert",
				BackColor = Color.FromArgb(68, 114, 196),
				ForeColor = Color.White,
				Font = new Font("Segoe UI", 9F, FontStyle.Bold),
				FlatStyle = FlatStyle.Flat,
				Location = new Point(716, 548),
				Size = new Size(90, 32),
				DialogResult = DialogResult.OK
			};
			btnInsert.Click += (s, e) => { /* SelectedTemplateName already set */ };

			btnCancel = new Button
			{
				Text = "Cancel",
				Font = new Font("Segoe UI", 9F),
				FlatStyle = FlatStyle.Flat,
				Location = new Point(812, 548),
				Size = new Size(84, 32),
				DialogResult = DialogResult.Cancel
			};

			this.AcceptButton = btnInsert;
			this.CancelButton = btnCancel;

			this.Controls.Add(lblTitle);
			this.Controls.Add(separator);
			this.Controls.Add(txtSearch);
			this.Controls.Add(lstCategories);
			this.Controls.Add(flowTemplates);
			this.Controls.Add(btnInsert);
			this.Controls.Add(btnCancel);
		}

		private void LoadTemplates()
		{
			// Categories
			var categories = new[] { "Basic", "Shapes", "Text", "Navigation", "Colors", "Elements", "Position", "File", "Others" };
			lstCategories.Items.AddRange(categories);
			lstCategories.SelectedIndex = 0;

			// Define a set of templates. Images are loaded from embedded resources (icons/*/*.png)
			AddTemplate("line", "Basic", "icons8-open-file-48.png");
			AddTemplate("square_thin", "Shapes", "icons8-align-center-64.png", resourceFolder: "position");
			AddTemplate("round_thin", "Shapes", "icons8-align-justify-64.png", resourceFolder: "position");
			AddTemplate("textbox_no_frame", "Text", "icons8-text.png", resourceFolder: "color");
			AddTemplate("textbox_frame", "Text", "icons8-color-palette-48.png", resourceFolder: "elements");
			AddTemplate("textbox_frame_colored", "Text", "icons8-object-50.png", resourceFolder: "elements");
			AddTemplate("box_A", "Elements", "icons8-table-50.png", resourceFolder: "elements");
			AddTemplate("box_B", "Elements", "icons8-object-50.png", resourceFolder: "elements");
			AddTemplate("box_C", "Elements", "icons8-chart-60.png", resourceFolder: "elements");
			AddTemplate("circle_left", "Navigation", "icons8-previous-slide.png", resourceFolder: "navigation");
			AddTemplate("circle_right", "Navigation", "icons8-next-slide.png", resourceFolder: "navigation");
			AddTemplate("circle_up", "Navigation", "icons8-zoom-in.png", resourceFolder: "navigation");
			AddTemplate("fill_color", "Colors", "icons8-fill.png", resourceFolder: "color");
			AddTemplate("text_color", "Colors", "icons8-text.png", resourceFolder: "color");
			AddTemplate("outline_color", "Colors", "icons8-outline.png", resourceFolder: "color");
		}

		private void AddTemplate(string name, string category, string fileName, string resourceFolder = null)
		{
			var item = new TemplateItem
			{
				Name = name,
				Category = category,
				Image = LoadEmbeddedPng(fileName, resourceFolder)
			};
			allTemplates.Add(item);
		}

		private Image LoadEmbeddedPng(string fileName, string resourceFolder)
		{
			try
			{
				var asm = Assembly.GetExecutingAssembly();
				var names = asm.GetManifestResourceNames();
				string probe = fileName;
				if (!string.IsNullOrEmpty(resourceFolder))
				{
					probe = $".{resourceFolder}.{fileName}";
				}
				var resourceName = names.FirstOrDefault(n => n.EndsWith(probe, StringComparison.OrdinalIgnoreCase));
				if (resourceName == null)
				{
					// Try common prefix used by this project (my_addin.icons.*)
					resourceName = names.FirstOrDefault(n => n.EndsWith($".icons.{resourceFolder}.{fileName}", StringComparison.OrdinalIgnoreCase) || n.EndsWith($".icons.{fileName}", StringComparison.OrdinalIgnoreCase));
				}
				if (resourceName == null)
					return null;
				using (var s = asm.GetManifestResourceStream(resourceName))
				{
					return s != null ? Image.FromStream(s) : null;
				}
			}
			catch
			{
				return null;
			}
		}

		private void ApplyFilters()
		{
			var term = (txtSearch.Text ?? string.Empty).Trim();
			if (txtSearch.ForeColor == Color.Gray) term = string.Empty;
			var cat = lstCategories.SelectedItem as string;

			var filtered = allTemplates.AsEnumerable();
			if (!string.IsNullOrEmpty(cat))
				filtered = filtered.Where(t => string.Equals(t.Category, cat, StringComparison.OrdinalIgnoreCase));
			if (!string.IsNullOrEmpty(term))
				filtered = filtered.Where(t => t.Name.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0);

			RenderTiles(filtered.ToList());
		}

		private void RenderTiles(List<TemplateItem> items)
		{
			flowTemplates.SuspendLayout();
			flowTemplates.Controls.Clear();

			foreach (var item in items)
			{
				var tile = CreateTile(item);
				flowTemplates.Controls.Add(tile);
			}

			flowTemplates.ResumeLayout();
		}

		private Control CreateTile(TemplateItem item)
		{
			var tile = new Panel
			{
				Width = 200,
				Height = 140,
				Margin = new Padding(10),
				BackColor = Color.White,
				BorderStyle = BorderStyle.FixedSingle,
				Tag = item
			};

			var pic = new PictureBox
			{
				Image = item.Image,
				SizeMode = PictureBoxSizeMode.Zoom,
				Location = new Point(10, 10),
				Size = new Size(178, 86)
			};

			var lbl = new Label
			{
				Text = item.Name,
				Font = new Font("Segoe UI", 9F),
				AutoSize = false,
				TextAlign = ContentAlignment.MiddleLeft,
				Location = new Point(10, 104),
				Size = new Size(178, 24)
			};

			void selectThis()
			{
				selectedTemplate = (TemplateItem)tile.Tag;
				foreach (Control c in flowTemplates.Controls)
				{
					if (c is Panel p)
						p.BackColor = Color.White;
				}
				tile.BackColor = Color.FromArgb(240, 248, 255);
			}

			tile.Click += (s, e) => selectThis();
			pic.Click += (s, e) => selectThis();
			lbl.Click += (s, e) => selectThis();
			tile.DoubleClick += (s, e) => { selectThis(); this.DialogResult = DialogResult.OK; };
			pic.DoubleClick += (s, e) => { selectThis(); this.DialogResult = DialogResult.OK; };
			lbl.DoubleClick += (s, e) => { selectThis(); this.DialogResult = DialogResult.OK; };

			tile.Controls.Add(pic);
			tile.Controls.Add(lbl);
			return tile;
		}

		private class TemplateItem
		{
			public string Name { get; set; }
			public string Category { get; set; }
			public Image Image { get; set; }
		}
	}
} 