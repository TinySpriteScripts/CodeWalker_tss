using CodeWalker.Utils;
using SixLabors.ImageSharp.Processing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CodeWalker.Tools
{
    public class ObjectLibraryForm : Form
    {
        private class ObjectEntry
        {
            public string FilePath { get; set; }
            public string ModelName { get; set; }
        }

        [DataContract]
        private class ObjectSetFile
        {
            [DataMember(Name = "name")]
            public string Name { get; set; }

            [DataMember(Name = "author")]
            public string Author { get; set; }

            [DataMember(Name = "contents")]
            public List<string> Contents { get; set; } = new List<string>();

            public override string ToString() => Name ?? "Unnamed Set";
        }

        private const int MaxDisplayedCards = 300;
        private const int CardWidth = 190;
        private const int CardHeight = 220;

        private readonly TextBox SearchTextBox;
        private readonly Label StatusLabel;
        private readonly Label ResultLabel;
        private readonly ComboBox SetComboBox;
        private readonly Label SetAuthorLabel;
        private readonly TabControl ResultsTabControl;
        private readonly FlowLayoutPanel AllResultsPanel;
        private readonly FlowLayoutPanel FavoriteResultsPanel;
        private readonly FlowLayoutPanel SetResultsPanel;
        private readonly System.Windows.Forms.Timer FilterDebounceTimer;
        private readonly ToolTip CardToolTip;

        private readonly List<ObjectEntry> AllEntries = new List<ObjectEntry>();
        private readonly HashSet<string> FavoriteModels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Image> ThumbnailCache = new Dictionary<string, Image>(StringComparer.OrdinalIgnoreCase);
        private readonly object ThumbnailCacheSync = new object();
        private readonly SemaphoreSlim ThumbnailLoadSemaphore = new SemaphoreSlim(2, 2);

        private readonly Dictionary<string, ObjectSetFile> ObjectSetsByPath = new Dictionary<string, ObjectSetFile>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, FlowLayoutPanel> SetPanelsByPath = new Dictionary<string, FlowLayoutPanel>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, TabPage> SetTabsByPath = new Dictionary<string, TabPage>(StringComparer.OrdinalIgnoreCase);
        private CancellationTokenSource LoadCancellation;

        public ObjectLibraryForm()
        {
            Text = "Object Library";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(920, 560);
            Size = new Size(1180, 760);

            var topPanel = new Panel { Dock = DockStyle.Top, Height = 96, Padding = new Padding(10, 10, 10, 8) };

            var searchLabel = new Label { Text = "Search:", AutoSize = true, Location = new Point(10, 14) };
            SearchTextBox = new TextBox { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, Location = new Point(64, 10), Width = 520 };
            SearchTextBox.TextChanged += SearchTextBox_TextChanged;

            var setLabel = new Label { Text = "Set:", AutoSize = true, Location = new Point(10, 46) };
            SetComboBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(64, 42), Width = 280 };
            SetComboBox.SelectedIndexChanged += SetComboBox_SelectedIndexChanged;

            var newSetButton = new Button { Text = "New Set", Location = new Point(350, 41), Width = 76, Height = 25 };
            newSetButton.Click += NewSetButton_Click;
            var importSetButton = new Button { Text = "Import", Location = new Point(432, 41), Width = 70, Height = 25 };
            importSetButton.Click += ImportSetButton_Click;
            var exportSetButton = new Button { Text = "Export", Location = new Point(508, 41), Width = 70, Height = 25 };
            exportSetButton.Click += ExportSetButton_Click;

            SetAuthorLabel = new Label { AutoSize = true, Location = new Point(586, 46) };
            ResultLabel = new Label { AutoSize = true, Location = new Point(64, 74) };
            StatusLabel = new Label { Anchor = AnchorStyles.Top | AnchorStyles.Right, AutoSize = true, Location = new Point(760, 74), TextAlign = ContentAlignment.MiddleRight };

            topPanel.Controls.Add(searchLabel);
            topPanel.Controls.Add(SearchTextBox);
            topPanel.Controls.Add(setLabel);
            topPanel.Controls.Add(SetComboBox);
            topPanel.Controls.Add(newSetButton);
            topPanel.Controls.Add(importSetButton);
            topPanel.Controls.Add(exportSetButton);
            topPanel.Controls.Add(SetAuthorLabel);
            topPanel.Controls.Add(ResultLabel);
            topPanel.Controls.Add(StatusLabel);

            ResultsTabControl = new TabControl { Dock = DockStyle.Fill };
            var allTab = new TabPage("All Objects");
            var favoriteTab = new TabPage("Favorites");
            var setTab = new TabPage("In Sets");

            AllResultsPanel = NewResultsPanel();
            FavoriteResultsPanel = NewResultsPanel();
            SetResultsPanel = NewResultsPanel();

            allTab.Controls.Add(AllResultsPanel);
            favoriteTab.Controls.Add(FavoriteResultsPanel);
            setTab.Controls.Add(SetResultsPanel);

            ResultsTabControl.TabPages.Add(allTab);
            ResultsTabControl.TabPages.Add(favoriteTab);
            ResultsTabControl.TabPages.Add(setTab);
            ResultsTabControl.SelectedIndexChanged += ResultsTabControl_SelectedIndexChanged;

            Controls.Add(ResultsTabControl);
            Controls.Add(topPanel);

            FilterDebounceTimer = new System.Windows.Forms.Timer { Interval = 180 };
            FilterDebounceTimer.Tick += FilterDebounceTimer_Tick;
            CardToolTip = new ToolTip();

            FormClosed += ObjectLibraryForm_FormClosed;
            Load += ObjectLibraryForm_Load;
            Resize += ObjectLibraryForm_Resize;

            ApplyThemeFromSettings();
        }

        private void ResultsTabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            var path = GetActiveSetPath();
            if (!string.IsNullOrWhiteSpace(path) && SetComboBox.Items.Contains(path) && !string.Equals(SetComboBox.SelectedItem as string, path, StringComparison.OrdinalIgnoreCase))
            {
                SetComboBox.SelectedItem = path;
            }
            UpdateSetHeader();
        }

        private static FlowLayoutPanel NewResultsPanel()
        {
            return new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(8), WrapContents = true, FlowDirection = FlowDirection.LeftToRight };
        }

        private async void ObjectLibraryForm_Load(object sender, EventArgs e)
        {
            LoadFavorites();
            LoadSets();
            await LoadIndexAsync();
            ApplyFilter();
        }

        private void ObjectLibraryForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            SaveFavorites();
            LoadCancellation?.Cancel();
            LoadCancellation?.Dispose();
            LoadCancellation = null;
            ThumbnailLoadSemaphore.Dispose();
            lock (ThumbnailCacheSync)
            {
                foreach (var image in ThumbnailCache.Values) image.Dispose();
                ThumbnailCache.Clear();
            }
            FilterDebounceTimer.Stop();
            FilterDebounceTimer.Dispose();
            CardToolTip.Dispose();
        }

        private void ObjectLibraryForm_Resize(object sender, EventArgs e)
        {
            StatusLabel.Left = Math.Max(330, ClientSize.Width - StatusLabel.Width - 20);
            SearchTextBox.Width = Math.Max(280, ClientSize.Width - 680);
        }

        private void SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            FilterDebounceTimer.Stop();
            FilterDebounceTimer.Start();
        }

        private void FilterDebounceTimer_Tick(object sender, EventArgs e)
        {
            FilterDebounceTimer.Stop();
            ApplyFilter();
        }

        private void SetComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateSetHeader();
            UpdateSetButtonsInPanel(AllResultsPanel, null);
            UpdateSetButtonsInPanel(FavoriteResultsPanel, null);
            UpdateSetButtonsInPanel(SetResultsPanel, null);
        }

        private async Task LoadIndexAsync()
        {
            SetStatus("Loading object index...");
            var imageDbPath = FindImageDbPath();
            if (string.IsNullOrEmpty(imageDbPath))
            {
                SetStatus("ImageDB folder not found.");
                return;
            }

            var files = await Task.Run(() =>
            {
                try { return Directory.EnumerateFiles(imageDbPath, "*.webp", SearchOption.AllDirectories).ToList(); }
                catch { return new List<string>(); }
            });

            AllEntries.Clear();
            AllEntries.AddRange(files.Select(file => new ObjectEntry { FilePath = file, ModelName = Path.GetFileNameWithoutExtension(file) })
                                     .OrderBy(entry => entry.ModelName, StringComparer.OrdinalIgnoreCase));
            SetStatus($"Loaded {AllEntries.Count:N0} objects.");
        }

        private void ApplyFilter()
        {
            LoadCancellation?.Cancel();
            LoadCancellation?.Dispose();
            LoadCancellation = new CancellationTokenSource();
            var token = LoadCancellation.Token;

            var term = (SearchTextBox.Text ?? string.Empty).Trim();
            IEnumerable<ObjectEntry> query = AllEntries;
            if (!string.IsNullOrEmpty(term)) query = query.Where(entry => entry.ModelName.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0);

            var filteredAll = query.ToList();
            var filteredFavorites = filteredAll.Where(entry => IsFavorite(entry.ModelName)).ToList();
            var filteredSet = filteredAll.Where(entry => IsInAnySet(entry.ModelName)).ToList();

            PopulateResultsPanel(AllResultsPanel, filteredAll.Take(MaxDisplayedCards), token);
            PopulateResultsPanel(FavoriteResultsPanel, filteredFavorites.Take(MaxDisplayedCards), token);
            PopulateResultsPanel(SetResultsPanel, filteredSet.Take(MaxDisplayedCards), token);
            PopulateAllSetTabs(filteredAll, token);

            ResultLabel.Text =
                $"All: {filteredAll.Count:N0}" + (filteredAll.Count > MaxDisplayedCards ? $" (showing {MaxDisplayedCards:N0})" : string.Empty) +
                $" | Favorites: {filteredFavorites.Count:N0}" + (filteredFavorites.Count > MaxDisplayedCards ? $" (showing {MaxDisplayedCards:N0})" : string.Empty) +
                $" | Set: {filteredSet.Count:N0}" + (filteredSet.Count > MaxDisplayedCards ? $" (showing {MaxDisplayedCards:N0})" : string.Empty);

            if (AllEntries.Count == 0) SetStatus("No images indexed.");
            else if (filteredAll.Count == 0) SetStatus("No matches.");
            else SetStatus("Ready.");
        }

        private void PopulateAllSetTabs(List<ObjectEntry> filteredAll, CancellationToken token)
        {
            foreach (var kvp in SetPanelsByPath)
            {
                var setPath = kvp.Key;
                var panel = kvp.Value;
                var set = ObjectSetsByPath.ContainsKey(setPath) ? ObjectSetsByPath[setPath] : null;
                if (set == null)
                {
                    PopulateResultsPanel(panel, Enumerable.Empty<ObjectEntry>(), token);
                    continue;
                }

                var entries = filteredAll.Where(entry => IsInSet(set, entry.ModelName)).Take(MaxDisplayedCards);
                PopulateResultsPanel(panel, entries, token);
            }
        }

        private void PopulateResultsPanel(FlowLayoutPanel panel, IEnumerable<ObjectEntry> entries, CancellationToken token)
        {
            panel.SuspendLayout();
            while (panel.Controls.Count > 0)
            {
                var control = panel.Controls[0];
                panel.Controls.RemoveAt(0);
                control.Dispose();
            }
            foreach (var entry in entries) panel.Controls.Add(CreateCard(entry, token));
            panel.ResumeLayout();
        }

        private Control CreateCard(ObjectEntry entry, CancellationToken token)
        {
            var card = new Panel { Width = CardWidth, Height = CardHeight, Margin = new Padding(8), BorderStyle = BorderStyle.FixedSingle, Tag = entry.ModelName };

            var pictureBox = new PictureBox
            {
                Width = CardWidth - 12,
                Height = 132,
                Left = 6,
                Top = 6,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.FromArgb(245, 245, 245),
                BorderStyle = BorderStyle.FixedSingle
            };

            var nameLabel = new Label
            {
                Left = 8,
                Top = 146,
                Width = CardWidth - 16,
                Height = 34,
                Text = entry.ModelName,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.TopLeft
            };

            var favoriteButton = new Button
            {
                Left = 8,
                Top = 184,
                Width = 34,
                Height = 28,
                FlatStyle = FlatStyle.Flat,
                Font = new Font(FontFamily.GenericSansSerif, 12.0f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Tag = "fav"
            };
            UpdateFavoriteButton(favoriteButton, entry.ModelName);
            favoriteButton.Click += (s, e) => ToggleFavorite(entry.ModelName);

            var setButton = new Button
            {
                Left = 46,
                Top = 184,
                Width = 44,
                Height = 28,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Tag = "set"
            };
            UpdateSetButton(setButton, entry.ModelName);
            setButton.Click += (s, e) => ShowSetMenu(setButton, entry.ModelName);

            var copyButton = new Button
            {
                Left = CardWidth - 44,
                Top = 184,
                Width = 34,
                Height = 28,
                Text = "📋",
                Cursor = Cursors.Hand,
                Tag = "copy"
            };
            copyButton.Click += (s, e) => CopyModelName(entry.ModelName);
            CardToolTip.SetToolTip(copyButton, "Copy model name");

            card.Controls.Add(pictureBox);
            card.Controls.Add(nameLabel);
            card.Controls.Add(favoriteButton);
            card.Controls.Add(setButton);
            card.Controls.Add(copyButton);

            _ = LoadThumbnailAsync(entry, pictureBox, token);
            return card;
        }

        private void ToggleFavorite(string modelName)
        {
            if (string.IsNullOrWhiteSpace(modelName)) return;

            var added = FavoriteModels.Add(modelName);
            if (!added) FavoriteModels.Remove(modelName);

            SaveFavorites();
            SetStatus((added ? "Added to favorites: " : "Removed from favorites: ") + modelName);
            UpdateFavoriteButtonsInPanel(AllResultsPanel, modelName);
            UpdateFavoriteButtonsInPanel(FavoriteResultsPanel, modelName);
            UpdateFavoriteButtonsInPanel(SetResultsPanel, modelName);
            RefreshFavoritesPanelForModel(modelName, added);
            UpdateResultCounts();
        }

        private void ShowSetMenu(Control anchor, string modelName)
        {
            if (ObjectSetsByPath.Count == 0)
            {
                SetStatus("No sets found. Create or import a set first.");
                return;
            }

            var menu = new ContextMenuStrip();
            foreach (var path in ObjectSetsByPath.Keys.OrderBy(p => ObjectSetsByPath[p].Name, StringComparer.OrdinalIgnoreCase))
            {
                var set = ObjectSetsByPath[path];
                var item = new ToolStripMenuItem(set.Name)
                {
                    Checked = IsInSet(set, modelName),
                    Tag = path
                };
                item.Click += (s, e) =>
                {
                    ToggleMembershipInSet((string)item.Tag, modelName);
                };
                menu.Items.Add(item);
            }
            menu.Show(anchor, new Point(0, anchor.Height));
        }

        private void ToggleMembershipInSet(string setPath, string modelName)
        {
            if (string.IsNullOrWhiteSpace(setPath) || string.IsNullOrWhiteSpace(modelName)) return;
            ObjectSetFile set;
            if (!ObjectSetsByPath.TryGetValue(setPath, out set) || set == null) return;

            var contents = set.Contents ?? new List<string>();
            var idx = contents.FindIndex(x => string.Equals(x, modelName, StringComparison.OrdinalIgnoreCase));
            var added = idx < 0;
            if (added) contents.Add(modelName);
            else contents.RemoveAt(idx);
            set.Contents = contents
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            WriteJson(setPath, set);
            SetStatus((added ? "Added to set: " : "Removed from set: ") + modelName + " [" + set.Name + "]");

            UpdateSetButtonsInPanel(AllResultsPanel, modelName);
            UpdateSetButtonsInPanel(FavoriteResultsPanel, modelName);
            UpdateSetButtonsInPanel(SetResultsPanel, modelName);
            foreach (var panel in SetPanelsByPath.Values)
            {
                UpdateSetButtonsInPanel(panel, modelName);
            }

            RefreshSingleSetPanelForModel(setPath, modelName, added);
            RefreshAggregateSetPanelForModel(modelName);
            UpdateResultCounts();
        }

        private bool IsFavorite(string modelName) => !string.IsNullOrWhiteSpace(modelName) && FavoriteModels.Contains(modelName);

        private bool IsInSelectedSet(string modelName)
        {
            var selected = GetSelectedSet();
            return (selected != null) && (selected.Contents != null) && selected.Contents.Any(x => string.Equals(x, modelName, StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsInSet(ObjectSetFile set, string modelName)
        {
            return (set != null) && (set.Contents != null) && !string.IsNullOrWhiteSpace(modelName) &&
                   set.Contents.Any(x => string.Equals(x, modelName, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsInAnySet(string modelName)
        {
            if (string.IsNullOrWhiteSpace(modelName)) return false;
            return ObjectSetsByPath.Values.Any(set => IsInSet(set, modelName));
        }

        private void UpdateFavoriteButton(Button button, string modelName)
        {
            var isFavorite = IsFavorite(modelName);
            button.Text = "♥";
            button.ForeColor = isFavorite ? Color.Red : Color.Gray;
            CardToolTip.SetToolTip(button, isFavorite ? "Remove from favorites" : "Add to favorites");
        }

        private void RefreshFavoritesPanelForModel(string modelName, bool isFavoriteNow)
        {
            if (string.IsNullOrWhiteSpace(modelName)) return;
            if (!MatchesSearch(modelName)) return;

            for (int i = FavoriteResultsPanel.Controls.Count - 1; i >= 0; i--)
            {
                var card = FavoriteResultsPanel.Controls[i] as Panel;
                if ((card != null) && string.Equals(card.Tag as string, modelName, StringComparison.OrdinalIgnoreCase))
                {
                    FavoriteResultsPanel.Controls.RemoveAt(i);
                    card.Dispose();
                }
            }

            if (!isFavoriteNow) return;
            if (FavoriteResultsPanel.Controls.Count >= MaxDisplayedCards) return;

            var token = (LoadCancellation != null) ? LoadCancellation.Token : CancellationToken.None;
            var matches = AllEntries.Where(entry => string.Equals(entry.ModelName, modelName, StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (var entry in matches)
            {
                if (FavoriteResultsPanel.Controls.Count >= MaxDisplayedCards) break;
                FavoriteResultsPanel.Controls.Add(CreateCard(entry, token));
            }
        }

        private bool MatchesSearch(string modelName)
        {
            var term = (SearchTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(term)) return true;
            return modelName.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private int GetFilteredAllCount()
        {
            var term = (SearchTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(term)) return AllEntries.Count;
            return AllEntries.Count(entry => entry.ModelName.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private int GetFilteredFavoritesCount()
        {
            var term = (SearchTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(term))
            {
                return AllEntries.Count(entry => IsFavorite(entry.ModelName));
            }
            return AllEntries.Count(entry => IsFavorite(entry.ModelName) && entry.ModelName.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private int GetFilteredSetCount()
        {
            var term = (SearchTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(term))
            {
                return AllEntries.Count(entry => IsInAnySet(entry.ModelName));
            }
            return AllEntries.Count(entry => IsInAnySet(entry.ModelName) && entry.ModelName.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void UpdateResultCounts()
        {
            var allCount = GetFilteredAllCount();
            var favoriteCount = GetFilteredFavoritesCount();
            var setCount = GetFilteredSetCount();

            ResultLabel.Text =
                $"All: {allCount:N0}" + (allCount > MaxDisplayedCards ? $" (showing {MaxDisplayedCards:N0})" : string.Empty) +
                $" | Favorites: {favoriteCount:N0}" + (favoriteCount > MaxDisplayedCards ? $" (showing {MaxDisplayedCards:N0})" : string.Empty) +
                $" | Set: {setCount:N0}" + (setCount > MaxDisplayedCards ? $" (showing {MaxDisplayedCards:N0})" : string.Empty);
        }

        private void UpdateSetButton(Button button, string modelName)
        {
            var inSet = IsInAnySet(modelName);
            button.Text = "SET";
            button.ForeColor = inSet ? Color.MediumBlue : Color.Gray;
            if (!inSet)
            {
                CardToolTip.SetToolTip(button, "Add to one or more sets");
            }
            else
            {
                var setNames = ObjectSetsByPath.Values
                    .Where(set => IsInSet(set, modelName))
                    .Select(set => set.Name)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var tip = "In sets: " + string.Join(", ", setNames.Take(6));
                if (setNames.Count > 6) tip += ", ...";
                CardToolTip.SetToolTip(button, tip);
            }
        }

        private void RefreshSingleSetPanelForModel(string setPath, string modelName, bool nowInSet)
        {
            FlowLayoutPanel panel;
            if (!SetPanelsByPath.TryGetValue(setPath, out panel) || panel == null) return;
            if (!MatchesSearch(modelName)) return;

            for (int i = panel.Controls.Count - 1; i >= 0; i--)
            {
                var card = panel.Controls[i] as Panel;
                if ((card != null) && string.Equals(card.Tag as string, modelName, StringComparison.OrdinalIgnoreCase))
                {
                    panel.Controls.RemoveAt(i);
                    card.Dispose();
                }
            }

            if (!nowInSet) return;
            if (panel.Controls.Count >= MaxDisplayedCards) return;

            var token = (LoadCancellation != null) ? LoadCancellation.Token : CancellationToken.None;
            var matches = AllEntries.Where(entry => string.Equals(entry.ModelName, modelName, StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (var entry in matches)
            {
                if (panel.Controls.Count >= MaxDisplayedCards) break;
                panel.Controls.Add(CreateCard(entry, token));
            }
        }

        private void RefreshAggregateSetPanelForModel(string modelName)
        {
            if (string.IsNullOrWhiteSpace(modelName)) return;
            if (!MatchesSearch(modelName)) return;

            for (int i = SetResultsPanel.Controls.Count - 1; i >= 0; i--)
            {
                var card = SetResultsPanel.Controls[i] as Panel;
                if ((card != null) && string.Equals(card.Tag as string, modelName, StringComparison.OrdinalIgnoreCase))
                {
                    SetResultsPanel.Controls.RemoveAt(i);
                    card.Dispose();
                }
            }

            if (!IsInAnySet(modelName)) return;
            if (SetResultsPanel.Controls.Count >= MaxDisplayedCards) return;

            var token = (LoadCancellation != null) ? LoadCancellation.Token : CancellationToken.None;
            var matches = AllEntries.Where(entry => string.Equals(entry.ModelName, modelName, StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (var entry in matches)
            {
                if (SetResultsPanel.Controls.Count >= MaxDisplayedCards) break;
                SetResultsPanel.Controls.Add(CreateCard(entry, token));
            }
        }

        private void UpdateFavoriteButtonsInPanel(FlowLayoutPanel panel, string modelName)
        {
            foreach (Control control in panel.Controls)
            {
                var card = control as Panel;
                if (card == null) continue;
                var cardModel = card.Tag as string;
                if ((modelName != null) && !string.Equals(cardModel, modelName, StringComparison.OrdinalIgnoreCase)) continue;

                foreach (Control child in card.Controls)
                {
                    var btn = child as Button;
                    if ((btn != null) && string.Equals(btn.Tag as string, "fav", StringComparison.Ordinal))
                    {
                        UpdateFavoriteButton(btn, cardModel);
                    }
                }
            }
        }

        private void UpdateSetButtonsInPanel(FlowLayoutPanel panel, string modelName)
        {
            foreach (Control control in panel.Controls)
            {
                var card = control as Panel;
                if (card == null) continue;
                var cardModel = card.Tag as string;
                if ((modelName != null) && !string.Equals(cardModel, modelName, StringComparison.OrdinalIgnoreCase)) continue;

                foreach (Control child in card.Controls)
                {
                    var btn = child as Button;
                    if ((btn != null) && string.Equals(btn.Tag as string, "set", StringComparison.Ordinal))
                    {
                        UpdateSetButton(btn, cardModel);
                    }
                }
            }
        }

        private void NewSetButton_Click(object sender, EventArgs e)
        {
            var nameInput = new TextInputForm
            {
                TitleText = "New Object Set",
                PromptText = "Set name:",
                MainText = string.Empty
            };
            if (nameInput.ShowDialog(this) != DialogResult.OK) return;

            var setName = (nameInput.MainText ?? string.Empty).Trim();
            if (setName.Length == 0)
            {
                SetStatus("Set name is required.");
                return;
            }

            var authorInput = new TextInputForm
            {
                TitleText = "New Object Set",
                PromptText = "Author:",
                MainText = Environment.UserName
            };
            if (authorInput.ShowDialog(this) != DialogResult.OK) return;

            var set = new ObjectSetFile
            {
                Name = setName,
                Author = string.IsNullOrWhiteSpace(authorInput.MainText) ? "Unknown" : authorInput.MainText.Trim(),
                Contents = new List<string>()
            };

            var path = GetUniqueSetPath(setName);
            if (!WriteJson(path, set))
            {
                SetStatus("Failed to save new set.");
                return;
            }

            LoadSets(path);
            SetStatus("Created set: " + setName);
            ApplyFilter();
        }

        private void ImportSetButton_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog { Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*", Title = "Import Object Set" };
            if (ofd.ShowDialog(this) != DialogResult.OK) return;

            ObjectSetFile set;
            if (!TryReadJson(ofd.FileName, out set) || (set == null) || string.IsNullOrWhiteSpace(set.Name))
            {
                SetStatus("Invalid set file.");
                return;
            }

            set.Name = set.Name.Trim();
            set.Author = string.IsNullOrWhiteSpace(set.Author) ? "Unknown" : set.Author.Trim();
            set.Contents = (set.Contents ?? new List<string>()).Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            var path = GetUniqueSetPath(set.Name);
            if (!WriteJson(path, set))
            {
                SetStatus("Failed to import set.");
                return;
            }

            LoadSets(path);
            SetStatus("Imported set: " + set.Name);
            ApplyFilter();
        }

        private void ExportSetButton_Click(object sender, EventArgs e)
        {
            var selected = GetSelectedSet();
            if (selected == null)
            {
                SetStatus("Select a set to export.");
                return;
            }

            var sfd = new SaveFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                FileName = MakeSafeFileName(selected.Name) + ".json",
                Title = "Export Object Set"
            };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            if (WriteJson(sfd.FileName, selected)) SetStatus("Exported set: " + selected.Name);
            else SetStatus("Export failed.");
        }

        private void LoadSets(string selectPath = null)
        {
            ObjectSetsByPath.Clear();
            var folder = GetSetsFolderPath();
            Directory.CreateDirectory(folder);

            foreach (var path in Directory.EnumerateFiles(folder, "*.json", SearchOption.TopDirectoryOnly))
            {
                ObjectSetFile set;
                if (!TryReadJson(path, out set) || (set == null) || string.IsNullOrWhiteSpace(set.Name)) continue;
                set.Name = set.Name.Trim();
                set.Author = string.IsNullOrWhiteSpace(set.Author) ? "Unknown" : set.Author.Trim();
                set.Contents = (set.Contents ?? new List<string>()).Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                ObjectSetsByPath[path] = set;
            }

            var selectedPath = selectPath;
            if (selectedPath == null && SetComboBox.SelectedItem != null) selectedPath = SetComboBox.SelectedItem.ToString();

            SetComboBox.BeginUpdate();
            SetComboBox.Items.Clear();
            foreach (var path in ObjectSetsByPath.Keys.OrderBy(p => ObjectSetsByPath[p].Name, StringComparer.OrdinalIgnoreCase))
            {
                SetComboBox.Items.Add(path);
            }
            SetComboBox.EndUpdate();

            if ((selectedPath != null) && SetComboBox.Items.Contains(selectedPath)) SetComboBox.SelectedItem = selectedPath;
            else if (SetComboBox.Items.Count > 0) SetComboBox.SelectedIndex = 0;
            else UpdateSetHeader();

            RebuildSetTabs();
            ApplyThemeFromSettings();
        }

        private void RebuildSetTabs()
        {
            foreach (var kvp in SetTabsByPath.ToList())
            {
                ResultsTabControl.TabPages.Remove(kvp.Value);
                kvp.Value.Dispose();
            }
            SetTabsByPath.Clear();
            SetPanelsByPath.Clear();

            foreach (var path in ObjectSetsByPath.Keys.OrderBy(p => ObjectSetsByPath[p].Name, StringComparer.OrdinalIgnoreCase))
            {
                var set = ObjectSetsByPath[path];
                var tab = new TabPage(set.Name);
                tab.Tag = path;
                var panel = NewResultsPanel();
                tab.Controls.Add(panel);
                ResultsTabControl.TabPages.Add(tab);
                SetTabsByPath[path] = tab;
                SetPanelsByPath[path] = panel;
            }
        }

        private void UpdateSetHeader()
        {
            var path = GetActiveSetPath() ?? (SetComboBox.SelectedItem as string);
            ObjectSetFile selected = null;
            if (!string.IsNullOrWhiteSpace(path)) ObjectSetsByPath.TryGetValue(path, out selected);
            if (selected == null)
            {
                SetAuthorLabel.Text = "Author: -";
            }
            else
            {
                SetAuthorLabel.Text = "Author: " + selected.Author;
            }
        }

        private ObjectSetFile GetSelectedSet()
        {
            var path = GetActiveSetPath() ?? (SetComboBox.SelectedItem as string);
            if (string.IsNullOrWhiteSpace(path)) return null;
            ObjectSetFile set;
            return ObjectSetsByPath.TryGetValue(path, out set) ? set : null;
        }

        private void SaveSelectedSet()
        {
            var path = GetActiveSetPath() ?? (SetComboBox.SelectedItem as string);
            var selected = GetSelectedSet();
            if (string.IsNullOrWhiteSpace(path) || selected == null) return;
            WriteJson(path, selected);
        }

        private string GetActiveSetPath()
        {
            var tab = ResultsTabControl.SelectedTab;
            if (tab == null) return null;
            var path = tab.Tag as string;
            if (string.IsNullOrWhiteSpace(path)) return null;
            return ObjectSetsByPath.ContainsKey(path) ? path : null;
        }

        private string GetSetsFolderPath()
        {
            var app = Application.StartupPath;
            if (string.IsNullOrWhiteSpace(app)) app = AppDomain.CurrentDomain.BaseDirectory;
            return Path.Combine(app, "ObjectSets");
        }

        private string GetUniqueSetPath(string setName)
        {
            var folder = GetSetsFolderPath();
            Directory.CreateDirectory(folder);

            var name = MakeSafeFileName(setName);
            if (name.Length == 0) name = "set";
            var path = Path.Combine(folder, name + ".json");
            var i = 2;
            while (File.Exists(path))
            {
                path = Path.Combine(folder, name + "_" + i + ".json");
                i++;
            }
            return path;
        }

        private static string MakeSafeFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            var chars = value.Trim().ToCharArray();
            var bad = Path.GetInvalidFileNameChars();
            for (var i = 0; i < chars.Length; i++) if (bad.Contains(chars[i])) chars[i] = '_';
            return new string(chars);
        }

        private static bool TryReadJson<T>(string path, out T value)
        {
            value = default(T);
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var ser = new DataContractJsonSerializer(typeof(T));
                    value = (T)ser.ReadObject(fs);
                    return true;
                }
            }
            catch { return false; }
        }

        private static bool WriteJson<T>(string path, T value)
        {
            try
            {
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    var ser = new DataContractJsonSerializer(typeof(T));
                    ser.WriteObject(fs, value);
                }
                return true;
            }
            catch { return false; }
        }

        private void LoadFavorites()
        {
            FavoriteModels.Clear();
            var path = GetFavoritesFilePath();
            if (!File.Exists(path)) return;
            try
            {
                foreach (var line in File.ReadAllLines(path))
                {
                    var model = (line ?? string.Empty).Trim();
                    if (model.Length > 0) FavoriteModels.Add(model);
                }
            }
            catch { }
        }

        private void SaveFavorites()
        {
            var path = GetFavoritesFilePath();
            try
            {
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.WriteAllLines(path, FavoriteModels.OrderBy(x => x, StringComparer.OrdinalIgnoreCase));
            }
            catch { }
        }

        private static string GetFavoritesFilePath()
        {
            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(docs, "CW_TSS", "object_library_favorites.txt");
        }

        private async Task LoadThumbnailAsync(ObjectEntry entry, PictureBox target, CancellationToken token)
        {
            if (token.IsCancellationRequested || target.IsDisposed) return;

            Image cachedImage;
            lock (ThumbnailCacheSync)
            {
                ThumbnailCache.TryGetValue(entry.FilePath, out cachedImage);
            }
            if (cachedImage != null)
            {
                SetPictureImage(target, cachedImage);
                return;
            }

            var lockTaken = false;
            try
            {
                await ThumbnailLoadSemaphore.WaitAsync(token);
                lockTaken = true;
                if (token.IsCancellationRequested || target.IsDisposed) return;

                lock (ThumbnailCacheSync)
                {
                    ThumbnailCache.TryGetValue(entry.FilePath, out cachedImage);
                }
                if (cachedImage == null)
                {
                    cachedImage = await Task.Run(() => LoadWebpThumbnail(entry.FilePath, 170, 128), token);
                    if (cachedImage != null)
                    {
                        lock (ThumbnailCacheSync)
                        {
                            ThumbnailCache[entry.FilePath] = cachedImage;
                        }
                    }
                }

                if (cachedImage != null) SetPictureImage(target, cachedImage);
            }
            catch (OperationCanceledException) { }
            catch { }
            finally
            {
                if (lockTaken) ThumbnailLoadSemaphore.Release();
            }
        }

        private static Image LoadWebpThumbnail(string filePath, int maxWidth, int maxHeight)
        {
            using (var image = SixLabors.ImageSharp.Image.Load(filePath))
            {
                image.Mutate(context => context.Resize(new ResizeOptions
                {
                    Size = new SixLabors.ImageSharp.Size(maxWidth, maxHeight),
                    Mode = ResizeMode.Max
                }));

                using (var ms = new MemoryStream())
                {
                    image.Save(ms, new SixLabors.ImageSharp.Formats.Bmp.BmpEncoder());
                    ms.Position = 0;
                    using (var tempBitmap = new Bitmap(ms))
                    {
                        return new Bitmap(tempBitmap);
                    }
                }
            }
        }

        private void SetPictureImage(PictureBox pictureBox, Image image)
        {
            if (pictureBox.IsDisposed) return;

            if (pictureBox.InvokeRequired)
            {
                pictureBox.BeginInvoke((MethodInvoker)delegate
                {
                    if (!pictureBox.IsDisposed) pictureBox.Image = image;
                });
            }
            else
            {
                pictureBox.Image = image;
            }
        }

        private void CopyModelName(string modelName)
        {
            try
            {
                Clipboard.SetText(modelName ?? string.Empty);
                SetStatus("Copied: " + modelName);
            }
            catch (Exception ex)
            {
                SetStatus("Copy failed: " + ex.Message);
            }
        }

        private void SetStatus(string text)
        {
            if (StatusLabel.InvokeRequired) StatusLabel.BeginInvoke((MethodInvoker)delegate { StatusLabel.Text = text; });
            else StatusLabel.Text = text;
        }

        private static string FindImageDbPath()
        {
            var probeRoots = new[]
            {
                Application.StartupPath,
                AppDomain.CurrentDomain.BaseDirectory,
                Environment.CurrentDirectory
            }.Where(path => !string.IsNullOrWhiteSpace(path))
             .Distinct(StringComparer.OrdinalIgnoreCase)
             .ToList();

            foreach (var root in probeRoots)
            {
                var current = new DirectoryInfo(root);
                for (var i = 0; i < 8 && current != null; i++)
                {
                    var candidate = Path.Combine(current.FullName, "ImageDB");
                    if (Directory.Exists(candidate)) return candidate;
                    current = current.Parent;
                }
            }
            return null;
        }

        public void ApplyThemeFromSettings()
        {
            var theme = Properties.Settings.Default.GlobalUITheme;
            var dark = AppThemeManager.IsDarkTheme(theme);
            var formBack = AppThemeManager.GetThemeFormBackColor(theme);
            var panelBack = AppThemeManager.GetThemePanelBackColor(theme);

            AppThemeManager.ApplyToForm(this, theme);

            BackColor = formBack;
            ResultsTabControl.BackColor = panelBack;

            foreach (TabPage tab in ResultsTabControl.TabPages)
            {
                tab.UseVisualStyleBackColor = !dark;
                tab.BackColor = dark ? formBack : SystemColors.Control;
                tab.ForeColor = dark ? Color.Gainsboro : SystemColors.ControlText;
            }

            AllResultsPanel.BackColor = dark ? panelBack : SystemColors.Control;
            FavoriteResultsPanel.BackColor = dark ? panelBack : SystemColors.Control;
            SetResultsPanel.BackColor = dark ? panelBack : SystemColors.Control;

            foreach (var panel in SetPanelsByPath.Values)
            {
                if (panel != null)
                {
                    panel.BackColor = dark ? panelBack : SystemColors.Control;
                }
            }
        }
    }
}
