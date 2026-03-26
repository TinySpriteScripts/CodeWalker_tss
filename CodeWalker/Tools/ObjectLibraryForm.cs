using SixLabors.ImageSharp.Processing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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

        private const int MaxDisplayedCards = 300;
        private const int CardWidth = 190;
        private const int CardHeight = 220;

        private readonly TextBox SearchTextBox;
        private readonly Label StatusLabel;
        private readonly Label ResultLabel;
        private readonly FlowLayoutPanel ResultsPanel;
        private readonly System.Windows.Forms.Timer FilterDebounceTimer;
        private readonly ToolTip CardToolTip;

        private readonly List<ObjectEntry> AllEntries = new List<ObjectEntry>();
        private readonly Dictionary<string, System.Drawing.Image> ThumbnailCache = new Dictionary<string, System.Drawing.Image>(StringComparer.OrdinalIgnoreCase);
        private readonly object ThumbnailCacheSync = new object();
        private readonly SemaphoreSlim ThumbnailLoadSemaphore = new SemaphoreSlim(2, 2);

        private CancellationTokenSource LoadCancellation;

        public ObjectLibraryForm()
        {
            Text = "Object Library";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(760, 500);
            Size = new Size(1080, 720);

            var topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 66,
                Padding = new Padding(10, 10, 10, 8)
            };

            var searchLabel = new Label
            {
                Text = "Search:",
                AutoSize = true,
                Location = new Point(10, 14)
            };

            SearchTextBox = new TextBox
            {
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Location = new Point(64, 10),
                Width = 640
            };
            SearchTextBox.TextChanged += SearchTextBox_TextChanged;

            ResultLabel = new Label
            {
                AutoSize = true,
                Location = new Point(64, 40)
            };

            StatusLabel = new Label
            {
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                AutoSize = true,
                Location = new Point(710, 40),
                TextAlign = ContentAlignment.MiddleRight
            };

            topPanel.Controls.Add(searchLabel);
            topPanel.Controls.Add(SearchTextBox);
            topPanel.Controls.Add(ResultLabel);
            topPanel.Controls.Add(StatusLabel);

            ResultsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(8),
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight
            };

            Controls.Add(ResultsPanel);
            Controls.Add(topPanel);

            FilterDebounceTimer = new System.Windows.Forms.Timer { Interval = 180 };
            FilterDebounceTimer.Tick += FilterDebounceTimer_Tick;
            CardToolTip = new ToolTip();

            FormClosed += ObjectLibraryForm_FormClosed;
            Load += ObjectLibraryForm_Load;
            Resize += ObjectLibraryForm_Resize;
        }

        private async void ObjectLibraryForm_Load(object sender, EventArgs e)
        {
            await LoadIndexAsync();
            ApplyFilter();
        }

        private void ObjectLibraryForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            LoadCancellation?.Cancel();
            LoadCancellation?.Dispose();
            LoadCancellation = null;

            ThumbnailLoadSemaphore.Dispose();

            lock (ThumbnailCacheSync)
            {
                foreach (var image in ThumbnailCache.Values)
                {
                    image.Dispose();
                }
                ThumbnailCache.Clear();
            }

            FilterDebounceTimer.Stop();
            FilterDebounceTimer.Dispose();
            CardToolTip.Dispose();
        }

        private void ObjectLibraryForm_Resize(object sender, EventArgs e)
        {
            StatusLabel.Left = Math.Max(280, ClientSize.Width - StatusLabel.Width - 20);
            SearchTextBox.Width = Math.Max(240, ClientSize.Width - 220);
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
                try
                {
                    return Directory.EnumerateFiles(imageDbPath, "*.webp", SearchOption.AllDirectories).ToList();
                }
                catch
                {
                    return new List<string>();
                }
            });

            AllEntries.Clear();
            AllEntries.AddRange(files.Select(file => new ObjectEntry
            {
                FilePath = file,
                ModelName = Path.GetFileNameWithoutExtension(file)
            }).OrderBy(entry => entry.ModelName, StringComparer.OrdinalIgnoreCase));

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
            if (!string.IsNullOrEmpty(term))
            {
                query = query.Where(entry => entry.ModelName.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            var filtered = query.ToList();
            var display = filtered.Take(MaxDisplayedCards).ToList();

            ResultLabel.Text = filtered.Count > MaxDisplayedCards
                ? $"Showing first {MaxDisplayedCards:N0} of {filtered.Count:N0} matches. Refine search to narrow results."
                : $"Showing {filtered.Count:N0} match(es).";

            ResultsPanel.SuspendLayout();
            while (ResultsPanel.Controls.Count > 0)
            {
                var control = ResultsPanel.Controls[0];
                ResultsPanel.Controls.RemoveAt(0);
                control.Dispose();
            }

            foreach (var entry in display)
            {
                ResultsPanel.Controls.Add(CreateCard(entry, token));
            }

            ResultsPanel.ResumeLayout();

            if (AllEntries.Count == 0)
            {
                SetStatus("No images indexed.");
            }
            else if (display.Count == 0)
            {
                SetStatus("No matches.");
            }
            else
            {
                SetStatus("Ready.");
            }
        }

        private Control CreateCard(ObjectEntry entry, CancellationToken token)
        {
            var card = new Panel
            {
                Width = CardWidth,
                Height = CardHeight,
                Margin = new Padding(8),
                BorderStyle = BorderStyle.FixedSingle
            };

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

            var copyButton = new Button
            {
                Left = CardWidth - 44,
                Top = 184,
                Width = 34,
                Height = 28,
                Text = "📋",
                Cursor = Cursors.Hand
            };
            copyButton.Click += (s, e) => CopyModelName(entry.ModelName);
            CardToolTip.SetToolTip(copyButton, "Copy model name");

            card.Controls.Add(pictureBox);
            card.Controls.Add(nameLabel);
            card.Controls.Add(copyButton);

            _ = LoadThumbnailAsync(entry, pictureBox, token);

            return card;
        }

        private async Task LoadThumbnailAsync(ObjectEntry entry, PictureBox target, CancellationToken token)
        {
            if (token.IsCancellationRequested || target.IsDisposed) return;

            System.Drawing.Image cachedImage;
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

                if (cachedImage != null)
                {
                    SetPictureImage(target, cachedImage);
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch
            {
            }
            finally
            {
                if (lockTaken)
                {
                    ThumbnailLoadSemaphore.Release();
                }
            }
        }

        private static System.Drawing.Image LoadWebpThumbnail(string filePath, int maxWidth, int maxHeight)
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

        private void SetPictureImage(PictureBox pictureBox, System.Drawing.Image image)
        {
            if (pictureBox.IsDisposed) return;

            if (pictureBox.InvokeRequired)
            {
                pictureBox.BeginInvoke((MethodInvoker)delegate
                {
                    if (!pictureBox.IsDisposed)
                    {
                        pictureBox.Image = image;
                    }
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
                SetStatus($"Copied: {modelName}");
            }
            catch (Exception ex)
            {
                SetStatus($"Copy failed: {ex.Message}");
            }
        }

        private void SetStatus(string text)
        {
            if (StatusLabel.InvokeRequired)
            {
                StatusLabel.BeginInvoke((MethodInvoker)delegate { StatusLabel.Text = text; });
            }
            else
            {
                StatusLabel.Text = text;
            }
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
                    if (Directory.Exists(candidate))
                    {
                        return candidate;
                    }
                    current = current.Parent;
                }
            }

            return null;
        }
    }
}
