using CodeWalker.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace CodeWalker.Utils
{
    public static class AppThemeManager
    {
        public const string ThemeTss = "TSS";
        public const string ThemeTssBlue = "TSS Blue";
        public const string ThemeDefaultCw = "Default CW";

        private sealed class ThemePalette
        {
            public Color FormBack;
            public Color PanelBack;
            public Color InputBack;
            public Color Text;
            public Color Accent;
            public Color StatusBack;
            public Color TabSelected;
            public Color TabBorder;
        }

        public static string NormalizeTheme(string themeName)
        {
            if (string.Equals(themeName, ThemeTssBlue, StringComparison.OrdinalIgnoreCase)) return ThemeTssBlue;
            if (string.Equals(themeName, ThemeDefaultCw, StringComparison.OrdinalIgnoreCase)) return ThemeDefaultCw;
            return ThemeTss;
        }

        public static bool IsDarkTheme(string themeName)
        {
            return NormalizeTheme(themeName) != ThemeDefaultCw;
        }

        public static Color GetThemeFormBackColor(string themeName)
        {
            if (!IsDarkTheme(themeName)) return SystemColors.Control;
            return NormalizeTheme(themeName) == ThemeTssBlue ? CreateBluePalette().FormBack : CreateRedPalette().FormBack;
        }

        public static Color GetThemePanelBackColor(string themeName)
        {
            if (!IsDarkTheme(themeName)) return SystemColors.Control;
            return NormalizeTheme(themeName) == ThemeTssBlue ? CreateBluePalette().PanelBack : CreateRedPalette().PanelBack;
        }

        public static string GetProjectWindowThemeForGlobal(string globalTheme)
        {
            switch (NormalizeTheme(globalTheme))
            {
                case ThemeDefaultCw: return "Blue";
                default: return "Dark";
            }
        }

        public static string GetExplorerWindowThemeForGlobal(string globalTheme)
        {
            switch (NormalizeTheme(globalTheme))
            {
                case ThemeDefaultCw: return "Windows";
                default: return "Dark";
            }
        }

        public static void SyncLegacyThemeSettings(string globalTheme)
        {
            Settings.Default.ProjectWindowTheme = GetProjectWindowThemeForGlobal(globalTheme);
            Settings.Default.ExplorerWindowTheme = GetExplorerWindowThemeForGlobal(globalTheme);
        }

        public static void ApplyToForm(Form form, string globalTheme)
        {
            if (form == null) return;

            var theme = NormalizeTheme(globalTheme);
            if (theme == ThemeDefaultCw)
            {
                ApplyDefaultControlColors(form);
                foreach (var strip in GetToolStrips(form))
                {
                    ApplyDefaultToolStripColors(strip);
                }
                return;
            }

            var palette = (theme == ThemeTssBlue)
                ? CreateBluePalette()
                : CreateRedPalette();

            ApplyDarkControlColors(form, palette);
            foreach (var strip in GetToolStrips(form))
            {
                ApplyDarkToolStripColors(strip, palette);
            }
        }

        private static ThemePalette CreateRedPalette()
        {
            return new ThemePalette
            {
                FormBack = Color.FromArgb(22, 22, 24),
                PanelBack = Color.FromArgb(32, 32, 36),
                InputBack = Color.FromArgb(42, 42, 48),
                Text = Color.FromArgb(236, 236, 236),
                Accent = Color.FromArgb(211, 47, 47),
                StatusBack = Color.FromArgb(28, 28, 32),
                TabSelected = Color.FromArgb(46, 46, 52),
                TabBorder = Color.FromArgb(70, 70, 80)
            };
        }

        private static ThemePalette CreateBluePalette()
        {
            return new ThemePalette
            {
                FormBack = Color.FromArgb(22, 22, 24),
                PanelBack = Color.FromArgb(32, 32, 36),
                InputBack = Color.FromArgb(42, 42, 48),
                Text = Color.FromArgb(236, 236, 236),
                Accent = Color.FromArgb(0, 120, 215),
                StatusBack = Color.FromArgb(28, 28, 32),
                TabSelected = Color.FromArgb(46, 46, 52),
                TabBorder = Color.FromArgb(70, 70, 80)
            };
        }

        private static void ApplyDarkControlColors(Control control, ThemePalette palette)
        {
            if (control == null) return;

            if (control is Form || control is TabPage)
            {
                control.BackColor = palette.FormBack;
                control.ForeColor = palette.Text;
            }
            if (control is TabPage tabPage)
            {
                tabPage.UseVisualStyleBackColor = false;
            }
            else if (control is Panel || control is GroupBox || control is TabControl || control is SplitContainer)
            {
                control.BackColor = palette.PanelBack;
                control.ForeColor = palette.Text;
            }
            else if (control is TextBox || control is RichTextBox || control is ListView || control is TreeView || control is NumericUpDown)
            {
                control.BackColor = palette.InputBack;
                control.ForeColor = palette.Text;
            }
            else if (control is ComboBox)
            {
                control.BackColor = palette.InputBack;
                control.ForeColor = palette.Text;
            }
            else if (control is Button button)
            {
                button.BackColor = palette.PanelBack;
                button.ForeColor = palette.Text;
                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderColor = palette.Accent;
            }
            else if (control is CheckBox || control is RadioButton || control is Label)
            {
                control.ForeColor = palette.Text;
            }
            else if (control is TrackBar)
            {
                control.BackColor = palette.PanelBack;
                control.ForeColor = palette.Text;
            }

            if (control is PropertyGrid propertyGrid)
            {
                propertyGrid.BackColor = palette.PanelBack;
                propertyGrid.ViewBackColor = palette.InputBack;
                propertyGrid.ViewForeColor = palette.Text;
                propertyGrid.HelpBackColor = palette.PanelBack;
                propertyGrid.HelpForeColor = palette.Text;
                propertyGrid.LineColor = palette.TabBorder;
                propertyGrid.CategoryForeColor = palette.Text;
                propertyGrid.CommandsBackColor = palette.PanelBack;
                propertyGrid.CommandsForeColor = palette.Text;
            }

            if (control is CheckBox checkBox)
            {
                checkBox.FlatStyle = FlatStyle.Flat;
                checkBox.BackColor = palette.PanelBack;
            }
            if (control is RadioButton radioButton)
            {
                radioButton.FlatStyle = FlatStyle.Flat;
                radioButton.BackColor = palette.PanelBack;
            }

            if (control is ComboBox comboBox)
            {
                comboBox.FlatStyle = FlatStyle.Flat;
                comboBox.DrawMode = DrawMode.OwnerDrawFixed;
                comboBox.DrawItem -= ThemedComboBox_DrawItem;
                comboBox.DrawItem += ThemedComboBox_DrawItem;
            }

            if (control is TabControl tabControl)
            {
                tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
                tabControl.DrawItem -= ThemedTabControl_DrawItem;
                tabControl.DrawItem += ThemedTabControl_DrawItem;
            }

            if (control is StatusStrip statusStrip)
            {
                statusStrip.BackColor = palette.StatusBack;
                statusStrip.ForeColor = palette.Text;
            }

            // WinForms leaves some surfaces in system colors unless explicitly
            // touched; keep them coherent with the active dark palette.
            if (!(control is PictureBox) &&
                !(control is ToolStrip) &&
                !(control is MenuStrip) &&
                !(control is StatusStrip))
            {
                if (control.BackColor == SystemColors.Control ||
                    control.BackColor == SystemColors.Window ||
                    control.BackColor == Color.White)
                {
                    control.BackColor = palette.PanelBack;
                }
            }

            foreach (Control child in control.Controls)
            {
                ApplyDarkControlColors(child, palette);
            }
        }

        private static void ApplyDefaultControlColors(Control control)
        {
            if (control == null) return;

            if (control is Form)
            {
                control.BackColor = SystemColors.Control;
                control.ForeColor = SystemColors.ControlText;
            }
            else if (control is Panel || control is GroupBox || control is TabControl || control is TabPage || control is SplitContainer)
            {
                control.BackColor = SystemColors.Control;
                control.ForeColor = SystemColors.ControlText;
            }
            else if (control is TextBox || control is RichTextBox || control is ComboBox || control is ListView || control is TreeView || control is NumericUpDown)
            {
                control.BackColor = SystemColors.Window;
                control.ForeColor = SystemColors.WindowText;
            }
            else if (control is Button button)
            {
                button.BackColor = SystemColors.Control;
                button.ForeColor = SystemColors.ControlText;
                button.FlatStyle = FlatStyle.Standard;
                button.UseVisualStyleBackColor = true;
            }
            else if (control is CheckBox || control is RadioButton || control is Label)
            {
                control.ForeColor = SystemColors.ControlText;
            }
            else if (control is TrackBar)
            {
                control.BackColor = SystemColors.Control;
                control.ForeColor = SystemColors.ControlText;
            }

            if (control is TabPage defaultTabPage)
            {
                defaultTabPage.UseVisualStyleBackColor = true;
            }

            if (control is PropertyGrid propertyGrid)
            {
                propertyGrid.BackColor = SystemColors.Control;
                propertyGrid.ViewBackColor = SystemColors.Window;
                propertyGrid.ViewForeColor = SystemColors.WindowText;
                propertyGrid.HelpBackColor = SystemColors.Control;
                propertyGrid.HelpForeColor = SystemColors.ControlText;
                propertyGrid.LineColor = SystemColors.ControlDark;
                propertyGrid.CategoryForeColor = SystemColors.ControlText;
                propertyGrid.CommandsBackColor = SystemColors.Control;
                propertyGrid.CommandsForeColor = SystemColors.ControlText;
            }

            if (control is CheckBox checkBox)
            {
                checkBox.FlatStyle = FlatStyle.Standard;
                checkBox.BackColor = SystemColors.Control;
            }
            if (control is RadioButton radioButton)
            {
                radioButton.FlatStyle = FlatStyle.Standard;
                radioButton.BackColor = SystemColors.Control;
            }

            if (control is ComboBox comboBox)
            {
                comboBox.DrawItem -= ThemedComboBox_DrawItem;
                comboBox.DrawMode = DrawMode.Normal;
                comboBox.FlatStyle = FlatStyle.Standard;
            }

            if (control is TabControl tabControl)
            {
                tabControl.DrawItem -= ThemedTabControl_DrawItem;
                tabControl.DrawMode = TabDrawMode.Normal;
            }

            if (control is StatusStrip statusStrip)
            {
                statusStrip.BackColor = SystemColors.Control;
                statusStrip.ForeColor = SystemColors.ControlText;
            }

            foreach (Control child in control.Controls)
            {
                ApplyDefaultControlColors(child);
            }
        }

        private static IEnumerable<ToolStrip> GetToolStrips(Control root)
        {
            var strips = new List<ToolStrip>();
            CollectToolStrips(root, strips);
            return strips;
        }

        private static void CollectToolStrips(Control control, List<ToolStrip> strips)
        {
            if (control == null) return;

            if (control is ToolStrip strip)
            {
                strips.Add(strip);
            }

            foreach (Control child in control.Controls)
            {
                CollectToolStrips(child, strips);
            }

            if (control.ContextMenuStrip != null)
            {
                strips.Add(control.ContextMenuStrip);
            }
        }

        private static void ApplyDarkToolStripColors(ToolStrip strip, ThemePalette palette)
        {
            if (strip == null) return;

            strip.BackColor = palette.PanelBack;
            strip.ForeColor = palette.Text;
            foreach (ToolStripItem item in strip.Items)
            {
                item.BackColor = palette.PanelBack;
                item.ForeColor = palette.Text;
            }
        }

        private static void ApplyDefaultToolStripColors(ToolStrip strip)
        {
            if (strip == null) return;

            strip.BackColor = SystemColors.Control;
            strip.ForeColor = SystemColors.ControlText;
            foreach (ToolStripItem item in strip.Items)
            {
                item.BackColor = SystemColors.Control;
                item.ForeColor = SystemColors.ControlText;
            }
        }

        private static ThemePalette GetCurrentPalette()
        {
            return NormalizeTheme(Settings.Default.GlobalUITheme) == ThemeTssBlue ? CreateBluePalette() : CreateRedPalette();
        }

        private static void ThemedTabControl_DrawItem(object sender, DrawItemEventArgs e)
        {
            var tabControl = sender as TabControl;
            if (tabControl == null) return;
            if (e.Index < 0 || e.Index >= tabControl.TabPages.Count) return;

            var palette = GetCurrentPalette();
            var selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            var tabRect = tabControl.GetTabRect(e.Index);
            var backColor = selected ? palette.TabSelected : palette.PanelBack;

            using (var backBrush = new SolidBrush(backColor))
            using (var borderPen = new Pen(selected ? palette.Accent : palette.TabBorder))
            using (var textBrush = new SolidBrush(palette.Text))
            {
                e.Graphics.FillRectangle(backBrush, tabRect);
                e.Graphics.DrawRectangle(borderPen, tabRect);

                var text = tabControl.TabPages[e.Index].Text;
                TextRenderer.DrawText(
                    e.Graphics,
                    text,
                    tabControl.Font,
                    tabRect,
                    textBrush.Color,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
        }

        private static void ThemedComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            var comboBox = sender as ComboBox;
            if (comboBox == null) return;
            if (e.Index < 0) return;

            var palette = GetCurrentPalette();
            var selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            var backColor = selected ? palette.TabSelected : palette.InputBack;

            using (var backBrush = new SolidBrush(backColor))
            using (var textBrush = new SolidBrush(palette.Text))
            {
                e.Graphics.FillRectangle(backBrush, e.Bounds);
                var text = comboBox.GetItemText(comboBox.Items[e.Index]);
                TextRenderer.DrawText(
                    e.Graphics,
                    text,
                    comboBox.Font,
                    e.Bounds,
                    textBrush.Color,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            }

            e.DrawFocusRectangle();
        }
    }
}
