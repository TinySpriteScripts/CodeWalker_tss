using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CodeWalker.Tools
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
            ApplyBranding();
        }

        private void ApplyBranding()
        {
            Text = "About CodeWalker x TSS";
            label2.Text = "CodeWalker x TSS";
            label2.ForeColor = Color.FromArgb(220, 45, 45);

            var logoPath = PathUtil.GetFilePath("icons\\tss_red_128.png");
            if (File.Exists(logoPath))
            {
                var logo = new PictureBox
                {
                    Left = (ClientSize.Width - 128) / 2,
                    Top = 8,
                    Width = 128,
                    Height = 32,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    ImageLocation = logoPath
                };
                Controls.Add(logo);
                logo.BringToFront();

                label2.Top = 44;
                label1.Top = 66;
                label1.Height = 108;
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
