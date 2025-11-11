// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Help
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;


namespace MOS_PPT_LEARN
{
    public class Help : Form
    {
        private IContainer components = (IContainer)null;
        private RichTextBox richTextBox1;
        private Button button1;

        public Help()
        {
            this.InitializeComponent();
            this.SizeChanged += new EventHandler(this.Help_SizeChanged);
        }

        private void Help_SizeChanged(object sender, EventArgs e)
        {
            this.richTextBox1.Width = this.Width - 15;
            this.richTextBox1.Height = this.Height - 2 * this.button1.Height;
        }

        private void Help_Load(object sender, EventArgs e)
        {
            this.Width = Screen.PrimaryScreen.Bounds.Width / 2;
            this.Height = Screen.PrimaryScreen.Bounds.Height / 3;
            this.richTextBox1.Location = new Point(0, 0);
            this.richTextBox1.Width = this.Width - 15;
            this.richTextBox1.Height = this.Height - 2 * this.button1.Height;
            Rectangle bounds = Screen.PrimaryScreen.Bounds;
            int x = bounds.Width - this.Width;
            bounds = Screen.PrimaryScreen.Bounds;
            int y = bounds.Height - this.Height;
            this.Location = new Point(x, y);
            string path = Path.Combine(Application.StartupPath, "tam\\help");
            if (!File.Exists(path))
                return;
            this.richTextBox1.LoadFile(path);
        }

        public void refresh()
        {
            string path = Path.Combine(Application.StartupPath, "tam\\help");
            if (!File.Exists(path))
                return;
            this.richTextBox1.LoadFile(path);
        }

        private void button1_Click(object sender, EventArgs e) => this.Hide();

        protected override void Dispose(bool disposing)
        {
            if (disposing && this.components != null)
                this.components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.richTextBox1 = new RichTextBox();
            this.button1 = new Button();
            this.SuspendLayout();
            this.richTextBox1.Location = new Point(61, 49);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new Size(277, 247);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = "";
            this.button1.Font = new Font("Tahoma", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.button1.Location = new Point(0, -3);
            this.button1.Name = "button1";
            this.button1.Size = new Size(40, 33);
            this.button1.TabIndex = 1;
            this.button1.Text = "Ẩn";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new EventHandler(this.button1_Click);
            this.AutoScaleDimensions = new SizeF(6f, 13f);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new Size(426, 365);
            this.ControlBox = false;
            this.Controls.Add((Control)this.button1);
            this.Controls.Add((Control)this.richTextBox1);
            this.Name = nameof(Help);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = nameof(Help);
            this.TopMost = true;
            this.Load += new EventHandler(this.Help_Load);
            this.ResumeLayout(false);
        }
    }
}