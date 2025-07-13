// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.Form4
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace MOS_EXCEL_TEST
{
  public class Form4_cu : Form
  {
    private IContainer components = (IContainer) null;
    private Button buttonExit;
    private Button buttonNext;
    private RichTextBox richTextBox1;

    protected override void Dispose(bool disposing)
    {
      if (disposing && this.components != null)
        this.components.Dispose();
      base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
      this.buttonExit = new Button();
      this.richTextBox1 = new RichTextBox();
      this.buttonNext = new Button();
      this.SuspendLayout();
      this.buttonExit.Location = new Point(910, 3);
      this.buttonExit.Name = "buttonExit";
      this.buttonExit.Size = new Size(75, 23);
      this.buttonExit.TabIndex = 0;
      this.buttonExit.Text = "E&xit";
      this.buttonExit.UseVisualStyleBackColor = true;
      this.buttonExit.Click += new EventHandler(this.button1_Click);
      this.richTextBox1.Location = new Point(184, 314);
      this.richTextBox1.Name = "richTextBox1";
      this.richTextBox1.ReadOnly = true;
      this.richTextBox1.Size = new Size(183, 115);
      this.richTextBox1.TabIndex = 0;
      this.richTextBox1.Text = "";
      this.buttonNext.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
      this.buttonNext.Location = new Point(433, 618);
      this.buttonNext.Name = "buttonNext";
      this.buttonNext.Size = new Size(489, 72);
      this.buttonNext.TabIndex = 6;
      this.buttonNext.Text = "Next";
      this.buttonNext.UseVisualStyleBackColor = true;
      this.buttonNext.Click += new EventHandler(this.buttonNext_Click);
      this.BackColor = Color.Black;
      this.ClientSize = new Size(1299, 702);
      this.Controls.Add((Control) this.richTextBox1);
      this.Controls.Add((Control) this.buttonNext);
      this.Controls.Add((Control) this.buttonExit);
      this.FormBorderStyle = FormBorderStyle.None;
      this.Name = nameof (Form4_cu);
      this.WindowState = FormWindowState.Maximized;
      this.Load += new EventHandler(this.Form4_Load);
      this.ResumeLayout(false);
    }

    public Form4_cu() => this.InitializeComponent();

    private void button1_Click(object sender, EventArgs e)
    {
      Program.key = "";
      this.Close();
    }

    private void buttonNext_Click(object sender, EventArgs e)
    {
      Program.key = "Next";
      this.Close();
    }

    private void Form4_Load(object sender, EventArgs e)
    {
      this.buttonExit.Location = new Point(this.Width - this.buttonExit.Width, this.Height - this.buttonExit.Height);
      this.richTextBox1.Width = this.Width;
      this.richTextBox1.Height = this.Height - this.buttonNext.Height;
      this.richTextBox1.Location = new Point(0, 0);
      this.buttonNext.Location = new Point((this.Width - this.buttonNext.Width) / 2, this.Height - this.buttonNext.Height);
      if (!File.Exists("d:\\config\\cs2.rtf"))
        return;
      this.richTextBox1.LoadFile("d:\\config\\cs2.rtf");
    }
  }
}
