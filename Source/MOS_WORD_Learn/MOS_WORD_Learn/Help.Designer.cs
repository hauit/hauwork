using System;
using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    partial class Help
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
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
            this.Name = "Help";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Help";
            this.TopMost = true;
            this.Load += new EventHandler(this.Help_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private RichTextBox richTextBox1;
        private Button button1;
    }
}