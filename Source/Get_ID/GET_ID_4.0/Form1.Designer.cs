using System;
using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    partial class Form1
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
            this.button1 = new Button();
            this.textBox1 = new TextBox();
            this.button2 = new Button();
            this.SuspendLayout();
            this.button1.Location = new Point(178, 311);
            this.button1.Name = "button1";
            this.button1.Size = new Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "&OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new EventHandler(this.button1_Click);
            this.textBox1.Font = new Font("Microsoft Sans Serif", 24f, FontStyle.Bold, GraphicsUnit.Point, (byte)163);
            this.textBox1.Location = new Point(64, 52);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = ScrollBars.Both;
            this.textBox1.Size = new Size(344, 214);
            this.textBox1.TabIndex = 1;
            this.textBox1.Text = "No your Computer";
            this.button2.Location = new Point(351, 330);
            this.button2.Name = "button2";
            this.button2.Size = new Size(109, 34);
            this.button2.TabIndex = 2;
            this.button2.Text = "Cho phep Chay";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Visible = false;
            this.button2.Click += new EventHandler(this.button2_Click);
            this.AutoScaleDimensions = new SizeF(6f, 13f);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(472, 366);
            this.Controls.Add((Control)this.button2);
            this.Controls.Add((Control)this.textBox1);
            this.Controls.Add((Control)this.button1);
            this.Name = "Form1";
            this.Text = "Get ID";
            this.Load += new EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private Button button1;
        private TextBox textBox1;
        private Button button2;
    }
}