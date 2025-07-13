using System;
using System.Drawing;
using System.Windows.Forms;

namespace MOS_EXCEL_LEARN
{
    partial class Form4
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
            this.buttonNext.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.buttonNext.Location = new Point(433, 618);
            this.buttonNext.Name = "buttonNext";
            this.buttonNext.Size = new Size(489, 72);
            this.buttonNext.TabIndex = 6;
            this.buttonNext.Text = "Next";
            this.buttonNext.UseVisualStyleBackColor = true;
            this.buttonNext.Click += new EventHandler(this.buttonNext_Click);
            // 
            // Form4_New
            // 
            this.BackColor = Color.Black;
            this.ClientSize = new Size(1299, 702);
            this.Controls.Add((Control)this.richTextBox1);
            this.Controls.Add((Control)this.buttonNext);
            this.Controls.Add((Control)this.buttonExit);
            this.FormBorderStyle = FormBorderStyle.None;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.WindowState = FormWindowState.Maximized;
            this.Name = "Form4_New";
            this.Load += new System.EventHandler(this.Form4_New_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private Button buttonExit;
        private Button buttonNext;
        private RichTextBox richTextBox1;
    }
}