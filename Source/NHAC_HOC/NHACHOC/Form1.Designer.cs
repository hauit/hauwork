using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using NHACHOC.Properties;

namespace NHACHOC
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
            this.components = (IContainer)new System.ComponentModel.Container();
            ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(Form1));
            this.timer1 = new Timer(this.components);
            this.textBox2 = new TextBox();
            this.textBox1 = new TextBox();
            this.pictureBox1 = new PictureBox();
            ((ISupportInitialize)this.pictureBox1).BeginInit();
            this.SuspendLayout();
            this.timer1.Interval = 30000;
            this.timer1.Tick += new EventHandler(this.timer1_Tick);
            this.textBox2.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Regular, GraphicsUnit.Point, (byte)163);
            this.textBox2.Location = new Point(318, 221);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.ScrollBars = ScrollBars.Both;
            this.textBox2.Size = new Size(673, 194);
            this.textBox2.TabIndex = 1;
            this.textBox2.Text = "Thường Xuyên Ôn bài để kịp thi nha các bạn cố gắn mỗi môn 3 lần trên 950, IIG sẽ không chấp nhận lý do học không kịp nha các bạn";
            this.textBox1.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Bold, GraphicsUnit.Point, (byte)163);
            this.textBox1.ForeColor = Color.Red;
            this.textBox1.Location = new Point(12, 21);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new Size(979, 194);
            this.textBox1.TabIndex = 2;
            this.pictureBox1.Image = (Image)Resources._2022_08_25;
            this.pictureBox1.Location = new Point(12, 221);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new Size(300, 194);
            this.pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            this.AutoScaleDimensions = new SizeF(6f, 13f);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1021, 448);
            this.Controls.Add((Control)this.pictureBox1);
            this.Controls.Add((Control)this.textBox1);
            this.Controls.Add((Control)this.textBox2);
            //this.Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
            this.Name = "Form1";
            this.Text = "NHẮC HỌC HÀNH";
            this.Load += new EventHandler(this.Form1_Load);
            ((ISupportInitialize)this.pictureBox1).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Timer timer1;
        private TextBox textBox2;
        private TextBox textBox1;
        private PictureBox pictureBox1;
    }
}