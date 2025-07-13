using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace MOS_EXCEL_TEST
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
            this.checkedListBox1 = new CheckedListBox();
            this.buttonReset = new System.Windows.Forms.Button();
            this.buttonCheck = new System.Windows.Forms.Button();
            this.buttonSubmit = new System.Windows.Forms.Button();
            this.buttonExit = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.timer1 = new Timer(this.components);
            this.tabControl1 = new TabControl();
            this.tabPage2 = new TabPage();
            this.panel2 = new Panel();
            this.pictureBox2 = new PictureBox();
            this.tabPage1 = new TabPage();
            this.panel1 = new Panel();
            this.pictureBox1 = new PictureBox();
            this.buttonSave = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.panel4 = new Panel();
            this.richTextBox1 = new RichTextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.buttonHelp = new System.Windows.Forms.Button();
            this.buttonShowHide = new System.Windows.Forms.Button();
            this.buttonxhdh = new System.Windows.Forms.Button();
            this.buttonRefresh = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.comboBox1 = new ComboBox();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panel2.SuspendLayout();
            ((ISupportInitialize)this.pictureBox2).BeginInit();
            this.tabPage1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((ISupportInitialize)this.pictureBox1).BeginInit();
            this.panel4.SuspendLayout();
            this.SuspendLayout();
            this.checkedListBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(29, 30);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new Size(61, 172);
            this.checkedListBox1.TabIndex = 0;
            this.checkedListBox1.SelectedIndexChanged += new EventHandler(this.checkedListBox1_SelectedIndexChanged);
            this.buttonReset.Location = new System.Drawing.Point(70, 266);
            this.buttonReset.Name = "buttonReset";
            this.buttonReset.Size = new Size(75, 23);
            this.buttonReset.TabIndex = 2;
            this.buttonReset.Text = "Reset";
            this.buttonReset.UseVisualStyleBackColor = true;
            this.buttonReset.Click += new EventHandler(this.button1_Click);
            this.buttonCheck.Location = new System.Drawing.Point(70, 295);
            this.buttonCheck.Name = "buttonCheck";
            this.buttonCheck.Size = new Size(75, 23);
            this.buttonCheck.TabIndex = 3;
            this.buttonCheck.Text = "&Check";
            this.buttonCheck.UseVisualStyleBackColor = true;
            this.buttonCheck.Click += new EventHandler(this.button2_Click);
            this.buttonSubmit.Location = new System.Drawing.Point(249, 372);
            this.buttonSubmit.Name = "buttonSubmit";
            this.buttonSubmit.Size = new Size(75, 23);
            this.buttonSubmit.TabIndex = 4;
            this.buttonSubmit.Text = "Submit";
            this.buttonSubmit.UseVisualStyleBackColor = true;
            this.buttonSubmit.Click += new EventHandler(this.button3_Click);
            this.buttonExit.Location = new System.Drawing.Point(413, 393);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new Size(75, 23);
            this.buttonExit.TabIndex = 5;
            this.buttonExit.Text = "E&xit";
            this.buttonExit.UseVisualStyleBackColor = true;
            this.buttonExit.Click += new EventHandler(this.button4_Click);
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.label1.Location = new System.Drawing.Point(514, 347);
            this.label1.Name = "label1";
            this.label1.Size = new Size(100, 25);
            this.label1.TabIndex = 6;
            this.label1.Text = "00:00";
            this.label1.TextAlign = ContentAlignment.MiddleCenter;
            this.timer1.Tick += new EventHandler(this.timer1_Tick);
            this.tabControl1.Controls.Add((Control)this.tabPage2);
            this.tabControl1.Controls.Add((Control)this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(164, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new Size(529, 309);
            this.tabControl1.TabIndex = 8;
            this.tabPage2.Controls.Add((Control)this.panel2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new Padding(3);
            this.tabPage2.Size = new Size(521, 283);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "T A";
            this.tabPage2.UseVisualStyleBackColor = true;
            this.panel2.AutoScroll = true;
            this.panel2.Controls.Add((Control)this.pictureBox2);
            this.panel2.Location = new System.Drawing.Point(3, 3);
            this.panel2.Name = "panel2";
            this.panel2.Size = new Size(512, 264);
            this.panel2.TabIndex = 0;
            this.pictureBox2.Location = new System.Drawing.Point(3, 3);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new Size(1028, 3615);
            this.pictureBox2.SizeMode = PictureBoxSizeMode.AutoSize;
            this.pictureBox2.TabIndex = 0;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.MouseHover += new EventHandler(this.pictureBox2_MouseHover);
            this.tabPage1.Controls.Add((Control)this.panel1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new Size(521, 283);
            this.tabPage1.TabIndex = 2;
            this.tabPage1.Text = "TV";
            this.tabPage1.UseVisualStyleBackColor = true;
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add((Control)this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(70, 68);
            this.panel1.Name = "panel1";
            this.panel1.Size = new Size(389, 176);
            this.panel1.TabIndex = 2;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new Size(100, 50);
            this.pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.MouseHover += new EventHandler(this.pictureBox1_MouseHover);
            this.buttonSave.Location = new System.Drawing.Point(114, 372);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new Size(83, 23);
            this.buttonSave.TabIndex = 16;
            this.buttonSave.Text = "Lưu lai hỏi sau";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new EventHandler(this.buttonSave_Click);
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(699, 193);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new Size(100, 20);
            this.textBox1.TabIndex = 15;
            this.textBox1.Visible = false;
            this.panel4.AutoScroll = true;
            this.panel4.Controls.Add((Control)this.richTextBox1);
            this.panel4.Location = new System.Drawing.Point(713, 307);
            this.panel4.Name = "panel4";
            this.panel4.Size = new Size(98, 88);
            this.panel4.TabIndex = 17;
            this.richTextBox1.BorderStyle = BorderStyle.None;
            this.richTextBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.richTextBox1.ForeColor = Color.Red;
            this.richTextBox1.Location = new System.Drawing.Point(3, 3);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new Size(83, 75);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = "";
            this.button1.Location = new System.Drawing.Point(629, 393);
            this.button1.Name = "button1";
            this.button1.Size = new Size(29, 23);
            this.button1.TabIndex = 24;
            this.button1.Text = "Go";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new EventHandler(this.button1_Click_1);
            this.textBox2.Location = new System.Drawing.Point(601, 396);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new Size(22, 20);
            this.textBox2.TabIndex = 23;
            this.textBox2.Text = "1";
            this.buttonHelp.Location = new System.Drawing.Point(343, 351);
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Size = new Size(111, 23);
            this.buttonHelp.TabIndex = 25;
            this.buttonHelp.Text = "Video chỉ làm bài";
            this.buttonHelp.UseVisualStyleBackColor = true;
            this.buttonHelp.Click += new EventHandler(this.buttonHelp_Click);
            this.buttonShowHide.Location = new System.Drawing.Point(712, 260);
            this.buttonShowHide.Name = "buttonShowHide";
            this.buttonShowHide.Size = new Size(66, 23);
            this.buttonShowHide.TabIndex = 26;
            this.buttonShowHide.Text = "Gọi Thầy";
            this.buttonShowHide.UseVisualStyleBackColor = true;
            this.buttonShowHide.Click += new EventHandler(this.button2_Click_1);
            this.buttonxhdh.Location = new System.Drawing.Point(114, 429);
            this.buttonxhdh.Name = "buttonxhdh";
            this.buttonxhdh.Size = new Size(114, 23);
            this.buttonxhdh.TabIndex = 28;
            this.buttonxhdh.Text = "Hướng dẫn sử dụng";
            this.buttonxhdh.UseVisualStyleBackColor = true;
            this.buttonxhdh.Click += new EventHandler(this.buttonxhdh_Click);
            this.buttonRefresh.Location = new System.Drawing.Point(29, 343);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new Size(61, 23);
            this.buttonRefresh.TabIndex = 27;
            this.buttonRefresh.Text = "Bố tri lại";
            this.buttonRefresh.UseVisualStyleBackColor = true;
            this.buttonRefresh.Click += new EventHandler(this.buttonRefresh_Click);
            this.button2.Location = new System.Drawing.Point(328, 429);
            this.button2.Name = "button2";
            this.button2.Size = new Size(41, 23);
            this.button2.TabIndex = 29;
            this.button2.Text = "Mở";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new EventHandler(this.button2_Click_2);
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(544, 431);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new Size(44, 21);
            this.comboBox1.TabIndex = 30;
            this.AutoScaleDimensions = new SizeF(6f, 13f);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(823, 464);
            this.Controls.Add((Control)this.comboBox1);
            this.Controls.Add((Control)this.button2);
            this.Controls.Add((Control)this.buttonShowHide);
            this.Controls.Add((Control)this.buttonxhdh);
            this.Controls.Add((Control)this.buttonRefresh);
            this.Controls.Add((Control)this.buttonHelp);
            this.Controls.Add((Control)this.button1);
            this.Controls.Add((Control)this.textBox2);
            this.Controls.Add((Control)this.panel4);
            this.Controls.Add((Control)this.buttonSave);
            this.Controls.Add((Control)this.textBox1);
            this.Controls.Add((Control)this.label1);
            this.Controls.Add((Control)this.buttonExit);
            this.Controls.Add((Control)this.buttonSubmit);
            this.Controls.Add((Control)this.buttonCheck);
            this.Controls.Add((Control)this.buttonReset);
            this.Controls.Add((Control)this.checkedListBox1);
            this.Controls.Add((Control)this.tabControl1);
            this.Name = "Form1_New";
            this.Text = "Form1_New";
            this.Load += new EventHandler(this.Form1_New_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((ISupportInitialize)this.pictureBox2).EndInit();
            this.tabPage1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((ISupportInitialize)this.pictureBox1).EndInit();
            this.panel4.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private CheckedListBox checkedListBox1;
        private System.Windows.Forms.Button buttonReset;
        private System.Windows.Forms.Button buttonCheck;
        private System.Windows.Forms.Button buttonSubmit;
        private System.Windows.Forms.Button buttonExit;
        private System.Windows.Forms.Label label1;
        private Timer timer1;
        private TabControl tabControl1;
        private TabPage tabPage2;
        private Panel panel2;
        private PictureBox pictureBox2;
        private System.Windows.Forms.Button buttonSave;
        private System.Windows.Forms.TextBox textBox1;
        private Panel panel4;
        private RichTextBox richTextBox1;
        private TabPage tabPage1;
        private PictureBox pictureBox1;
        private Panel panel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button buttonHelp;
        private System.Windows.Forms.Button buttonShowHide;
        private System.Windows.Forms.Button buttonxhdh;
        private System.Windows.Forms.Button buttonRefresh;
        private System.Windows.Forms.Button button2;
        private ComboBox comboBox1;
    }
}