using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace MOS_EXCEL_LEARN
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
            this.buttonExit = new Button();
            this.buttonNext = new Button();
            this.buttonReset = new Button();
            this.buttonEV = new Button();
            this.buttonHelp = new Button();
            this.buttonRefresh = new Button();
            this.richTextBox1 = new RichTextBox();
            this.buttonCheck = new Button();
            this.pictureBox1 = new PictureBox();
            this.labelKQ = new Label();
            this.labelCauHienTai = new Label();
            this.comboBoxCauNext = new ComboBox();
            this.panel1 = new Panel();
            this.buttonZoom = new Button();
            this.buttonSummary = new Button();
            this.buttonVideoHelp = new Button();
            this.button2 = new Button();
            this.comboBoxCauDaDanhDau = new ComboBox();
            this.buttonHelpVideo = new Button();
            ((ISupportInitialize)this.pictureBox1).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            this.buttonExit.BackColor = Color.Transparent;
            this.buttonExit.Location = new Point(100, 68);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new Size(75, 23);
            this.buttonExit.TabIndex = 0;
            this.buttonExit.Text = "Exit";
            this.buttonExit.UseVisualStyleBackColor = false;
            this.buttonExit.Click += new EventHandler(this.button1_Click);
            this.buttonNext.Location = new Point(216, 68);
            this.buttonNext.Name = "buttonNext";
            this.buttonNext.Size = new Size(75, 23);
            this.buttonNext.TabIndex = 1;
            this.buttonNext.Text = "Next";
            this.buttonNext.UseVisualStyleBackColor = true;
            this.buttonNext.Click += new EventHandler(this.button2_Click);
            this.buttonReset.Location = new Point(119, 126);
            this.buttonReset.Name = "buttonReset";
            this.buttonReset.Size = new Size(75, 23);
            this.buttonReset.TabIndex = 2;
            this.buttonReset.Text = "Reset";
            this.buttonReset.UseVisualStyleBackColor = true;
            this.buttonReset.Click += new EventHandler(this.buttonReset_Click);
            this.buttonEV.Location = new Point(235, 130);
            this.buttonEV.Name = "buttonEV";
            this.buttonEV.Size = new Size(75, 23);
            this.buttonEV.TabIndex = 3;
            this.buttonEV.Text = "Tiếng Việt";
            this.buttonEV.UseVisualStyleBackColor = true;
            this.buttonEV.Click += new EventHandler(this.buttonEV_Click);
            this.buttonHelp.Location = new Point(365, 234);
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Size = new Size(114, 23);
            this.buttonHelp.TabIndex = 4;
            this.buttonHelp.Text = "Về chế độ thường";
            this.buttonHelp.UseVisualStyleBackColor = true;
            this.buttonHelp.Click += new EventHandler(this.buttonHelp_Click);
            this.buttonRefresh.Location = new Point(365, 263);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new Size(87, 23);
            this.buttonRefresh.TabIndex = 5;
            this.buttonRefresh.Text = "Phục hồi Excel";
            this.buttonRefresh.UseVisualStyleBackColor = true;
            this.buttonRefresh.Click += new EventHandler(this.buttonRefresh_Click);
            this.richTextBox1.Font = new System.Drawing.Font("Times New Roman", 18f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.richTextBox1.Location = new Point(61, 296);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new Size(103, 82);
            this.richTextBox1.TabIndex = 6;
            this.richTextBox1.Text = "";
            this.buttonCheck.BackColor = Color.FromArgb((int)byte.MaxValue, 192, 128);
            this.buttonCheck.Location = new Point(285, 184);
            this.buttonCheck.Name = "buttonCheck";
            this.buttonCheck.Size = new Size(97, 23);
            this.buttonCheck.TabIndex = 7;
            this.buttonCheck.Text = "Check đúng/sai";
            this.buttonCheck.UseVisualStyleBackColor = false;
            this.buttonCheck.Click += new EventHandler(this.buttonCheck_Click);
            this.pictureBox1.Location = new Point(37, 33);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new Size(171, 218);
            this.pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 8;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new EventHandler(this.pictureBox1_Click);
            this.pictureBox1.MouseHover += new EventHandler(this.pictureBox1_MouseHover);
            this.labelKQ.AutoSize = true;
            this.labelKQ.Font = new System.Drawing.Font("Times New Roman", 14.25f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.labelKQ.ForeColor = SystemColors.Highlight;
            this.labelKQ.Location = new Point(516, 220);
            this.labelKQ.Name = "labelKQ";
            this.labelKQ.Size = new Size(0, 21);
            this.labelKQ.TabIndex = 10;
            this.labelCauHienTai.AutoSize = true;
            this.labelCauHienTai.Font = new System.Drawing.Font("Times New Roman", 14.25f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.labelCauHienTai.ForeColor = Color.DeepPink;
            this.labelCauHienTai.Location = new Point(569, 274);
            this.labelCauHienTai.Name = "labelCauHienTai";
            this.labelCauHienTai.Size = new Size(72, 21);
            this.labelCauHienTai.TabIndex = 12;
            this.labelCauHienTai.Text = "Câu 150";
            this.comboBoxCauNext.BackColor = SystemColors.MenuText;
            this.comboBoxCauNext.Font = new System.Drawing.Font("Times New Roman", 14.25f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.comboBoxCauNext.ForeColor = Color.Coral;
            this.comboBoxCauNext.FormattingEnabled = true;
            this.comboBoxCauNext.Location = new Point(720, 369);
            this.comboBoxCauNext.Name = "comboBoxCauNext";
            this.comboBoxCauNext.Size = new Size(52, 29);
            this.comboBoxCauNext.TabIndex = 13;
            this.comboBoxCauNext.Text = "1";
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add((Control)this.pictureBox1);
            this.panel1.Location = new Point(573, 35);
            this.panel1.Name = "panel1";
            this.panel1.Size = new Size(282, 172);
            this.panel1.TabIndex = 15;
            this.buttonZoom.BackColor = Color.FromArgb(192, (int)byte.MaxValue, (int)byte.MaxValue);
            this.buttonZoom.Location = new Point(487, 288);
            this.buttonZoom.Name = "buttonZoom";
            this.buttonZoom.Size = new Size(76, 23);
            this.buttonZoom.TabIndex = 21;
            this.buttonZoom.Text = "Đánh dấu";
            this.buttonZoom.UseVisualStyleBackColor = false;
            this.buttonZoom.Click += new EventHandler(this.buttonZoom_Click);
            this.buttonSummary.Location = new Point(285, 263);
            this.buttonSummary.Name = "buttonSummary";
            this.buttonSummary.Size = new Size(42, 23);
            this.buttonSummary.TabIndex = 23;
            this.buttonSummary.Text = "Xóa";
            this.buttonSummary.UseVisualStyleBackColor = true;
            this.buttonSummary.Click += new EventHandler(this.buttonSummary_Click);
            this.buttonVideoHelp.BackColor = Color.FromArgb((int)byte.MaxValue, (int)byte.MaxValue, 192);
            this.buttonVideoHelp.Location = new Point(545, 218);
            this.buttonVideoHelp.Name = "buttonVideoHelp";
            this.buttonVideoHelp.Size = new Size(108, 23);
            this.buttonVideoHelp.TabIndex = 24;
            this.buttonVideoHelp.Text = "HDSD phần mềm";
            this.buttonVideoHelp.UseVisualStyleBackColor = false;
            this.buttonVideoHelp.Click += new EventHandler(this.buttonVideoHelp_Click);
            this.button2.BackColor = Color.Transparent;
            this.button2.ForeColor = Color.Black;
            this.button2.Location = new Point(389, 374);
            this.button2.Name = "button2";
            this.button2.Size = new Size(77, 23);
            this.button2.TabIndex = 25;
            this.button2.Text = "Gọi Thầy";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new EventHandler(this.button2_Click_1);
            this.comboBoxCauDaDanhDau.BackColor = Color.White;
            this.comboBoxCauDaDanhDau.Font = new System.Drawing.Font("Times New Roman", 12f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.comboBoxCauDaDanhDau.ForeColor = Color.Red;
            this.comboBoxCauDaDanhDau.FormattingEnabled = true;
            this.comboBoxCauDaDanhDau.Location = new Point(835, 298);
            this.comboBoxCauDaDanhDau.Name = "comboBoxCauDaDanhDau";
            this.comboBoxCauDaDanhDau.Size = new Size(43, 27);
            this.comboBoxCauDaDanhDau.TabIndex = 26;
            this.buttonHelpVideo.BackColor = Color.FromArgb((int)byte.MaxValue, 192, 192);
            this.buttonHelpVideo.Location = new Point(365, 333);
            this.buttonHelpVideo.Name = "buttonHelpVideo";
            this.buttonHelpVideo.Size = new Size(75, 23);
            this.buttonHelpVideo.TabIndex = 30;
            this.buttonHelpVideo.Text = "Chỉ làm bài";
            this.buttonHelpVideo.UseVisualStyleBackColor = false;
            this.buttonHelpVideo.Click += new EventHandler(this.buttonHelpVideo_Click);
            // 
            // Form1_New
            // 
            this.AutoScaleDimensions = new SizeF(6f, 13f);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.BackColor = Color.Black;
            this.ClientSize = new Size(1016, 599);
            this.Controls.Add((Control)this.buttonHelpVideo);
            this.Controls.Add((Control)this.labelKQ);
            this.Controls.Add((Control)this.comboBoxCauDaDanhDau);
            this.Controls.Add((Control)this.button2);
            this.Controls.Add((Control)this.buttonZoom);
            this.Controls.Add((Control)this.panel1);
            this.Controls.Add((Control)this.comboBoxCauNext);
            this.Controls.Add((Control)this.labelCauHienTai);
            this.Controls.Add((Control)this.buttonCheck);
            this.Controls.Add((Control)this.richTextBox1);
            this.Controls.Add((Control)this.buttonRefresh);
            this.Controls.Add((Control)this.buttonHelp);
            this.Controls.Add((Control)this.buttonEV);
            this.Controls.Add((Control)this.buttonReset);
            this.Controls.Add((Control)this.buttonNext);
            this.Controls.Add((Control)this.buttonExit);
            this.Controls.Add((Control)this.buttonSummary);
            this.Controls.Add((Control)this.buttonVideoHelp);

            this.FormBorderStyle = FormBorderStyle.None;
            this.Name = "Form1_New";
            this.Text = "Form1_New";
            this.TopMost = true;
            this.WindowState = FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_New_Load);
            ((ISupportInitialize)this.pictureBox1).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();


            //this.FormBorderStyle = FormBorderStyle.None;
            //this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            //this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            //this.ClientSize = new System.Drawing.Size(800, 450);
            //this.Name = "Form1_New";
            //this.Text = "Form1_New";
            //this.Load += new System.EventHandler(this.Form1_New_Load);
            //this.ResumeLayout(false);

        }

        #endregion
        private Button buttonExit;
        private Button buttonNext;
        private Button buttonReset;
        private Button buttonEV;
        private Button buttonHelp;
        private Button buttonRefresh;
        private RichTextBox richTextBox1;
        private Button buttonCheck;
        private PictureBox pictureBox1;
        private Label labelKQ;
        private Label labelCauHienTai;
        private ComboBox comboBoxCauNext;
        private Panel panel1;
        private Button buttonZoom;
        private Button buttonSummary;
        private Button buttonVideoHelp;
        private Button button2;
        private ComboBox comboBoxCauDaDanhDau;
        private Button buttonHelpVideo;
    }
}