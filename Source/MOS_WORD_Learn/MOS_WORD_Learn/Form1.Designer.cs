using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace MOS_WORD_LEARN
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
            this.buttonExit = new System.Windows.Forms.Button();
            this.buttonNext = new System.Windows.Forms.Button();
            this.buttonReset = new System.Windows.Forms.Button();
            this.buttonEV = new System.Windows.Forms.Button();
            this.buttonRefresh = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.buttonCheck = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.labelKQ = new System.Windows.Forms.Label();
            this.labelCauHienTai = new System.Windows.Forms.Label();
            this.comboBoxCauNext = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.buttonSummary = new System.Windows.Forms.Button();
            this.buttonZoom = new System.Windows.Forms.Button();
            this.buttonVideoHelp = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.comboBoxCauDaDanhDau = new System.Windows.Forms.ComboBox();
            this.buttonHelpVideo = new System.Windows.Forms.Button();
            this.buttonHelp = new System.Windows.Forms.Button();
            this.richTextBox2 = new System.Windows.Forms.RichTextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonExit
            // 
            this.buttonExit.Location = new System.Drawing.Point(100, 68);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new System.Drawing.Size(75, 23);
            this.buttonExit.TabIndex = 0;
            this.buttonExit.Text = "Exit";
            this.buttonExit.UseVisualStyleBackColor = true;
            this.buttonExit.Click += new System.EventHandler(this.button1_Click);
            // 
            // buttonNext
            // 
            this.buttonNext.Location = new System.Drawing.Point(216, 68);
            this.buttonNext.Name = "buttonNext";
            this.buttonNext.Size = new System.Drawing.Size(75, 23);
            this.buttonNext.TabIndex = 1;
            this.buttonNext.Text = "Next";
            this.buttonNext.UseVisualStyleBackColor = true;
            this.buttonNext.Click += new System.EventHandler(this.button2_Click);
            // 
            // buttonReset
            // 
            this.buttonReset.Location = new System.Drawing.Point(119, 126);
            this.buttonReset.Name = "buttonReset";
            this.buttonReset.Size = new System.Drawing.Size(75, 23);
            this.buttonReset.TabIndex = 2;
            this.buttonReset.Text = "Reset";
            this.buttonReset.UseVisualStyleBackColor = true;
            this.buttonReset.Click += new System.EventHandler(this.buttonReset_Click);
            // 
            // buttonEV
            // 
            this.buttonEV.BackColor = System.Drawing.Color.Transparent;
            this.buttonEV.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.buttonEV.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonEV.Location = new System.Drawing.Point(235, 130);
            this.buttonEV.Name = "buttonEV";
            this.buttonEV.Size = new System.Drawing.Size(75, 23);
            this.buttonEV.TabIndex = 3;
            this.buttonEV.Text = "Tiếng Việt";
            this.buttonEV.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.buttonEV.UseVisualStyleBackColor = false;
            this.buttonEV.Visible = false;
            this.buttonEV.Click += new System.EventHandler(this.buttonEV_Click);
            // 
            // buttonRefresh
            // 
            this.buttonRefresh.Location = new System.Drawing.Point(216, 216);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new System.Drawing.Size(94, 23);
            this.buttonRefresh.TabIndex = 5;
            this.buttonRefresh.Text = "Phục hồi Word";
            this.buttonRefresh.UseVisualStyleBackColor = true;
            this.buttonRefresh.Click += new System.EventHandler(this.buttonRefresh_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Font = new System.Drawing.Font("Times New Roman", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBox1.Location = new System.Drawing.Point(61, 296);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(103, 82);
            this.richTextBox1.TabIndex = 6;
            this.richTextBox1.Text = "";
            // 
            // buttonCheck
            // 
            this.buttonCheck.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.buttonCheck.Location = new System.Drawing.Point(337, 181);
            this.buttonCheck.Name = "buttonCheck";
            this.buttonCheck.Size = new System.Drawing.Size(105, 23);
            this.buttonCheck.TabIndex = 7;
            this.buttonCheck.Text = "Check đúng/ sai";
            this.buttonCheck.UseVisualStyleBackColor = false;
            this.buttonCheck.Click += new System.EventHandler(this.buttonCheck_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(9, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(289, 84);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 8;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            this.pictureBox1.MouseHover += new System.EventHandler(this.pictureBox1_MouseHover);
            // 
            // labelKQ
            // 
            this.labelKQ.AutoSize = true;
            this.labelKQ.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelKQ.ForeColor = System.Drawing.SystemColors.Highlight;
            this.labelKQ.Location = new System.Drawing.Point(606, 218);
            this.labelKQ.Name = "labelKQ";
            this.labelKQ.Size = new System.Drawing.Size(0, 21);
            this.labelKQ.TabIndex = 10;
            // 
            // labelCauHienTai
            // 
            this.labelCauHienTai.AutoSize = true;
            this.labelCauHienTai.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelCauHienTai.ForeColor = System.Drawing.Color.Magenta;
            this.labelCauHienTai.Location = new System.Drawing.Point(587, 217);
            this.labelCauHienTai.Name = "labelCauHienTai";
            this.labelCauHienTai.Size = new System.Drawing.Size(79, 22);
            this.labelCauHienTai.TabIndex = 12;
            this.labelCauHienTai.Text = "Câu 150";
            // 
            // comboBoxCauNext
            // 
            this.comboBoxCauNext.BackColor = System.Drawing.SystemColors.MenuText;
            this.comboBoxCauNext.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxCauNext.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.comboBoxCauNext.FormattingEnabled = true;
            this.comboBoxCauNext.Location = new System.Drawing.Point(720, 369);
            this.comboBoxCauNext.Name = "comboBoxCauNext";
            this.comboBoxCauNext.Size = new System.Drawing.Size(52, 29);
            this.comboBoxCauNext.TabIndex = 13;
            this.comboBoxCauNext.Text = "1";
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(490, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(282, 172);
            this.panel1.TabIndex = 15;
            // 
            // buttonSummary
            // 
            this.buttonSummary.Location = new System.Drawing.Point(471, 288);
            this.buttonSummary.Name = "buttonSummary";
            this.buttonSummary.Size = new System.Drawing.Size(38, 23);
            this.buttonSummary.TabIndex = 17;
            this.buttonSummary.Text = "Xóa";
            this.buttonSummary.UseVisualStyleBackColor = true;
            this.buttonSummary.Click += new System.EventHandler(this.buttonSummary_Click);
            // 
            // buttonZoom
            // 
            this.buttonZoom.BackColor = System.Drawing.Color.Cyan;
            this.buttonZoom.Location = new System.Drawing.Point(362, 296);
            this.buttonZoom.Name = "buttonZoom";
            this.buttonZoom.Size = new System.Drawing.Size(68, 23);
            this.buttonZoom.TabIndex = 20;
            this.buttonZoom.Text = "Đánh Dấu";
            this.buttonZoom.UseVisualStyleBackColor = false;
            this.buttonZoom.Click += new System.EventHandler(this.buttonZoom_Click);
            // 
            // buttonVideoHelp
            // 
            this.buttonVideoHelp.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.buttonVideoHelp.Location = new System.Drawing.Point(198, 288);
            this.buttonVideoHelp.Name = "buttonVideoHelp";
            this.buttonVideoHelp.Size = new System.Drawing.Size(112, 23);
            this.buttonVideoHelp.TabIndex = 24;
            this.buttonVideoHelp.Text = "HDSD phần mềm";
            this.buttonVideoHelp.UseVisualStyleBackColor = false;
            this.buttonVideoHelp.Click += new System.EventHandler(this.buttonVideoHelp_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(410, 355);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(74, 23);
            this.button2.TabIndex = 25;
            this.button2.Text = "Gọi Thầy";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.buttonTopMost_Click);
            // 
            // comboBoxCauDaDanhDau
            // 
            this.comboBoxCauDaDanhDau.BackColor = System.Drawing.SystemColors.MenuText;
            this.comboBoxCauDaDanhDau.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxCauDaDanhDau.ForeColor = System.Drawing.Color.Blue;
            this.comboBoxCauDaDanhDau.FormattingEnabled = true;
            this.comboBoxCauDaDanhDau.Location = new System.Drawing.Point(815, 282);
            this.comboBoxCauDaDanhDau.Name = "comboBoxCauDaDanhDau";
            this.comboBoxCauDaDanhDau.Size = new System.Drawing.Size(52, 29);
            this.comboBoxCauDaDanhDau.TabIndex = 27;
            // 
            // buttonHelpVideo
            // 
            this.buttonHelpVideo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.buttonHelpVideo.Location = new System.Drawing.Point(256, 355);
            this.buttonHelpVideo.Name = "buttonHelpVideo";
            this.buttonHelpVideo.Size = new System.Drawing.Size(73, 23);
            this.buttonHelpVideo.TabIndex = 29;
            this.buttonHelpVideo.Text = "Chỉ làm bài";
            this.buttonHelpVideo.UseVisualStyleBackColor = false;
            this.buttonHelpVideo.Click += new System.EventHandler(this.buttonHelpVideo_Click);
            // 
            // buttonHelp
            // 
            this.buttonHelp.Location = new System.Drawing.Point(451, 288);
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Size = new System.Drawing.Size(114, 23);
            this.buttonHelp.TabIndex = 30;
            this.buttonHelp.Text = "Về chế độ thường";
            this.buttonHelp.UseVisualStyleBackColor = true;
            this.buttonHelp.Click += new System.EventHandler(this.buttonHelp_Click_1);
            // 
            // richTextBox2
            // 
            this.richTextBox2.Font = new System.Drawing.Font("Times New Roman", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBox2.Location = new System.Drawing.Point(159, 395);
            this.richTextBox2.Name = "richTextBox2";
            this.richTextBox2.ReadOnly = true;
            this.richTextBox2.Size = new System.Drawing.Size(103, 82);
            this.richTextBox2.TabIndex = 6;
            this.richTextBox2.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.ClientSize = new System.Drawing.Size(1016, 599);
            this.Controls.Add(this.buttonHelp);
            this.Controls.Add(this.buttonHelpVideo);
            this.Controls.Add(this.labelKQ);
            this.Controls.Add(this.comboBoxCauDaDanhDau);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.buttonVideoHelp);
            this.Controls.Add(this.buttonZoom);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.comboBoxCauNext);
            this.Controls.Add(this.labelCauHienTai);
            this.Controls.Add(this.buttonCheck);
            this.Controls.Add(this.richTextBox2);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.buttonRefresh);
            this.Controls.Add(this.buttonEV);
            this.Controls.Add(this.buttonReset);
            this.Controls.Add(this.buttonNext);
            this.Controls.Add(this.buttonExit);
            this.Controls.Add(this.buttonSummary);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Form1";
            this.Text = "Form1";
            this.TopMost = true;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private Button buttonExit;
        private Button buttonNext;
        private Button buttonReset;
        private Button buttonEV;
        private Button buttonRefresh;
        private RichTextBox richTextBox1;
        private Button buttonCheck;
        private PictureBox pictureBox1;
        private Label labelKQ;
        private Label labelCauHienTai;
        private ComboBox comboBoxCauNext;
        private Panel panel1;
        private Button buttonSummary;
        private Button buttonZoom;
        private Button buttonVideoHelp;
        private Button button2;
        private ComboBox comboBoxCauDaDanhDau;
        private Button buttonHelpVideo;
        private Button buttonHelp;
        private RichTextBox richTextBox2;
    }
}