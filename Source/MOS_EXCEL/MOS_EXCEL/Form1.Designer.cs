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
            this.buttonExit = new MOS_EXCEL_LEARN.Base.RJButton();
            this.buttonNext = new MOS_EXCEL_LEARN.Base.RJButton();
            this.buttonReset = new MOS_EXCEL_LEARN.Base.RJButton();
            this.buttonEV = new MOS_EXCEL_LEARN.Base.RJButton();
            this.buttonHelp = new MOS_EXCEL_LEARN.Base.RJButton();
            this.buttonRefresh = new MOS_EXCEL_LEARN.Base.RJButton();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.buttonCheck = new MOS_EXCEL_LEARN.Base.RJButton();
            this.labelKQ = new System.Windows.Forms.Label();
            this.labelCauHienTai = new System.Windows.Forms.Label();
            this.comboBoxCauNext = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.buttonZoom = new MOS_EXCEL_LEARN.Base.RJButton();
            this.buttonSummary = new MOS_EXCEL_LEARN.Base.RJButton();
            this.buttonVideoHelp = new MOS_EXCEL_LEARN.Base.RJButton();
            this.comboBoxCauDaDanhDau = new System.Windows.Forms.ComboBox();
            this.buttonHelpVideo = new MOS_EXCEL_LEARN.Base.RJButton();
            this.richTextQuestion = new System.Windows.Forms.RichTextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonExit
            // 
            this.buttonExit.AutoSize = true;
            this.buttonExit.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonExit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(3)))), ((int)(((byte)(3)))));
            this.buttonExit.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(3)))), ((int)(((byte)(3)))));
            this.buttonExit.BorderColor = System.Drawing.Color.White;
            this.buttonExit.BorderRadius = 5;
            this.buttonExit.BorderSize = 2;
            this.buttonExit.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonExit.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonExit.ForeColor = System.Drawing.Color.White;
            this.buttonExit.Location = new System.Drawing.Point(61, 455);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Padding = new System.Windows.Forms.Padding(3);
            this.buttonExit.Size = new System.Drawing.Size(36, 37);
            this.buttonExit.TabIndex = 0;
            this.buttonExit.Text = "X";
            this.buttonExit.TextColor = System.Drawing.Color.White;
            this.buttonExit.UseVisualStyleBackColor = false;
            this.buttonExit.Click += new System.EventHandler(this.button1_Click);
            // 
            // buttonNext
            // 
            this.buttonNext.AutoSize = true;
            this.buttonNext.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonNext.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonNext.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonNext.BorderColor = System.Drawing.Color.White;
            this.buttonNext.BorderRadius = 5;
            this.buttonNext.BorderSize = 2;
            this.buttonNext.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonNext.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonNext.ForeColor = System.Drawing.Color.White;
            this.buttonNext.Location = new System.Drawing.Point(559, 532);
            this.buttonNext.Name = "buttonNext";
            this.buttonNext.Padding = new System.Windows.Forms.Padding(3);
            this.buttonNext.Size = new System.Drawing.Size(90, 37);
            this.buttonNext.TabIndex = 1;
            this.buttonNext.Text = "Chọn câu";
            this.buttonNext.TextColor = System.Drawing.Color.White;
            this.buttonNext.UseVisualStyleBackColor = false;
            // 
            // buttonReset
            // 
            this.buttonReset.AutoSize = true;
            this.buttonReset.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonReset.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonReset.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonReset.BorderColor = System.Drawing.Color.White;
            this.buttonReset.BorderRadius = 5;
            this.buttonReset.BorderSize = 2;
            this.buttonReset.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonReset.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonReset.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonReset.ForeColor = System.Drawing.Color.White;
            this.buttonReset.Location = new System.Drawing.Point(492, 532);
            this.buttonReset.Name = "buttonReset";
            this.buttonReset.Padding = new System.Windows.Forms.Padding(3);
            this.buttonReset.Size = new System.Drawing.Size(64, 37);
            this.buttonReset.TabIndex = 2;
            this.buttonReset.Text = "Reset";
            this.buttonReset.TextColor = System.Drawing.Color.White;
            this.buttonReset.UseVisualStyleBackColor = false;
            this.buttonReset.Click += new System.EventHandler(this.buttonReset_Click);
            // 
            // buttonEV
            // 
            this.buttonEV.AutoSize = true;
            this.buttonEV.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonEV.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonEV.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonEV.BorderColor = System.Drawing.Color.White;
            this.buttonEV.BorderRadius = 5;
            this.buttonEV.BorderSize = 2;
            this.buttonEV.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonEV.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonEV.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonEV.ForeColor = System.Drawing.Color.White;
            this.buttonEV.Location = new System.Drawing.Point(711, 528);
            this.buttonEV.Name = "buttonEV";
            this.buttonEV.Padding = new System.Windows.Forms.Padding(3);
            this.buttonEV.Size = new System.Drawing.Size(96, 37);
            this.buttonEV.TabIndex = 3;
            this.buttonEV.Text = "Tiếng Anh";
            this.buttonEV.TextColor = System.Drawing.Color.White;
            this.buttonEV.UseVisualStyleBackColor = false;
            this.buttonEV.Click += new System.EventHandler(this.buttonEV_Click);
            // 
            // buttonHelp
            // 
            this.buttonHelp.AutoSize = true;
            this.buttonHelp.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonHelp.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonHelp.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonHelp.BorderColor = System.Drawing.Color.White;
            this.buttonHelp.BorderRadius = 5;
            this.buttonHelp.BorderSize = 2;
            this.buttonHelp.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonHelp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonHelp.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonHelp.ForeColor = System.Drawing.Color.White;
            this.buttonHelp.Location = new System.Drawing.Point(416, 455);
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Padding = new System.Windows.Forms.Padding(3);
            this.buttonHelp.Size = new System.Drawing.Size(114, 37);
            this.buttonHelp.TabIndex = 4;
            this.buttonHelp.Text = "Hiện Taskbar";
            this.buttonHelp.TextColor = System.Drawing.Color.White;
            this.buttonHelp.UseVisualStyleBackColor = false;
            this.buttonHelp.Click += new System.EventHandler(this.buttonHelp_Click);
            // 
            // buttonRefresh
            // 
            this.buttonRefresh.AutoSize = true;
            this.buttonRefresh.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonRefresh.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonRefresh.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonRefresh.BorderColor = System.Drawing.Color.White;
            this.buttonRefresh.BorderRadius = 5;
            this.buttonRefresh.BorderSize = 2;
            this.buttonRefresh.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonRefresh.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonRefresh.ForeColor = System.Drawing.Color.White;
            this.buttonRefresh.Location = new System.Drawing.Point(288, 455);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Padding = new System.Windows.Forms.Padding(3);
            this.buttonRefresh.Size = new System.Drawing.Size(121, 37);
            this.buttonRefresh.TabIndex = 5;
            this.buttonRefresh.Text = "Hiện/Ẩn Excel";
            this.buttonRefresh.TextColor = System.Drawing.Color.White;
            this.buttonRefresh.UseVisualStyleBackColor = false;
            this.buttonRefresh.Click += new System.EventHandler(this.buttonRefresh_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Font = new System.Drawing.Font("Times New Roman", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBox1.Location = new System.Drawing.Point(61, 159);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(103, 82);
            this.richTextBox1.TabIndex = 6;
            this.richTextBox1.Text = "";
            // 
            // buttonCheck
            // 
            this.buttonCheck.AutoSize = true;
            this.buttonCheck.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonCheck.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonCheck.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonCheck.BorderColor = System.Drawing.Color.White;
            this.buttonCheck.BorderRadius = 5;
            this.buttonCheck.BorderSize = 2;
            this.buttonCheck.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonCheck.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonCheck.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonCheck.ForeColor = System.Drawing.Color.White;
            this.buttonCheck.Location = new System.Drawing.Point(675, 455);
            this.buttonCheck.Name = "buttonCheck";
            this.buttonCheck.Padding = new System.Windows.Forms.Padding(3);
            this.buttonCheck.Size = new System.Drawing.Size(90, 37);
            this.buttonCheck.TabIndex = 7;
            this.buttonCheck.Text = "Đúng/Sai";
            this.buttonCheck.TextColor = System.Drawing.Color.White;
            this.buttonCheck.UseVisualStyleBackColor = false;
            this.buttonCheck.Click += new System.EventHandler(this.buttonCheck_Click);
            // 
            // labelKQ
            // 
            this.labelKQ.AutoSize = true;
            this.labelKQ.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelKQ.ForeColor = System.Drawing.SystemColors.Highlight;
            this.labelKQ.Location = new System.Drawing.Point(516, 220);
            this.labelKQ.Name = "labelKQ";
            this.labelKQ.Size = new System.Drawing.Size(0, 21);
            this.labelKQ.TabIndex = 10;
            // 
            // labelCauHienTai
            // 
            this.labelCauHienTai.AutoSize = true;
            this.labelCauHienTai.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(3)))), ((int)(((byte)(3)))));
            this.labelCauHienTai.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.labelCauHienTai.ForeColor = System.Drawing.Color.White;
            this.labelCauHienTai.Location = new System.Drawing.Point(406, 532);
            this.labelCauHienTai.MinimumSize = new System.Drawing.Size(80, 37);
            this.labelCauHienTai.Name = "labelCauHienTai";
            this.labelCauHienTai.Padding = new System.Windows.Forms.Padding(9);
            this.labelCauHienTai.Size = new System.Drawing.Size(80, 37);
            this.labelCauHienTai.TabIndex = 12;
            this.labelCauHienTai.Text = "Câu 150";
            // 
            // comboBoxCauNext
            // 
            this.comboBoxCauNext.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.comboBoxCauNext.Cursor = System.Windows.Forms.Cursors.Hand;
            this.comboBoxCauNext.Font = new System.Drawing.Font("Calibri", 17F, System.Drawing.FontStyle.Bold);
            this.comboBoxCauNext.ForeColor = System.Drawing.Color.White;
            this.comboBoxCauNext.FormattingEnabled = true;
            this.comboBoxCauNext.Location = new System.Drawing.Point(653, 532);
            this.comboBoxCauNext.Name = "comboBoxCauNext";
            this.comboBoxCauNext.Size = new System.Drawing.Size(52, 36);
            this.comboBoxCauNext.TabIndex = 13;
            this.comboBoxCauNext.Text = "1";
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.richTextQuestion);
            this.panel1.Location = new System.Drawing.Point(573, 35);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(282, 172);
            this.panel1.TabIndex = 15;
            // 
            // buttonZoom
            // 
            this.buttonZoom.AutoSize = true;
            this.buttonZoom.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonZoom.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonZoom.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonZoom.BorderColor = System.Drawing.Color.White;
            this.buttonZoom.BorderRadius = 5;
            this.buttonZoom.BorderSize = 2;
            this.buttonZoom.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonZoom.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonZoom.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonZoom.ForeColor = System.Drawing.Color.White;
            this.buttonZoom.Location = new System.Drawing.Point(61, 528);
            this.buttonZoom.Name = "buttonZoom";
            this.buttonZoom.Padding = new System.Windows.Forms.Padding(3);
            this.buttonZoom.Size = new System.Drawing.Size(93, 37);
            this.buttonZoom.TabIndex = 21;
            this.buttonZoom.Text = "Đánh dấu";
            this.buttonZoom.TextColor = System.Drawing.Color.White;
            this.buttonZoom.UseVisualStyleBackColor = false;
            this.buttonZoom.Click += new System.EventHandler(this.buttonZoom_Click);
            // 
            // buttonSummary
            // 
            this.buttonSummary.AutoSize = true;
            this.buttonSummary.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonSummary.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonSummary.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonSummary.BorderColor = System.Drawing.Color.White;
            this.buttonSummary.BorderRadius = 5;
            this.buttonSummary.BorderSize = 2;
            this.buttonSummary.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonSummary.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSummary.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonSummary.ForeColor = System.Drawing.Color.White;
            this.buttonSummary.Location = new System.Drawing.Point(339, 532);
            this.buttonSummary.Name = "buttonSummary";
            this.buttonSummary.Padding = new System.Windows.Forms.Padding(3);
            this.buttonSummary.Size = new System.Drawing.Size(53, 37);
            this.buttonSummary.TabIndex = 23;
            this.buttonSummary.Text = "Xóa";
            this.buttonSummary.TextColor = System.Drawing.Color.White;
            this.buttonSummary.UseVisualStyleBackColor = false;
            this.buttonSummary.Click += new System.EventHandler(this.buttonSummary_Click);
            // 
            // buttonVideoHelp
            // 
            this.buttonVideoHelp.AutoSize = true;
            this.buttonVideoHelp.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonVideoHelp.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonVideoHelp.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonVideoHelp.BorderColor = System.Drawing.Color.White;
            this.buttonVideoHelp.BorderRadius = 5;
            this.buttonVideoHelp.BorderSize = 2;
            this.buttonVideoHelp.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonVideoHelp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonVideoHelp.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonVideoHelp.ForeColor = System.Drawing.Color.White;
            this.buttonVideoHelp.Location = new System.Drawing.Point(142, 455);
            this.buttonVideoHelp.Name = "buttonVideoHelp";
            this.buttonVideoHelp.Padding = new System.Windows.Forms.Padding(3);
            this.buttonVideoHelp.Size = new System.Drawing.Size(65, 37);
            this.buttonVideoHelp.TabIndex = 24;
            this.buttonVideoHelp.Text = "HDSD";
            this.buttonVideoHelp.TextColor = System.Drawing.Color.White;
            this.buttonVideoHelp.UseVisualStyleBackColor = false;
            this.buttonVideoHelp.Click += new System.EventHandler(this.buttonVideoHelp_Click);
            // 
            // comboBoxCauDaDanhDau
            // 
            this.comboBoxCauDaDanhDau.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.comboBoxCauDaDanhDau.Cursor = System.Windows.Forms.Cursors.Hand;
            this.comboBoxCauDaDanhDau.Font = new System.Drawing.Font("Calibri", 17F, System.Drawing.FontStyle.Bold);
            this.comboBoxCauDaDanhDau.ForeColor = System.Drawing.Color.White;
            this.comboBoxCauDaDanhDau.FormattingEnabled = true;
            this.comboBoxCauDaDanhDau.Location = new System.Drawing.Point(171, 528);
            this.comboBoxCauDaDanhDau.Name = "comboBoxCauDaDanhDau";
            this.comboBoxCauDaDanhDau.Size = new System.Drawing.Size(43, 36);
            this.comboBoxCauDaDanhDau.TabIndex = 26;
            // 
            // buttonHelpVideo
            // 
            this.buttonHelpVideo.AutoSize = true;
            this.buttonHelpVideo.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonHelpVideo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonHelpVideo.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.buttonHelpVideo.BorderColor = System.Drawing.Color.White;
            this.buttonHelpVideo.BorderRadius = 5;
            this.buttonHelpVideo.BorderSize = 2;
            this.buttonHelpVideo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonHelpVideo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonHelpVideo.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonHelpVideo.ForeColor = System.Drawing.Color.White;
            this.buttonHelpVideo.Location = new System.Drawing.Point(546, 457);
            this.buttonHelpVideo.Name = "buttonHelpVideo";
            this.buttonHelpVideo.Padding = new System.Windows.Forms.Padding(3);
            this.buttonHelpVideo.Size = new System.Drawing.Size(125, 37);
            this.buttonHelpVideo.TabIndex = 30;
            this.buttonHelpVideo.Text = "Video chữa đề";
            this.buttonHelpVideo.TextColor = System.Drawing.Color.White;
            this.buttonHelpVideo.UseVisualStyleBackColor = false;
            this.buttonHelpVideo.Click += new System.EventHandler(this.buttonHelpVideo_Click);
            // 
            // richTextQuestion
            // 
            this.richTextQuestion.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextQuestion.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextQuestion.Font = new System.Drawing.Font("Calibri", 20F);
            this.richTextQuestion.Location = new System.Drawing.Point(0, 0);
            this.richTextQuestion.Margin = new System.Windows.Forms.Padding(10);
            this.richTextQuestion.Name = "richTextQuestion";
            this.richTextQuestion.ReadOnly = true;
            this.richTextQuestion.Size = new System.Drawing.Size(282, 172);
            this.richTextQuestion.TabIndex = 1;
            this.richTextQuestion.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(106)))), ((int)(((byte)(63)))));
            this.ClientSize = new System.Drawing.Size(1016, 599);
            this.Controls.Add(this.buttonHelpVideo);
            this.Controls.Add(this.labelKQ);
            this.Controls.Add(this.comboBoxCauDaDanhDau);
            this.Controls.Add(this.buttonZoom);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.comboBoxCauNext);
            this.Controls.Add(this.labelCauHienTai);
            this.Controls.Add(this.buttonCheck);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.buttonRefresh);
            this.Controls.Add(this.buttonHelp);
            this.Controls.Add(this.buttonEV);
            this.Controls.Add(this.buttonReset);
            this.Controls.Add(this.buttonNext);
            this.Controls.Add(this.buttonExit);
            this.Controls.Add(this.buttonSummary);
            this.Controls.Add(this.buttonVideoHelp);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Form1";
            this.Text = "Form1_New";
            this.TopMost = true;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_New_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private Base.RJButton buttonExit;
        private Base.RJButton buttonNext;
        private Base.RJButton buttonReset;
        private Base.RJButton buttonEV;
        private Base.RJButton buttonHelp;
        private Base.RJButton buttonRefresh;
        private RichTextBox richTextBox1;
        private Base.RJButton buttonCheck;
        private Label labelKQ;
        private Label labelCauHienTai;
        private ComboBox comboBoxCauNext;
        private Panel panel1;
        private Base.RJButton buttonZoom;
        private Base.RJButton buttonSummary;
        private Base.RJButton buttonVideoHelp;
        private Base.RJButton button2;
        private ComboBox comboBoxCauDaDanhDau;
        private Base.RJButton buttonHelpVideo;
        private RichTextBox richTextQuestion;
    }
}