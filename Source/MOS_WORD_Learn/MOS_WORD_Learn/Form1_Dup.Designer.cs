using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    partial class Form1_Dup
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
            this.buttonRefresh = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.buttonCheck = new System.Windows.Forms.Button();
            this.labelKQ = new System.Windows.Forms.Label();
            this.labelCauHienTai = new System.Windows.Forms.Label();
            this.comboBoxCauNext = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.richTextQuestion = new System.Windows.Forms.RichTextBox();
            this.buttonSummary = new System.Windows.Forms.Button();
            this.buttonZoom = new System.Windows.Forms.Button();
            this.buttonVideoHelp = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.comboBoxCauDaDanhDau = new System.Windows.Forms.ComboBox();
            this.buttonHelpVideo = new System.Windows.Forms.Button();
            this.buttonHelp = new System.Windows.Forms.Button();
            this.buttonEV = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonExit
            // 
            this.buttonExit.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonExit.Location = new System.Drawing.Point(80, 248);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new System.Drawing.Size(75, 25);
            this.buttonExit.TabIndex = 0;
            this.buttonExit.Text = "Thoát";
            this.buttonExit.UseVisualStyleBackColor = true;
            this.buttonExit.Click += new System.EventHandler(this.buttonExit_Click);
            // 
            // buttonNext
            // 
            this.buttonNext.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonNext.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.buttonNext.Location = new System.Drawing.Point(716, 324);
            this.buttonNext.Name = "buttonNext";
            this.buttonNext.Size = new System.Drawing.Size(75, 25);
            this.buttonNext.TabIndex = 1;
            this.buttonNext.Text = "Câu tiếp";
            this.buttonNext.UseVisualStyleBackColor = true;
            this.buttonNext.Click += new System.EventHandler(this.buttonNext_Click);
            // 
            // buttonReset
            // 
            this.buttonReset.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonReset.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.buttonReset.Location = new System.Drawing.Point(629, 324);
            this.buttonReset.Name = "buttonReset";
            this.buttonReset.Size = new System.Drawing.Size(75, 25);
            this.buttonReset.TabIndex = 2;
            this.buttonReset.Text = "Reset";
            this.buttonReset.UseVisualStyleBackColor = true;
            this.buttonReset.Click += new System.EventHandler(this.buttonReset_Click);
            // 
            // buttonRefresh
            // 
            this.buttonRefresh.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonRefresh.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.buttonRefresh.Location = new System.Drawing.Point(161, 248);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new System.Drawing.Size(120, 25);
            this.buttonRefresh.TabIndex = 5;
            this.buttonRefresh.Text = "Phục hồi Word";
            this.buttonRefresh.UseVisualStyleBackColor = true;
            this.buttonRefresh.Click += new System.EventHandler(this.buttonRefresh_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBox1.Location = new System.Drawing.Point(80, 92);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(103, 82);
            this.richTextBox1.TabIndex = 6;
            this.richTextBox1.Text = "";
            // 
            // buttonCheck
            // 
            this.buttonCheck.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.buttonCheck.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonCheck.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.buttonCheck.Location = new System.Drawing.Point(574, 248);
            this.buttonCheck.Name = "buttonCheck";
            this.buttonCheck.Size = new System.Drawing.Size(120, 25);
            this.buttonCheck.TabIndex = 7;
            this.buttonCheck.Text = "Xem đúng/sai";
            this.buttonCheck.UseVisualStyleBackColor = false;
            this.buttonCheck.Click += new System.EventHandler(this.buttonCheck_Click);
            // 
            // labelKQ
            // 
            this.labelKQ.AutoSize = true;
            this.labelKQ.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelKQ.ForeColor = System.Drawing.Color.Aqua;
            this.labelKQ.Location = new System.Drawing.Point(606, 218);
            this.labelKQ.Name = "labelKQ";
            this.labelKQ.Size = new System.Drawing.Size(0, 24);
            this.labelKQ.TabIndex = 10;
            // 
            // labelCauHienTai
            // 
            this.labelCauHienTai.AutoSize = true;
            this.labelCauHienTai.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelCauHienTai.ForeColor = System.Drawing.Color.Yellow;
            this.labelCauHienTai.Location = new System.Drawing.Point(529, 325);
            this.labelCauHienTai.Name = "labelCauHienTai";
            this.labelCauHienTai.Size = new System.Drawing.Size(86, 24);
            this.labelCauHienTai.TabIndex = 12;
            this.labelCauHienTai.Text = "Câu 150";
            // 
            // comboBoxCauNext
            // 
            this.comboBoxCauNext.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.comboBoxCauNext.Cursor = System.Windows.Forms.Cursors.Hand;
            this.comboBoxCauNext.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F);
            this.comboBoxCauNext.ForeColor = System.Drawing.Color.Red;
            this.comboBoxCauNext.FormattingEnabled = true;
            this.comboBoxCauNext.Location = new System.Drawing.Point(410, 322);
            this.comboBoxCauNext.Name = "comboBoxCauNext";
            this.comboBoxCauNext.Size = new System.Drawing.Size(62, 28);
            this.comboBoxCauNext.TabIndex = 13;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.richTextQuestion);
            this.panel1.Location = new System.Drawing.Point(490, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(282, 172);
            this.panel1.TabIndex = 15;
            // 
            // richTextQuestion
            // 
            this.richTextQuestion.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextQuestion.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextQuestion.Location = new System.Drawing.Point(0, 0);
            this.richTextQuestion.Name = "richTextQuestion";
            this.richTextQuestion.ReadOnly = true;
            this.richTextQuestion.Size = new System.Drawing.Size(280, 170);
            this.richTextQuestion.TabIndex = 0;
            this.richTextQuestion.Text = "";
            // 
            // buttonSummary
            // 
            this.buttonSummary.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonSummary.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.buttonSummary.Location = new System.Drawing.Point(478, 324);
            this.buttonSummary.Name = "buttonSummary";
            this.buttonSummary.Size = new System.Drawing.Size(45, 25);
            this.buttonSummary.TabIndex = 17;
            this.buttonSummary.Text = "Xóa";
            this.buttonSummary.UseVisualStyleBackColor = true;
            this.buttonSummary.Click += new System.EventHandler(this.buttonSummary_Click);
            // 
            // buttonZoom
            // 
            this.buttonZoom.BackColor = System.Drawing.Color.Cyan;
            this.buttonZoom.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonZoom.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.buttonZoom.ForeColor = System.Drawing.Color.Black;
            this.buttonZoom.Location = new System.Drawing.Point(326, 324);
            this.buttonZoom.Name = "buttonZoom";
            this.buttonZoom.Size = new System.Drawing.Size(78, 25);
            this.buttonZoom.TabIndex = 20;
            this.buttonZoom.Text = "Đánh dấu";
            this.buttonZoom.UseVisualStyleBackColor = false;
            this.buttonZoom.Click += new System.EventHandler(this.buttonZoom_Click);
            // 
            // buttonVideoHelp
            // 
            this.buttonVideoHelp.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.buttonVideoHelp.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonVideoHelp.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.buttonVideoHelp.Location = new System.Drawing.Point(80, 324);
            this.buttonVideoHelp.Name = "buttonVideoHelp";
            this.buttonVideoHelp.Size = new System.Drawing.Size(160, 25);
            this.buttonVideoHelp.TabIndex = 24;
            this.buttonVideoHelp.Text = "Hướng dẫn phần mềm";
            this.buttonVideoHelp.UseVisualStyleBackColor = false;
            this.buttonVideoHelp.Click += new System.EventHandler(this.buttonVideoHelp_Click);
            // 
            // button2
            // 
            this.button2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.button2.Location = new System.Drawing.Point(246, 324);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(74, 25);
            this.button2.TabIndex = 25;
            this.button2.Text = "Gọi Cô";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.buttonTopMost_Click);
            // 
            // comboBoxCauDaDanhDau
            // 
            this.comboBoxCauDaDanhDau.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.comboBoxCauDaDanhDau.Cursor = System.Windows.Forms.Cursors.Hand;
            this.comboBoxCauDaDanhDau.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F);
            this.comboBoxCauDaDanhDau.ForeColor = System.Drawing.Color.Blue;
            this.comboBoxCauDaDanhDau.FormattingEnabled = true;
            this.comboBoxCauDaDanhDau.ItemHeight = 20;
            this.comboBoxCauDaDanhDau.Location = new System.Drawing.Point(797, 322);
            this.comboBoxCauDaDanhDau.Name = "comboBoxCauDaDanhDau";
            this.comboBoxCauDaDanhDau.Size = new System.Drawing.Size(62, 28);
            this.comboBoxCauDaDanhDau.TabIndex = 27;
            // 
            // buttonHelpVideo
            // 
            this.buttonHelpVideo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.buttonHelpVideo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonHelpVideo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonHelpVideo.Location = new System.Drawing.Point(428, 248);
            this.buttonHelpVideo.Name = "buttonHelpVideo";
            this.buttonHelpVideo.Size = new System.Drawing.Size(140, 25);
            this.buttonHelpVideo.TabIndex = 29;
            this.buttonHelpVideo.Text = "Hướng dẫn làm bài";
            this.buttonHelpVideo.UseVisualStyleBackColor = false;
            this.buttonHelpVideo.Click += new System.EventHandler(this.buttonHelpVideo_Click);
            // 
            // buttonHelp
            // 
            this.buttonHelp.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonHelp.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.buttonHelp.Location = new System.Drawing.Point(287, 248);
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Size = new System.Drawing.Size(135, 25);
            this.buttonHelp.TabIndex = 30;
            this.buttonHelp.Text = "Về chế độ thường";
            this.buttonHelp.UseVisualStyleBackColor = true;
            this.buttonHelp.Click += new System.EventHandler(this.buttonHelp_Click_1);
            // 
            // buttonEV
            // 
            this.buttonEV.BackgroundImage = global::WindowsFormsApplication1.Properties.Resources.vi;
            this.buttonEV.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.buttonEV.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonEV.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonEV.Font = new System.Drawing.Font("Microsoft Sans Serif", 0.001F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonEV.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(102)))), ((int)(((byte)(153)))));
            this.buttonEV.Location = new System.Drawing.Point(855, 326);
            this.buttonEV.Margin = new System.Windows.Forms.Padding(0);
            this.buttonEV.Name = "buttonEV";
            this.buttonEV.Size = new System.Drawing.Size(53, 25);
            this.buttonEV.TabIndex = 3;
            this.buttonEV.Text = "V";
            this.buttonEV.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            this.buttonEV.UseVisualStyleBackColor = true;
            this.buttonEV.Click += new System.EventHandler(this.buttonEV_Click);
            // 
            // Form1_Dup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(102)))), ((int)(((byte)(153)))));
            this.ClientSize = new System.Drawing.Size(1075, 422);
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
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.buttonRefresh);
            this.Controls.Add(this.buttonEV);
            this.Controls.Add(this.buttonReset);
            this.Controls.Add(this.buttonNext);
            this.Controls.Add(this.buttonExit);
            this.Controls.Add(this.buttonSummary);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Form1_Dup";
            this.Text = "Form1_Dup";
            this.TopMost = true;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Dup_Load);
            this.panel1.ResumeLayout(false);
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
        private RichTextBox richTextQuestion;
    }
}