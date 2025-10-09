using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using MOS_WORD_TEST.Properties;
using MOS_WORD_TEST.Base;

namespace MOS_WORD_TEST
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
            this.components = new System.ComponentModel.Container();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.richTextTA = new System.Windows.Forms.RichTextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.panelQuestion = new System.Windows.Forms.Panel();
            this.panelJumpQuestion = new System.Windows.Forms.Panel();
            this.panelQuestionContent = new System.Windows.Forms.Panel();
            this.labelProject = new System.Windows.Forms.Label();
            this.btnMaskReview = new MOS_WORD_TEST.Base.RJButton();
            this.btnMaskComplete = new MOS_WORD_TEST.Base.RJButton();
            this.button2 = new MOS_WORD_TEST.Base.RJButton();
            this.buttonShowHide = new MOS_WORD_TEST.Base.RJButton();
            this.buttonxhdh = new MOS_WORD_TEST.Base.RJButton();
            this.buttonRefresh = new MOS_WORD_TEST.Base.RJButton();
            this.buttonHelp = new MOS_WORD_TEST.Base.RJButton();
            this.button1 = new MOS_WORD_TEST.Base.RJButton();
            this.buttonSave = new MOS_WORD_TEST.Base.RJButton();
            this.buttonExit = new MOS_WORD_TEST.Base.RJButton();
            this.buttonSubmit = new MOS_WORD_TEST.Base.RJButton();
            this.buttonCheck = new MOS_WORD_TEST.Base.RJButton();
            this.buttonReset = new MOS_WORD_TEST.Base.RJButton();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panelQuestion.SuspendLayout();
            this.panelQuestionContent.SuspendLayout();
            this.SuspendLayout();
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(170, 28);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(122, 109);
            this.checkedListBox1.TabIndex = 0;
            this.checkedListBox1.Visible = false;
            this.checkedListBox1.SelectedIndexChanged += new System.EventHandler(this.checkedListBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 18F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(24, 177);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 29);
            this.label1.TabIndex = 6;
            this.label1.Text = "00:00";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(305, 192);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.panel2);
            this.tabPage2.Location = new System.Drawing.Point(4, 28);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(297, 160);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Question";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.AutoScroll = true;
            this.panel2.Controls.Add(this.richTextTA);
            this.panel2.Location = new System.Drawing.Point(3, 3);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(512, 264);
            this.panel2.TabIndex = 0;
            // 
            // richTextTA
            // 
            this.richTextTA.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextTA.Font = new System.Drawing.Font("Calibri", 18F);
            this.richTextTA.Location = new System.Drawing.Point(0, 0);
            this.richTextTA.Name = "richTextTA";
            this.richTextTA.Size = new System.Drawing.Size(512, 264);
            this.richTextTA.TabIndex = 0;
            this.richTextTA.Text = "";
            // 
            // textBox1
            // 
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(29, 33);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(100, 20);
            this.textBox1.TabIndex = 13;
            this.textBox1.Visible = false;
            // 
            // richTextBox1
            // 
            this.richTextBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.richTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBox1.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.richTextBox1.ForeColor = System.Drawing.Color.White;
            this.richTextBox1.Location = new System.Drawing.Point(3, 7);
            this.richTextBox1.MinimumSize = new System.Drawing.Size(150, 90);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(150, 90);
            this.richTextBox1.TabIndex = 16;
            this.richTextBox1.Text = "";
            // 
            // panel4
            // 
            this.panel4.AutoScroll = true;
            this.panel4.Controls.Add(this.richTextBox1);
            this.panel4.Location = new System.Drawing.Point(315, 33);
            this.panel4.MinimumSize = new System.Drawing.Size(190, 100);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(190, 100);
            this.panel4.TabIndex = 17;
            this.panel4.Visible = false;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(29, 90);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(22, 20);
            this.textBox2.TabIndex = 21;
            this.textBox2.Text = "1";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(604, 35);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(44, 21);
            this.comboBox1.TabIndex = 31;
            // 
            // panelQuestion
            // 
            this.panelQuestion.Controls.Add(this.panelJumpQuestion);
            this.panelQuestion.Controls.Add(this.panelQuestionContent);
            this.panelQuestion.Location = new System.Drawing.Point(29, 209);
            this.panelQuestion.Name = "panelQuestion";
            this.panelQuestion.Size = new System.Drawing.Size(305, 192);
            this.panelQuestion.TabIndex = 32;
            // 
            // panelJumpQuestion
            // 
            this.panelJumpQuestion.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.panelJumpQuestion.Location = new System.Drawing.Point(157, 97);
            this.panelJumpQuestion.Name = "panelJumpQuestion";
            this.panelJumpQuestion.Size = new System.Drawing.Size(184, 26);
            this.panelJumpQuestion.TabIndex = 33;
            // 
            // panelQuestionContent
            // 
            this.panelQuestionContent.Controls.Add(this.tabControl1);
            this.panelQuestionContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelQuestionContent.Location = new System.Drawing.Point(0, 0);
            this.panelQuestionContent.Name = "panelQuestionContent";
            this.panelQuestionContent.Size = new System.Drawing.Size(305, 192);
            this.panelQuestionContent.TabIndex = 0;
            // 
            // labelProject
            // 
            this.labelProject.AutoSize = true;
            this.labelProject.Cursor = System.Windows.Forms.Cursors.Hand;
            this.labelProject.Font = new System.Drawing.Font("Calibri", 14F, System.Drawing.FontStyle.Bold);
            this.labelProject.ForeColor = System.Drawing.Color.White;
            this.labelProject.Location = new System.Drawing.Point(245, 176);
            this.labelProject.Name = "labelProject";
            this.labelProject.Size = new System.Drawing.Size(151, 23);
            this.labelProject.TabIndex = 33;
            this.labelProject.Text = "Project ….. of …. : ";
            this.labelProject.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnMaskReview
            // 
            this.btnMaskReview.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMaskReview.AutoSize = true;
            this.btnMaskReview.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnMaskReview.BackColor = System.Drawing.Color.White;
            this.btnMaskReview.BackgroundColor = System.Drawing.Color.White;
            this.btnMaskReview.BorderColor = System.Drawing.Color.Black;
            this.btnMaskReview.BorderRadius = 5;
            this.btnMaskReview.BorderSize = 1;
            this.btnMaskReview.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMaskReview.Font = new System.Drawing.Font("Calibri", 10F);
            this.btnMaskReview.ForeColor = System.Drawing.Color.Black;
            this.btnMaskReview.Location = new System.Drawing.Point(667, 361);
            this.btnMaskReview.Name = "btnMaskReview";
            this.btnMaskReview.Size = new System.Drawing.Size(113, 29);
            this.btnMaskReview.TabIndex = 35;
            this.btnMaskReview.Text = "Mark for Review";
            this.btnMaskReview.TextColor = System.Drawing.Color.Black;
            this.btnMaskReview.UseVisualStyleBackColor = false;
            this.btnMaskReview.Click += new System.EventHandler(this.btnMaskReview_Click);
            // 
            // btnMaskComplete
            // 
            this.btnMaskComplete.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMaskComplete.AutoSize = true;
            this.btnMaskComplete.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnMaskComplete.BackColor = System.Drawing.Color.White;
            this.btnMaskComplete.BackgroundColor = System.Drawing.Color.White;
            this.btnMaskComplete.BorderColor = System.Drawing.Color.Black;
            this.btnMaskComplete.BorderRadius = 5;
            this.btnMaskComplete.BorderSize = 1;
            this.btnMaskComplete.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMaskComplete.Font = new System.Drawing.Font("Calibri", 10F);
            this.btnMaskComplete.ForeColor = System.Drawing.Color.Black;
            this.btnMaskComplete.Location = new System.Drawing.Point(554, 361);
            this.btnMaskComplete.Name = "btnMaskComplete";
            this.btnMaskComplete.Size = new System.Drawing.Size(107, 29);
            this.btnMaskComplete.TabIndex = 34;
            this.btnMaskComplete.Text = "Mask Complete";
            this.btnMaskComplete.TextColor = System.Drawing.Color.Black;
            this.btnMaskComplete.UseVisualStyleBackColor = false;
            this.btnMaskComplete.Click += new System.EventHandler(this.btnMaskComplete_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.button2.BackgroundColor = System.Drawing.Color.MediumSlateBlue;
            this.button2.BorderColor = System.Drawing.Color.PaleVioletRed;
            this.button2.BorderRadius = 0;
            this.button2.BorderSize = 0;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.ForeColor = System.Drawing.Color.White;
            this.button2.Location = new System.Drawing.Point(544, 65);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(41, 23);
            this.button2.TabIndex = 30;
            this.button2.Text = "Mở";
            this.button2.TextColor = System.Drawing.Color.White;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // buttonShowHide
            // 
            this.buttonShowHide.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.buttonShowHide.BackgroundColor = System.Drawing.Color.MediumSlateBlue;
            this.buttonShowHide.BorderColor = System.Drawing.Color.PaleVioletRed;
            this.buttonShowHide.BorderRadius = 0;
            this.buttonShowHide.BorderSize = 0;
            this.buttonShowHide.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonShowHide.ForeColor = System.Drawing.Color.White;
            this.buttonShowHide.Location = new System.Drawing.Point(522, 36);
            this.buttonShowHide.Name = "buttonShowHide";
            this.buttonShowHide.Size = new System.Drawing.Size(63, 23);
            this.buttonShowHide.TabIndex = 27;
            this.buttonShowHide.Text = "Gọi Thầy";
            this.buttonShowHide.TextColor = System.Drawing.Color.White;
            this.buttonShowHide.UseVisualStyleBackColor = true;
            this.buttonShowHide.Click += new System.EventHandler(this.buttonShowHide_Click);
            // 
            // buttonxhdh
            // 
            this.buttonxhdh.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.buttonxhdh.BackgroundColor = System.Drawing.Color.MediumSlateBlue;
            this.buttonxhdh.BorderColor = System.Drawing.Color.PaleVioletRed;
            this.buttonxhdh.BorderRadius = 0;
            this.buttonxhdh.BorderSize = 0;
            this.buttonxhdh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonxhdh.ForeColor = System.Drawing.Color.White;
            this.buttonxhdh.Location = new System.Drawing.Point(679, 35);
            this.buttonxhdh.Name = "buttonxhdh";
            this.buttonxhdh.Size = new System.Drawing.Size(114, 23);
            this.buttonxhdh.TabIndex = 25;
            this.buttonxhdh.Text = "Hướng dẫn sử dụng";
            this.buttonxhdh.TextColor = System.Drawing.Color.White;
            this.buttonxhdh.UseVisualStyleBackColor = true;
            this.buttonxhdh.Click += new System.EventHandler(this.buttonxhdh_Click);
            // 
            // buttonRefresh
            // 
            this.buttonRefresh.AutoSize = true;
            this.buttonRefresh.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonRefresh.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.buttonRefresh.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.buttonRefresh.BorderColor = System.Drawing.Color.White;
            this.buttonRefresh.BorderRadius = 5;
            this.buttonRefresh.BorderSize = 2;
            this.buttonRefresh.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonRefresh.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonRefresh.ForeColor = System.Drawing.Color.White;
            this.buttonRefresh.Location = new System.Drawing.Point(679, 171);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Padding = new System.Windows.Forms.Padding(2);
            this.buttonRefresh.Size = new System.Drawing.Size(121, 35);
            this.buttonRefresh.TabIndex = 24;
            this.buttonRefresh.Text = "Hiện / Ẩn Test";
            this.buttonRefresh.TextColor = System.Drawing.Color.White;
            this.buttonRefresh.UseVisualStyleBackColor = false;
            this.buttonRefresh.Click += new System.EventHandler(this.buttonRefresh_Click);
            // 
            // buttonHelp
            // 
            this.buttonHelp.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.buttonHelp.BackgroundColor = System.Drawing.Color.MediumSlateBlue;
            this.buttonHelp.BorderColor = System.Drawing.Color.PaleVioletRed;
            this.buttonHelp.BorderRadius = 0;
            this.buttonHelp.BorderSize = 0;
            this.buttonHelp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonHelp.ForeColor = System.Drawing.Color.White;
            this.buttonHelp.Location = new System.Drawing.Point(687, 93);
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Size = new System.Drawing.Size(106, 23);
            this.buttonHelp.TabIndex = 23;
            this.buttonHelp.Text = "Video chỉ làm bài";
            this.buttonHelp.TextColor = System.Drawing.Color.White;
            this.buttonHelp.UseVisualStyleBackColor = true;
            this.buttonHelp.Click += new System.EventHandler(this.buttonHelp_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.button1.BackgroundColor = System.Drawing.Color.MediumSlateBlue;
            this.button1.BorderColor = System.Drawing.Color.PaleVioletRed;
            this.button1.BorderRadius = 0;
            this.button1.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(58, 88);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(29, 23);
            this.button1.TabIndex = 22;
            this.button1.Text = "Go";
            this.button1.TextColor = System.Drawing.Color.White;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // buttonSave
            // 
            this.buttonSave.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.buttonSave.BackgroundColor = System.Drawing.Color.MediumSlateBlue;
            this.buttonSave.BorderColor = System.Drawing.Color.PaleVioletRed;
            this.buttonSave.BorderRadius = 0;
            this.buttonSave.BorderSize = 0;
            this.buttonSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSave.ForeColor = System.Drawing.Color.White;
            this.buttonSave.Location = new System.Drawing.Point(708, 64);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(85, 23);
            this.buttonSave.TabIndex = 14;
            this.buttonSave.Text = "Lưu lai hỏi sau";
            this.buttonSave.TextColor = System.Drawing.Color.White;
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // buttonExit
            // 
            this.buttonExit.AutoSize = true;
            this.buttonExit.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonExit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.buttonExit.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.buttonExit.BorderColor = System.Drawing.Color.White;
            this.buttonExit.BorderRadius = 5;
            this.buttonExit.BorderSize = 2;
            this.buttonExit.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonExit.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonExit.ForeColor = System.Drawing.Color.White;
            this.buttonExit.Location = new System.Drawing.Point(806, 171);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Padding = new System.Windows.Forms.Padding(2);
            this.buttonExit.Size = new System.Drawing.Size(50, 35);
            this.buttonExit.TabIndex = 5;
            this.buttonExit.Text = "Exit";
            this.buttonExit.TextColor = System.Drawing.Color.White;
            this.buttonExit.UseVisualStyleBackColor = false;
            this.buttonExit.Click += new System.EventHandler(this.button4_Click);
            // 
            // buttonSubmit
            // 
            this.buttonSubmit.AutoSize = true;
            this.buttonSubmit.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonSubmit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.buttonSubmit.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.buttonSubmit.BorderColor = System.Drawing.Color.White;
            this.buttonSubmit.BorderRadius = 5;
            this.buttonSubmit.BorderSize = 2;
            this.buttonSubmit.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonSubmit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSubmit.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonSubmit.ForeColor = System.Drawing.Color.White;
            this.buttonSubmit.Location = new System.Drawing.Point(537, 171);
            this.buttonSubmit.Name = "buttonSubmit";
            this.buttonSubmit.Padding = new System.Windows.Forms.Padding(2);
            this.buttonSubmit.Size = new System.Drawing.Size(127, 35);
            this.buttonSubmit.TabIndex = 4;
            this.buttonSubmit.Text = "Submit Project";
            this.buttonSubmit.TextColor = System.Drawing.Color.White;
            this.buttonSubmit.UseVisualStyleBackColor = false;
            this.buttonSubmit.Click += new System.EventHandler(this.button3_Click);
            // 
            // buttonCheck
            // 
            this.buttonCheck.AutoSize = true;
            this.buttonCheck.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonCheck.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.buttonCheck.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.buttonCheck.BorderColor = System.Drawing.Color.White;
            this.buttonCheck.BorderRadius = 5;
            this.buttonCheck.BorderSize = 2;
            this.buttonCheck.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonCheck.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonCheck.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonCheck.ForeColor = System.Drawing.Color.White;
            this.buttonCheck.Location = new System.Drawing.Point(98, 171);
            this.buttonCheck.Name = "buttonCheck";
            this.buttonCheck.Padding = new System.Windows.Forms.Padding(2);
            this.buttonCheck.Size = new System.Drawing.Size(133, 35);
            this.buttonCheck.TabIndex = 3;
            this.buttonCheck.Text = "Go To Summary";
            this.buttonCheck.TextColor = System.Drawing.Color.White;
            this.buttonCheck.UseVisualStyleBackColor = false;
            this.buttonCheck.Click += new System.EventHandler(this.button2_Click);
            // 
            // buttonReset
            // 
            this.buttonReset.AutoSize = true;
            this.buttonReset.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonReset.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.buttonReset.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.buttonReset.BorderColor = System.Drawing.Color.White;
            this.buttonReset.BorderRadius = 5;
            this.buttonReset.BorderSize = 2;
            this.buttonReset.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonReset.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonReset.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.buttonReset.ForeColor = System.Drawing.Color.White;
            this.buttonReset.Location = new System.Drawing.Point(402, 171);
            this.buttonReset.Name = "buttonReset";
            this.buttonReset.Padding = new System.Windows.Forms.Padding(2);
            this.buttonReset.Size = new System.Drawing.Size(127, 35);
            this.buttonReset.TabIndex = 2;
            this.buttonReset.Text = "Restart Project";
            this.buttonReset.TextColor = System.Drawing.Color.White;
            this.buttonReset.UseVisualStyleBackColor = false;
            this.buttonReset.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.ClientSize = new System.Drawing.Size(885, 457);
            this.Controls.Add(this.btnMaskReview);
            this.Controls.Add(this.btnMaskComplete);
            this.Controls.Add(this.labelProject);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.buttonShowHide);
            this.Controls.Add(this.buttonxhdh);
            this.Controls.Add(this.buttonRefresh);
            this.Controls.Add(this.buttonHelp);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.buttonSave);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonExit);
            this.Controls.Add(this.buttonSubmit);
            this.Controls.Add(this.buttonCheck);
            this.Controls.Add(this.buttonReset);
            this.Controls.Add(this.panelQuestion);
            this.Controls.Add(this.checkedListBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panelQuestion.ResumeLayout(false);
            this.panelQuestionContent.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private CheckedListBox checkedListBox1;
        private Base.RJButton buttonReset;
        private Base.RJButton buttonCheck;
        private Base.RJButton buttonSubmit;
        private Base.RJButton buttonExit;
        private Label label1;
        private Timer timer1;
        private TabControl tabControl1;
        private TabPage tabPage2;
        private Panel panel2;
        private Base.RJButton buttonSave;
        private TextBox textBox1;
        private RichTextBox richTextBox1;
        private Panel panel4;
        private Base.RJButton button1;
        private TextBox textBox2;
        private Base.RJButton buttonHelp;
        private Base.RJButton buttonRefresh;
        private Base.RJButton buttonxhdh;
        private Base.RJButton buttonShowHide;
        private Base.RJButton button2;
        private ComboBox comboBox1;
        private Panel panelQuestion;
        private Panel panelQuestionContent;
        private Panel panelJumpQuestion;
        private RichTextBox richTextTA;
        private Label labelProject;
        private RJButton btnMaskComplete;
        private RJButton btnMaskReview;
    }
}