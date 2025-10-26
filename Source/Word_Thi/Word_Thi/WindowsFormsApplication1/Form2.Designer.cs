using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Windows.Forms;
using MOS_WORD_TEST.Base;

namespace MOS_WORD_TEST
{
    partial class Form2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.cbExamList = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.radioVN = new System.Windows.Forms.RadioButton();
            this.radioEN = new System.Windows.Forms.RadioButton();
            this.btnCancel = new MOS_WORD_TEST.Base.RJButton();
            this.textBoxPass = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxUser = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.buttonOK = new MOS_WORD_TEST.Base.RJButton();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.richTextBox2 = new System.Windows.Forms.RichTextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // cbExamList
            // 
            this.cbExamList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbExamList.Font = new System.Drawing.Font("Calibri", 13F, System.Drawing.FontStyle.Bold);
            this.cbExamList.FormattingEnabled = true;
            this.cbExamList.Items.AddRange(new object[] {
            "Tổng",
            "Test"});
            this.cbExamList.Location = new System.Drawing.Point(90, 93);
            this.cbExamList.Name = "cbExamList";
            this.cbExamList.Size = new System.Drawing.Size(140, 29);
            this.cbExamList.TabIndex = 0;
            this.cbExamList.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 13F, System.Drawing.FontStyle.Bold);
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(8, 95);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 22);
            this.label3.TabIndex = 4;
            this.label3.Text = "Đề thi";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Font = new System.Drawing.Font("Calibri", 10F);
            this.tabControl1.ItemSize = new System.Drawing.Size(100, 25);
            this.tabControl1.Location = new System.Drawing.Point(12, 160);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.Padding = new System.Drawing.Point(10, 3);
            this.tabControl1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(273, 243);
            this.tabControl1.TabIndex = 14;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            this.tabPage1.Controls.Add(this.radioVN);
            this.tabPage1.Controls.Add(this.radioEN);
            this.tabPage1.Controls.Add(this.btnCancel);
            this.tabPage1.Controls.Add(this.textBoxPass);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.textBoxUser);
            this.tabPage1.Controls.Add(this.cbExamList);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.buttonOK);
            this.tabPage1.Font = new System.Drawing.Font("Calibri", 10F);
            this.tabPage1.Location = new System.Drawing.Point(4, 29);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(265, 210);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Đăng nhập";
            // 
            // radioVN
            // 
            this.radioVN.AutoSize = true;
            this.radioVN.Checked = true;
            this.radioVN.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.radioVN.Location = new System.Drawing.Point(96, 139);
            this.radioVN.Name = "radioVN";
            this.radioVN.Size = new System.Drawing.Size(43, 21);
            this.radioVN.TabIndex = 8;
            this.radioVN.TabStop = true;
            this.radioVN.Text = "VN";
            this.radioVN.UseVisualStyleBackColor = true;
            // 
            // radioEN
            // 
            this.radioEN.AutoSize = true;
            this.radioEN.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.radioEN.Location = new System.Drawing.Point(144, 139);
            this.radioEN.Name = "radioEN";
            this.radioEN.Size = new System.Drawing.Size(42, 21);
            this.radioEN.TabIndex = 8;
            this.radioEN.Text = "EN";
            this.radioEN.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.AutoSize = true;
            this.btnCancel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            this.btnCancel.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            this.btnCancel.BorderColor = System.Drawing.Color.White;
            this.btnCancel.BorderRadius = 5;
            this.btnCancel.BorderSize = 1;
            this.btnCancel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.Font = new System.Drawing.Font("Calibri", 12F);
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Location = new System.Drawing.Point(194, 165);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Padding = new System.Windows.Forms.Padding(2);
            this.btnCancel.Size = new System.Drawing.Size(62, 35);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "Thoát";
            this.btnCancel.TextColor = System.Drawing.Color.White;
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click_1);
            // 
            // textBoxPass
            // 
            this.textBoxPass.Cursor = System.Windows.Forms.Cursors.Hand;
            this.textBoxPass.Font = new System.Drawing.Font("Calibri", 13F, System.Drawing.FontStyle.Bold);
            this.textBoxPass.Location = new System.Drawing.Point(90, 49);
            this.textBoxPass.Name = "textBoxPass";
            this.textBoxPass.PasswordChar = '*';
            this.textBoxPass.Size = new System.Drawing.Size(166, 29);
            this.textBoxPass.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 13F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(6, 137);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 22);
            this.label1.TabIndex = 4;
            this.label1.Text = "Ngôn ngữ";
            // 
            // textBoxUser
            // 
            this.textBoxUser.Cursor = System.Windows.Forms.Cursors.Hand;
            this.textBoxUser.Font = new System.Drawing.Font("Calibri", 13F, System.Drawing.FontStyle.Bold);
            this.textBoxUser.Location = new System.Drawing.Point(90, 9);
            this.textBoxUser.Name = "textBoxUser";
            this.textBoxUser.Size = new System.Drawing.Size(166, 29);
            this.textBoxUser.TabIndex = 1;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label4.Font = new System.Drawing.Font("Calibri", 13F, System.Drawing.FontStyle.Bold);
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(8, 11);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 22);
            this.label4.TabIndex = 5;
            this.label4.Text = "Tài khoản";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label5.Font = new System.Drawing.Font("Calibri", 13F, System.Drawing.FontStyle.Bold);
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(8, 51);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 22);
            this.label5.TabIndex = 6;
            this.label5.Text = "Mật khẩu";
            // 
            // buttonOK
            // 
            this.buttonOK.AutoSize = true;
            this.buttonOK.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonOK.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            this.buttonOK.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            this.buttonOK.BorderColor = System.Drawing.Color.White;
            this.buttonOK.BorderRadius = 5;
            this.buttonOK.BorderSize = 1;
            this.buttonOK.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonOK.Font = new System.Drawing.Font("Calibri", 12F);
            this.buttonOK.ForeColor = System.Drawing.Color.White;
            this.buttonOK.Location = new System.Drawing.Point(90, 165);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Padding = new System.Windows.Forms.Padding(2);
            this.buttonOK.Size = new System.Drawing.Size(95, 35);
            this.buttonOK.TabIndex = 3;
            this.buttonOK.Text = "Đăng nhập";
            this.buttonOK.TextColor = System.Drawing.Color.White;
            this.buttonOK.UseVisualStyleBackColor = false;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154)))));
            this.tabPage2.Controls.Add(this.richTextBox1);
            this.tabPage2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.tabPage2.Font = new System.Drawing.Font("Calibri", 10F);
            this.tabPage2.ForeColor = System.Drawing.Color.Black;
            this.tabPage2.Location = new System.Drawing.Point(4, 29);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Size = new System.Drawing.Size(265, 210);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Hướng dẫn tài khoản";
            // 
            // richTextBox1
            // 
            this.richTextBox1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.richTextBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBox1.Font = new System.Drawing.Font("Calibri", 10F);
            this.richTextBox1.ForeColor = System.Drawing.Color.Black;
            this.richTextBox1.Location = new System.Drawing.Point(0, 0);
            this.richTextBox1.Margin = new System.Windows.Forms.Padding(0);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(265, 210);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = "Nếu bạn chưa có mật khẩu để đăng nhập, vui lòng gửi ID bên dưới cho Admin để được" +
    " cấp:\n\nAABBCCDDEEFF\n\nWebsite hỗ trợ: https://mos360.vn";
            // 
            // richTextBox2
            // 
            this.richTextBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            this.richTextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBox2.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Italic);
            this.richTextBox2.ForeColor = System.Drawing.Color.White;
            this.richTextBox2.Location = new System.Drawing.Point(12, 412);
            this.richTextBox2.Name = "richTextBox2";
            this.richTextBox2.ReadOnly = true;
            this.richTextBox2.Size = new System.Drawing.Size(273, 62);
            this.richTextBox2.TabIndex = 16;
            this.richTextBox2.Text = "Thành công không đến từ may mắn, mà bắt đầu từ từng giờ bạn ngồi học hôm nay.\n\nWe" +
    "bsite hỗ trợ: https://mos360.vn";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::MOS_WORD_TEST.Properties.Resources.w_mos_s;
            this.pictureBox1.Location = new System.Drawing.Point(44, 20);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(210, 119);
            this.pictureBox1.TabIndex = 17;
            this.pictureBox1.TabStop = false;
            // 
            // Form2
            // 
            this.AccessibleDescription = "MOS360_WORD_TEST";
            this.AccessibleName = "MOS360_WORD_TEST";
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ClientSize = new System.Drawing.Size(296, 487);
            this.ControlBox = false;
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.richTextBox2);
            this.Controls.Add(this.tabControl1);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.Form2_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private Base.RJButton buttonOK;
        private Base.RJButton btnCancel;
        private ComboBox cbExamList;
        private Label label3;
        private OpenFileDialog openFileDialog1;
        private TabControl tabControl1;
        private TabPage tabPage1;
        private TextBox textBoxPass;
        private TextBox textBoxUser;
        private Label label4;
        private Label label5;
        private TabPage tabPage2;
        private RichTextBox richTextBox1;
        private RichTextBox richTextBox2;
        private PictureBox pictureBox1;
        private RadioButton radioVN;
        private RadioButton radioEN;
        private Label label1;
    }
}