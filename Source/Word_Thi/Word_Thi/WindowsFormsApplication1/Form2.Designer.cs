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
            this.textBoxUser = new TextBox();
            this.textBoxPass = new TextBox();
            this.buttonOK = new Button();
            this.label1 = new Label();
            this.label2 = new Label();
            this.comboBox1 = new ComboBox();
            this.label3 = new Label();
            this.label4 = new Label();
            this.label5 = new Label();
            this.openFileDialog1 = new OpenFileDialog();
            this.SuspendLayout();
            this.textBoxUser.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.textBoxUser.Location = new Point(140, 384);
            this.textBoxUser.Name = "textBoxUser";
            this.textBoxUser.Size = new Size(398, 38);
            this.textBoxUser.TabIndex = 1;
            this.textBoxUser.Text = "hocmosonline";
            this.textBoxPass.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.textBoxPass.Location = new Point(140, 430);
            this.textBoxPass.Name = "textBoxPass";
            this.textBoxPass.PasswordChar = '*';
            this.textBoxPass.Size = new Size(398, 38);
            this.textBoxPass.TabIndex = 2;
            this.textBoxPass.Text = "123456";
            this.buttonOK.Location = new Point(307, 513);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new Size(75, 23);
            this.buttonOK.TabIndex = 3;
            this.buttonOK.Text = "&OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new EventHandler(this.buttonOK_Click);
            this.label1.AutoSize = true;
            this.label1.Font = new Font("Times New Roman", 15.75f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.label1.ForeColor = Color.Black;
            this.label1.Location = new Point(37, 384);
            this.label1.Name = "label1";
            this.label1.Size = new Size(51, 24);
            this.label1.TabIndex = 5;
            this.label1.Text = "User";
            this.label2.AutoSize = true;
            this.label2.Font = new Font("Times New Roman", 15.75f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.label2.ForeColor = Color.Black;
            this.label2.Location = new Point(35, 437);
            this.label2.Name = "label2";
            this.label2.Size = new Size(100, 24);
            this.label2.TabIndex = 6;
            this.label2.Text = "PassWord";
            this.comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            this.comboBox1.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[2]
            {
        (object) "Tổng",
        (object) "Test"
            });
            this.comboBox1.Location = new Point(140, 476);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new Size(92, 39);
            this.comboBox1.TabIndex = 0;
            this.comboBox1.Visible = false;
            this.comboBox1.SelectedIndexChanged += new EventHandler(this.comboBox1_SelectedIndexChanged);
            this.label3.AutoSize = true;
            this.label3.Font = new Font("Times New Roman", 15.75f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.label3.Location = new Point(35, 490);
            this.label3.Name = "label3";
            this.label3.Size = new Size(79, 24);
            this.label3.TabIndex = 4;
            this.label3.Text = "Lesson:";
            this.label4.AutoSize = true;
            this.label4.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.label4.ForeColor = Color.White;
            this.label4.Location = new Point(120, 53);
            this.label4.Name = "label4";
            this.label4.Size = new Size(412, 31);
            this.label4.TabIndex = 7;
            this.label4.Text = "PHẦN LÀM TEST WORD 2019";
            this.label5.AutoSize = true;
            this.label5.Font = new Font("Microsoft Sans Serif", 15.75f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.label5.Location = new Point(135, 347);
            this.label5.Name = "label5";
            this.label5.Size = new Size(180, 25);
            this.label5.TabIndex = 8;
            this.label5.Text = "User: thimos1000";
            this.openFileDialog1.FileName = "openFileDialog1";
            this.AutoScaleDimensions = new SizeF(6f, 13f);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.BackColor = Color.FromArgb(64, 64, 64);
            //this.BackgroundImage = (Image)componentResourceManager.GetObject("$this.BackgroundImage");
            this.BackgroundImageLayout = ImageLayout.Center;
            this.ClientSize = new Size(698, 586);
            this.Controls.Add((Control)this.label5);
            this.Controls.Add((Control)this.label4);
            this.Controls.Add((Control)this.label3);
            this.Controls.Add((Control)this.comboBox1);
            this.Controls.Add((Control)this.label2);
            this.Controls.Add((Control)this.label1);
            this.Controls.Add((Control)this.buttonOK);
            this.Controls.Add((Control)this.textBoxPass);
            this.Controls.Add((Control)this.textBoxUser);
            this.DoubleBuffered = true;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.Name = "Form2";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "Login";
            this.Load += new EventHandler(this.Form2_Load);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
        private TextBox textBoxUser;
        private TextBox textBoxPass;
        private Button buttonOK;
        private Label label1;
        private Label label2;
        private ComboBox comboBox1;
        private Label label3;
        private Label label4;
        private Label label5;
        private OpenFileDialog openFileDialog1;
    }
}