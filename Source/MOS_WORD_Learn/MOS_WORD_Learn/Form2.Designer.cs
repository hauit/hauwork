using System;
using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApplication1
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
            this.label3 = new Label();
            this.label4 = new Label();
            this.openFileDialog1 = new OpenFileDialog();
            this.textBoxPassSo = new TextBox();
            this.textBox1 = new TextBox();
            this.SuspendLayout();
            this.textBoxUser.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.textBoxUser.Location = new Point(154, 336);
            this.textBoxUser.Name = "textBoxUser";
            this.textBoxUser.Size = new Size(398, 38);
            this.textBoxUser.TabIndex = 1;
            this.textBoxUser.Text = "On Luyen MOS";
            this.textBoxPass.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.textBoxPass.Location = new Point(154, 386);
            this.textBoxPass.Name = "textBoxPass";
            this.textBoxPass.PasswordChar = '*';
            this.textBoxPass.Size = new Size(398, 38);
            this.textBoxPass.TabIndex = 2;
            this.buttonOK.Location = new Point(308, 446);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new Size(75, 23);
            this.buttonOK.TabIndex = 3;
            this.buttonOK.Text = "&OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new EventHandler(this.buttonOK_Click);
            this.label1.AutoSize = true;
            this.label1.Font = new Font("Times New Roman", 15.75f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.label1.ForeColor = Color.Black;
            this.label1.Location = new Point(49, 336);
            this.label1.Name = "label1";
            this.label1.Size = new Size(51, 24);
            this.label1.TabIndex = 5;
            this.label1.Text = "User";
            this.label2.AutoSize = true;
            this.label2.Font = new Font("Times New Roman", 15.75f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.label2.ForeColor = Color.Black;
            this.label2.Location = new Point(49, 394);
            this.label2.Name = "label2";
            this.label2.Size = new Size(100, 24);
            this.label2.TabIndex = 6;
            this.label2.Text = "PassWord";
            this.label3.AutoSize = true;
            this.label3.Font = new Font("Times New Roman", 15.75f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.label3.ForeColor = Color.Red;
            this.label3.Location = new Point(179, 37);
            this.label3.Name = "label3";
            this.label3.Size = new Size(145, 24);
            this.label3.TabIndex = 4;
            this.label3.Text = "Pass Của Ngày";
            this.label4.AutoSize = true;
            this.label4.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.label4.ForeColor = Color.White;
            this.label4.Location = new Point(236, 77);
            this.label4.Name = "label4";
            this.label4.Size = new Size(226, 31);
            this.label4.TabIndex = 8;
            this.label4.Text = "PM WORD 2019";
            this.openFileDialog1.FileName = "openFileDialog1";
            this.textBoxPassSo.Font = new Font("Microsoft Sans Serif", 15.75f, FontStyle.Regular, GraphicsUnit.Point, (byte)163);
            this.textBoxPassSo.ForeColor = Color.Red;
            this.textBoxPassSo.Location = new Point(330, 33);
            this.textBoxPassSo.Name = "textBoxPassSo";
            this.textBoxPassSo.ReadOnly = true;
            this.textBoxPassSo.Size = new Size(185, 31);
            this.textBoxPassSo.TabIndex = 10;
            this.textBox1.Location = new Point(183, 126);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new Size(318, 20);
            this.textBox1.TabIndex = 11;
            this.textBox1.Text = "https://meet.google.com/uay-ywnp-xzs";
            this.AutoScaleDimensions = new SizeF(6f, 13f);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.BackColor = Color.FromArgb(64, 64, 64);
            //this.BackgroundImage = (Image)componentResourceManager.GetObject("$this.BackgroundImage");
            this.BackgroundImageLayout = ImageLayout.Center;
            this.ClientSize = new Size(698, 493);
            this.Controls.Add((Control)this.textBox1);
            this.Controls.Add((Control)this.textBoxPassSo);
            this.Controls.Add((Control)this.label4);
            this.Controls.Add((Control)this.label3);
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
        private Label label3;
        private Label label4;
        private OpenFileDialog openFileDialog1;
        private TextBox textBoxPassSo;
        private TextBox textBox1;
    }
}