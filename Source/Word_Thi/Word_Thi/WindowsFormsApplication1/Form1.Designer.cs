using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using WindowsFormsApplication1.Properties;

namespace WindowsFormsApplication1
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
            this.buttonReset = new Button();
            this.buttonCheck = new Button();
            this.buttonSubmit = new Button();
            this.buttonExit = new Button();
            this.label1 = new Label();
            this.timer1 = new Timer(this.components);
            this.tabControl1 = new TabControl();
            this.tabPage2 = new TabPage();
            this.panel2 = new Panel();
            this.pictureBox2 = new PictureBox();
            this.tabPage1 = new TabPage();
            this.panel1 = new Panel();
            this.pictureBox1 = new PictureBox();
            this.buttonSave = new Button();
            this.textBox1 = new TextBox();
            this.richTextBox1 = new RichTextBox();
            this.panel4 = new Panel();
            this.button1 = new Button();
            this.textBox2 = new TextBox();
            this.buttonHelp = new Button();
            this.buttonRefresh = new Button();
            this.buttonxhdh = new Button();
            this.buttonShowHide = new Button();
            this.button2 = new Button();
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
            this.checkedListBox1.Location = new Point(29, 30);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new Size(122, 172);
            this.checkedListBox1.TabIndex = 0;
            this.checkedListBox1.SelectedIndexChanged += new EventHandler(this.checkedListBox1_SelectedIndexChanged);
            this.buttonReset.Location = new Point(70, 266);
            this.buttonReset.Name = "buttonReset";
            this.buttonReset.Size = new Size(75, 23);
            this.buttonReset.TabIndex = 2;
            this.buttonReset.Text = "Reset";
            this.buttonReset.UseVisualStyleBackColor = true;
            this.buttonReset.Click += new EventHandler(this.button1_Click);
            this.buttonCheck.Location = new Point(70, 295);
            this.buttonCheck.Name = "buttonCheck";
            this.buttonCheck.Size = new Size(75, 23);
            this.buttonCheck.TabIndex = 3;
            this.buttonCheck.Text = "Check";
            this.buttonCheck.UseVisualStyleBackColor = true;
            this.buttonCheck.Click += new EventHandler(this.button2_Click);
            this.buttonSubmit.Location = new Point(76, 324);
            this.buttonSubmit.Name = "buttonSubmit";
            this.buttonSubmit.Size = new Size(75, 23);
            this.buttonSubmit.TabIndex = 4;
            this.buttonSubmit.Text = "Submit";
            this.buttonSubmit.UseVisualStyleBackColor = true;
            this.buttonSubmit.Click += new EventHandler(this.button3_Click);
            this.buttonExit.Location = new Point(70, 237);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new Size(75, 23);
            this.buttonExit.TabIndex = 5;
            this.buttonExit.Text = "E&xit";
            this.buttonExit.UseVisualStyleBackColor = true;
            this.buttonExit.Click += new EventHandler(this.button4_Click);
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.label1.Location = new Point(514, 351);
            this.label1.Name = "label1";
            this.label1.Size = new Size(100, 20);
            this.label1.TabIndex = 6;
            this.label1.Text = "00:00";
            this.label1.TextAlign = ContentAlignment.MiddleCenter;
            this.timer1.Tick += new EventHandler(this.timer1_Tick);
            this.tabControl1.Controls.Add((Control)this.tabPage2);
            this.tabControl1.Controls.Add((Control)this.tabPage1);
            this.tabControl1.Location = new Point(176, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new Size(529, 309);
            this.tabControl1.TabIndex = 8;
            this.tabPage2.Controls.Add((Control)this.panel2);
            this.tabPage2.Location = new Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new Padding(3);
            this.tabPage2.Size = new Size(521, 283);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "T A";
            this.tabPage2.UseVisualStyleBackColor = true;
            this.panel2.AutoScroll = true;
            this.panel2.Controls.Add((Control)this.pictureBox2);
            this.panel2.Location = new Point(3, 3);
            this.panel2.Name = "panel2";
            this.panel2.Size = new Size(512, 264);
            this.panel2.TabIndex = 0;
            this.pictureBox2.ErrorImage = (Image)null;
            this.pictureBox2.Location = new Point(3, 3);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new Size(1147, 3615);
            this.pictureBox2.SizeMode = PictureBoxSizeMode.AutoSize;
            this.pictureBox2.TabIndex = 0;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.MouseHover += new EventHandler(this.pictureBox2_MouseHover);
            this.tabPage1.Controls.Add((Control)this.panel1);
            this.tabPage1.Location = new Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new Padding(3);
            this.tabPage1.Size = new Size(521, 283);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "TV";
            this.tabPage1.UseVisualStyleBackColor = true;
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add((Control)this.pictureBox1);
            this.panel1.Location = new Point(6, 6);
            this.panel1.Name = "panel1";
            this.panel1.Size = new Size(509, 249);
            this.panel1.TabIndex = 0;
            this.pictureBox1.Image = (Image)Resources.E01;
            this.pictureBox1.Location = new Point(3, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new Size(881, 478);
            this.pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.MouseHover += new EventHandler(this.pictureBox1_MouseHover);
            this.buttonSave.Location = new Point(209, 348);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new Size(85, 23);
            this.buttonSave.TabIndex = 14;
            this.buttonSave.Text = "Lưu lai hỏi sau";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new EventHandler(this.buttonSave_Click);
            this.textBox1.Enabled = false;
            this.textBox1.Location = new Point(711, 210);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new Size(100, 20);
            this.textBox1.TabIndex = 13;
            this.textBox1.Visible = false;
            this.richTextBox1.BorderStyle = BorderStyle.None;
            this.richTextBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.richTextBox1.ForeColor = Color.Red;
            this.richTextBox1.Location = new Point(3, 5);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new Size(78, 89);
            this.richTextBox1.TabIndex = 16;
            this.richTextBox1.Text = "";
            this.panel4.AutoScroll = true;
            this.panel4.Controls.Add((Control)this.richTextBox1);
            this.panel4.Location = new Point(594, 351);
            this.panel4.Name = "panel4";
            this.panel4.Size = new Size(101, 94);
            this.panel4.TabIndex = 17;
            this.button1.Location = new Point(438, 348);
            this.button1.Name = "button1";
            this.button1.Size = new Size(29, 23);
            this.button1.TabIndex = 22;
            this.button1.Text = "Go";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new EventHandler(this.button1_Click_1);
            this.textBox2.Location = new Point(410, 351);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new Size(22, 20);
            this.textBox2.TabIndex = 21;
            this.textBox2.Text = "1";
            this.buttonHelp.Location = new Point(188, 377);
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Size = new Size(106, 23);
            this.buttonHelp.TabIndex = 23;
            this.buttonHelp.Text = "Video chỉ làm bài";
            this.buttonHelp.UseVisualStyleBackColor = true;
            this.buttonHelp.Click += new EventHandler(this.buttonHelp_Click);
            this.buttonRefresh.Location = new Point(209, 406);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new Size(52, 23);
            this.buttonRefresh.TabIndex = 24;
            this.buttonRefresh.Text = "Bố tri lại";
            this.buttonRefresh.UseVisualStyleBackColor = true;
            this.buttonRefresh.Click += new EventHandler(this.buttonRefresh_Click);
            this.buttonxhdh.Location = new Point(235, 319);
            this.buttonxhdh.Name = "buttonxhdh";
            this.buttonxhdh.Size = new Size(114, 23);
            this.buttonxhdh.TabIndex = 25;
            this.buttonxhdh.Text = "Hướng dẫn sử dụng";
            this.buttonxhdh.UseVisualStyleBackColor = true;
            this.buttonxhdh.Click += new EventHandler(this.buttonxhdh_Click);
            this.buttonShowHide.Location = new Point(209, 435);
            this.buttonShowHide.Name = "buttonShowHide";
            this.buttonShowHide.Size = new Size(63, 23);
            this.buttonShowHide.TabIndex = 27;
            this.buttonShowHide.Text = "Gọi Thầy";
            this.buttonShowHide.UseVisualStyleBackColor = true;
            this.buttonShowHide.Click += new EventHandler(this.buttonShowHide_Click);
            this.button2.Location = new Point(754, 351);
            this.button2.Name = "button2";
            this.button2.Size = new Size(41, 23);
            this.button2.TabIndex = 30;
            this.button2.Text = "Mở";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new EventHandler(this.button2_Click_1);
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new Point(492, 397);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new Size(44, 21);
            this.comboBox1.TabIndex = 31;
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
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new EventHandler(this.Form1_Load);
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
        private Button buttonReset;
        private Button buttonCheck;
        private Button buttonSubmit;
        private Button buttonExit;
        private Label label1;
        private Timer timer1;
        private TabControl tabControl1;
        private TabPage tabPage1;
        private TabPage tabPage2;
        private Panel panel2;
        private PictureBox pictureBox2;
        private Button buttonSave;
        private TextBox textBox1;
        private RichTextBox richTextBox1;
        private Panel panel4;
        private Panel panel1;
        private PictureBox pictureBox1;
        private Button button1;
        private TextBox textBox2;
        private Button buttonHelp;
        private Button buttonRefresh;
        private Button buttonxhdh;
        private Button buttonShowHide;
        private Button button2;
        private ComboBox comboBox1;
    }
}