// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Form1
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using Point = System.Drawing.Point;


namespace MOS_PPT_LEARN
{

    public class Form1_Cu : Form
    {
        private IContainer components = (IContainer)null;
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
        private Button buttonZoom;
        private Button buttonSummary;
        private Button buttonVideoHelp;
        private Button button2;
        private ComboBox comboBoxCauDaDanhDau;
        private Button buttonHelpVideo;
        private Button buttonHelp;
        private int zoom = 0;
        private bool check = true;
        private int so_cau_dung = 0;
        private int so_cau_sai = 0;
        private DateTime currenTime;
        private int screen_height;
        private int screen_width;
        private Microsoft.Office.Interop.PowerPoint.Application a;
        private Presentation d;
        private int tong_so_cau = School.Tong();
        private int cau_User = 1;
        private bool chotat = false;
        public object readOnly = (object)false;
        public object isVisible = (object)true;
        public object missing = (object)Missing.Value;
        private enviroment paramater;
        private Help h;
        private int[] Diem;

        protected override void Dispose(bool disposing)
        {
            if (disposing && this.components != null)
                this.components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.buttonExit = new Button();
            this.buttonNext = new Button();
            this.buttonReset = new Button();
            this.buttonEV = new Button();
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
            this.buttonHelp = new Button();
            ((ISupportInitialize)this.pictureBox1).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            this.buttonExit.Location = new Point(100, 68);
            this.buttonExit.Name = "buttonExit";
            this.buttonExit.Size = new Size(75, 23);
            this.buttonExit.TabIndex = 0;
            this.buttonExit.Text = "Exit";
            this.buttonExit.UseVisualStyleBackColor = true;
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
            this.buttonRefresh.Location = new Point(256 /*0x0100*/, 210);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new Size(87, 23);
            this.buttonRefresh.TabIndex = 5;
            this.buttonRefresh.Text = "Phục hồi PPT";
            this.buttonRefresh.UseVisualStyleBackColor = true;
            this.buttonRefresh.Click += new EventHandler(this.buttonRefresh_Click);
            this.richTextBox1.Font = new System.Drawing.Font("Times New Roman", 18f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.richTextBox1.Location = new Point(61, 296);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new Size(103, 82);
            this.richTextBox1.TabIndex = 6;
            this.richTextBox1.Text = "";
            this.buttonCheck.BackColor = Color.FromArgb((int)byte.MaxValue, 192 /*0xC0*/, 128 /*0x80*/);
            this.buttonCheck.Location = new Point(391, 184);
            this.buttonCheck.Name = "buttonCheck";
            this.buttonCheck.Size = new Size(113, 23);
            this.buttonCheck.TabIndex = 7;
            this.buttonCheck.Text = "Check đúng/ sai";
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
            this.labelKQ.Location = new Point(606, 288);
            this.labelKQ.Name = "labelKQ";
            this.labelKQ.Size = new Size(0, 21);
            this.labelKQ.TabIndex = 10;
            this.labelCauHienTai.AutoSize = true;
            this.labelCauHienTai.Font = new System.Drawing.Font("Times New Roman", 14.25f, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.labelCauHienTai.ForeColor = Color.Yellow;
            this.labelCauHienTai.Location = new Point(587, 217);
            this.labelCauHienTai.Name = "labelCauHienTai";
            this.labelCauHienTai.Size = new Size(79, 22);
            this.labelCauHienTai.TabIndex = 12;
            this.labelCauHienTai.Text = "Câu 120";
            this.comboBoxCauNext.BackColor = SystemColors.MenuText;
            this.comboBoxCauNext.Font = new System.Drawing.Font("Times New Roman", 14.25f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.comboBoxCauNext.ForeColor = Color.Blue;
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
            this.buttonZoom.BackColor = Color.FromArgb(128 /*0x80*/, (int)byte.MaxValue, (int)byte.MaxValue);
            this.buttonZoom.Location = new Point(352, 291);
            this.buttonZoom.Name = "buttonZoom";
            this.buttonZoom.Size = new Size(72, 23);
            this.buttonZoom.TabIndex = 22;
            this.buttonZoom.Text = "Đánh Dấu";
            this.buttonZoom.UseVisualStyleBackColor = false;
            this.buttonZoom.Click += new EventHandler(this.buttonZoom_Click);
            this.buttonSummary.Location = new Point(320, 341);
            this.buttonSummary.Name = "buttonSummary";
            this.buttonSummary.Size = new Size(41, 23);
            this.buttonSummary.TabIndex = 21;
            this.buttonSummary.Text = "Xóa";
            this.buttonSummary.UseVisualStyleBackColor = true;
            this.buttonSummary.Click += new EventHandler(this.buttonSummary_Click);
            this.buttonVideoHelp.BackColor = Color.FromArgb((int)byte.MaxValue, (int)byte.MaxValue, 128 /*0x80*/);
            this.buttonVideoHelp.Location = new Point(471, 288);
            this.buttonVideoHelp.Name = "buttonVideoHelp";
            this.buttonVideoHelp.Size = new Size(114, 23);
            this.buttonVideoHelp.TabIndex = 23;
            this.buttonVideoHelp.Text = "Hướng dẫn sử dụng";
            this.buttonVideoHelp.UseVisualStyleBackColor = false;
            this.buttonVideoHelp.Click += new EventHandler(this.buttonVideoHelp_Click);
            this.button2.Location = new Point(483, 381);
            this.button2.Name = "button2";
            this.button2.Size = new Size(93, 23);
            this.button2.TabIndex = 24;
            this.button2.Text = "Gọi Thầy";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new EventHandler(this.buttonTopMost_Click);
            this.comboBoxCauDaDanhDau.BackColor = SystemColors.MenuText;
            this.comboBoxCauDaDanhDau.Font = new System.Drawing.Font("Times New Roman", 14.25f, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.comboBoxCauDaDanhDau.ForeColor = Color.Blue;
            this.comboBoxCauDaDanhDau.FormattingEnabled = true;
            this.comboBoxCauDaDanhDau.Location = new Point(845, 288);
            this.comboBoxCauDaDanhDau.Name = "comboBoxCauDaDanhDau";
            this.comboBoxCauDaDanhDau.Size = new Size(52, 29);
            this.comboBoxCauDaDanhDau.TabIndex = 27;
            this.buttonHelpVideo.BackColor = Color.Fuchsia;
            this.buttonHelpVideo.Location = new Point(216, 394);
            this.buttonHelpVideo.Name = "buttonHelpVideo";
            this.buttonHelpVideo.Size = new Size(73, 23);
            this.buttonHelpVideo.TabIndex = 28;
            this.buttonHelpVideo.Text = "Chỉ làm bài";
            this.buttonHelpVideo.UseVisualStyleBackColor = false;
            this.buttonHelpVideo.Click += new EventHandler(this.buttonHelpVideo_Click);
            this.buttonHelp.Location = new Point(390, 486);
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Size = new Size(114, 23);
            this.buttonHelp.TabIndex = 31 /*0x1F*/;
            this.buttonHelp.Text = "Về chế độ thường";
            this.buttonHelp.UseVisualStyleBackColor = true;
            this.buttonHelp.Click += new EventHandler(this.buttonHelp_Click_1);
            this.AutoScaleDimensions = new SizeF(6f, 13f);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.BackColor = Color.Black;
            this.ClientSize = new Size(1016, 599);
            this.Controls.Add((Control)this.buttonHelp);
            this.Controls.Add((Control)this.buttonHelpVideo);
            this.Controls.Add((Control)this.labelKQ);
            this.Controls.Add((Control)this.comboBoxCauDaDanhDau);
            this.Controls.Add((Control)this.button2);
            this.Controls.Add((Control)this.buttonVideoHelp);
            this.Controls.Add((Control)this.buttonZoom);
            this.Controls.Add((Control)this.panel1);
            this.Controls.Add((Control)this.comboBoxCauNext);
            this.Controls.Add((Control)this.labelCauHienTai);
            this.Controls.Add((Control)this.buttonCheck);
            this.Controls.Add((Control)this.richTextBox1);
            this.Controls.Add((Control)this.buttonRefresh);
            this.Controls.Add((Control)this.buttonEV);
            this.Controls.Add((Control)this.buttonReset);
            this.Controls.Add((Control)this.buttonNext);
            this.Controls.Add((Control)this.buttonExit);
            this.Controls.Add((Control)this.buttonSummary);
            this.FormBorderStyle = FormBorderStyle.None;
            this.Name = nameof(Form1_Cu);
            this.Text = nameof(Form1_Cu);
            this.TopMost = true;
            this.WindowState = FormWindowState.Maximized;
            this.Load += new EventHandler(this.Form1_Load);
            ((ISupportInitialize)this.pictureBox1).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void setDiem()
        {
            for (int index = 0; index < this.Diem.Length; ++index)
                this.Diem[index] = 0;
        }

        private void LuuDiem()
        {
            string str1 = "";
            int num = 0;
            for (int index = 0; index < this.Diem.Length; ++index)
            {
                num += this.Diem[index];
                if (this.Diem[index] == 1)
                    str1 = $"{str1}{(object)(index + 1)},";
            }
            string str2 = $"{(object)num}Câu,{str1}";
            string str3 = $"<p>{DateTime.Now.ToShortDateString()},{str2}" + "</p>";
            if (num <= 1)
                return;
            string str4 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Diem");
            if (!Directory.Exists(str4))
                Directory.CreateDirectory(str4);
            string path = Path.Combine(str4, "hoc_PPT.html");
            if (File.Exists(path))
            {
                TextWriter textWriter = (TextWriter)new StreamWriter(path, true);
                textWriter.WriteLine(str3);
                textWriter.Close();
            }
            else
            {
                TextWriter textWriter = (TextWriter)new StreamWriter(path, true);
                string str5 = "<P style=\"text-align:center;color:red;font-size:160%\"> Điều Kiện để được thi lại miễn phí là phải học hành chăm chỉ, gửi điểm thường xuyên cho thầy</p><P style=\"text-align:center;\">Khi Có Điểm bạn copy Điểm này gửi cho thầy qua 1 trong các face sau: </p> <P style=\"text-align:center;\">================================================================== </p> <P style=\"text-align:center;\"><a href=\"https://www.facebook.com/tinhoc.hoangkha.17\" target=\"_blank\">1. Minh Quang Vũ </a>  </p><P style=\"text-align:center;\"><a href=\"https://www.facebook.com/truongthinh.19\" target=\"_blank\">2. Trường Thịnh</a>  </p> <P style=\"text-align:center;\"><a href=\"https://www.facebook.com/tinhocmos900\" target=\"_blank\">3. Nguyễn Cữu Đàm</a>  </p><P style=\"text-align:center;\">================================================================== </p> ";
                textWriter.WriteLine(str5);
                textWriter.WriteLine(str3);
                textWriter.Close();
            }
            TextWriter textWriter1 = (TextWriter)new StreamWriter("C:\\MOS\\GOC HO TRO\\ngayhocgannhat", true);
            TextWriter textWriter2 = textWriter1;
            object[] objArray1 = new object[5]
            {
              (object) DateTime.Now.Day,
              (object) "|",
              null,
              null,
              null
            };
            object[] objArray2 = objArray1;
            DateTime now = DateTime.Now;
            // ISSUE: variable of a boxed type
            ValueType month = (ValueType)now.Month;
            objArray2[2] = (object)month;
            objArray1[3] = (object)"|";
            object[] objArray3 = objArray1;
            now = DateTime.Now;
            // ISSUE: variable of a boxed type
            ValueType year = (ValueType)now.Year;
            objArray3[4] = (object)year;
            string str6 = string.Concat(objArray1);
            textWriter2.WriteLine(str6);
            textWriter1.Close();
        }

        private void LuuDanhDau()
        {
            TextWriter textWriter = (TextWriter)new StreamWriter(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip//luudanhdau"));
            for (int index = 0; index < this.comboBoxCauDaDanhDau.Items.Count; ++index)
                textWriter.WriteLine(this.comboBoxCauDaDanhDau.Items[index].ToString());
            textWriter.Close();
        }

        private void loadDanhDau()
        {
            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "zip//luudanhdau");
            if (!File.Exists(path))
                return;
            TextReader textReader = (TextReader)new StreamReader(path);
            for (string str = textReader.ReadLine(); str != null; str = textReader.ReadLine())
                this.comboBoxCauDaDanhDau.Items.Add((object)str);
            textReader.Close();
        }

        public Form1_Cu()
        {
            this.h = new Help();
            this.InitializeComponent();
            this.FormClosed += new FormClosedEventHandler(this.Form1_FormClosed);
            this.a = (Microsoft.Office.Interop.PowerPoint.Application)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("91493441-5A91-11CF-8700-00AA0060263B")));
            this.load_cau_hoi(this.cau_User);
            // ISSUE: method pointer
            // ISSUE: object of a compiler-generated type is created
            this.a.PresentationBeforeClose += this.a_PresentationBeforeClose;
            // ISSUE: method pointer
            // ISSUE: object of a compiler-generated type is created
            this.a.ProtectedViewWindowBeforeClose += this.a_ProtectedViewWindowBeforeClose;
        }

        private void a_ProtectedViewWindowBeforeClose(
          ProtectedViewWindow ProtViewWindow,
          PpProtectedViewCloseReason ProtectedViewCloseReason,
          ref bool Cancel)
        {
            if (this.chotat)
                return;
            Cancel = true;
        }

        private void a_PresentationBeforeClose(Presentation Pres, ref bool Cancel)
        {
            if (this.chotat)
                return;
            Cancel = true;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.LuuDiem();
            this.LuuDanhDau();
            try
            {
                this.chotat = true;
                while (this.a.Windows.Count >= 1)
                {
                    // ISSUE: reference to a compiler-generated method
                    this.a.Windows[1].Close();
                }
                if (this.a != null)
                {
                    // ISSUE: reference to a compiler-generated method
                    this.a.Quit();
                }
                this.chotat = false;
            }
            catch (Exception ex)
            {
            }
            foreach (string file in Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "Word")))
                File.Delete(file);
            this.pictureBox1.Image.Dispose();
            foreach (string file in Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "tam")))
                File.Delete(file);
            this.Close();
        }

        private void turnOff()
        {
            this.chotat = true;
            while (this.a.Windows.Count >= 1)
            {
                // ISSUE: reference to a compiler-generated method
                this.a.Windows[1].Close();
            }
            while (this.a.ProtectedViewWindows.Count >= 1)
            {
                // ISSUE: reference to a compiler-generated method
                this.a.ProtectedViewWindows[1].Close();
            }
            this.chotat = false;
            this.pictureBox1.Image.Dispose();
            foreach (string file in Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "tam")))
                File.Delete(file);
            foreach (string file in Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "Word")))
                File.Delete(file);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (this.cau_User == this.tong_so_cau)
            {
                int num = (int)MessageBox.Show($"Điểm của bạn là:{(this.so_cau_dung * 1000 / this.tong_so_cau).ToString()}, số câu dúng:{this.so_cau_dung.ToString()}, số câu sai:{this.so_cau_sai.ToString()}");
            }
            try
            {
                this.turnOff();
            }
            catch (Exception ex)
            {
            }
            this.cau_User = int.Parse(this.comboBoxCauNext.Text.Trim());
            this.load_cau_hoi(this.cau_User);
            this.labelCauHienTai.Text = "Câu: " + this.cau_User.ToString();
            if (this.cau_User < this.tong_so_cau)
                this.comboBoxCauNext.Text = (this.cau_User + 1).ToString();
            else
                this.comboBoxCauNext.Text = "1";
            this.labelKQ.Text = "";
            this.check = true;
            this.buttonEV.Text = "Tiếng Việt";
        }

        private void loadTongSoCau()
        {
            for (int index = 0; index < School.Mn.Length; ++index)
            {
                string[] files = Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, $"Data\\Sec_{(object)index}\\hinh\\E"));
                School.Mn[index] = files.Length;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.loadDanhDau();
            try
            {
                this.loadTongSoCau();
                this.tong_so_cau = School.Tong();
                this.Diem = new int[this.tong_so_cau];
                int y = this.comboBoxCauNext.Height <= this.buttonExit.Height ? (this.buttonExit.Height - this.comboBoxCauNext.Height) / 2 : (this.comboBoxCauNext.Height - this.buttonExit.Height) / 2;
                this.screen_height = this.Height;
                this.screen_width = this.Width;
                this.WindowState = FormWindowState.Normal;
                this.Width = this.screen_width;
                this.Height = this.screen_height / 5;
                this.Location = new Point(0, this.screen_height * 4 / 5);
                this.buttonExit.Location = new Point(0, y);
                this.buttonRefresh.Location = new Point(this.buttonExit.Width, y);
                this.buttonHelp.Location = new Point(this.buttonRefresh.Location.X + this.buttonRefresh.Width, y);
                this.buttonHelpVideo.Location = new Point(this.buttonHelp.Location.X + this.buttonHelp.Width, y);
                this.buttonCheck.Location = new Point(this.buttonHelpVideo.Location.X + this.buttonHelpVideo.Width, y);
                this.labelKQ.Location = new Point(this.buttonCheck.Location.X + this.buttonCheck.Width, y);
                this.buttonEV.Location = new Point(this.Width - this.buttonEV.Width, y);
                this.comboBoxCauNext.Location = new Point(this.Width - this.buttonEV.Width - this.comboBoxCauNext.Width, 0);
                this.buttonNext.Location = new Point(this.Width - this.buttonEV.Width - this.comboBoxCauNext.Width - this.buttonNext.Width, y);
                this.buttonReset.Location = new Point(this.buttonNext.Location.X - this.buttonReset.Width, y);
                this.labelCauHienTai.Location = new Point(this.buttonReset.Location.X - this.labelCauHienTai.Width, y);
                this.buttonSummary.Location = new Point(this.buttonReset.Location.X - this.labelCauHienTai.Width - this.buttonZoom.Width, y);
                this.comboBoxCauDaDanhDau.Location = new Point(this.buttonReset.Location.X - this.comboBoxCauDaDanhDau.Width - this.labelCauHienTai.Width - this.buttonZoom.Width, 0);
                this.buttonZoom.Location = new Point(this.buttonReset.Location.X - this.comboBoxCauDaDanhDau.Width - this.labelCauHienTai.Width - this.buttonZoom.Width * 2, y);
                this.button2.Location = new Point(this.buttonReset.Location.X - this.comboBoxCauDaDanhDau.Width - this.labelCauHienTai.Width - this.buttonZoom.Width * 2 - this.button2.Width, y);
                this.buttonVideoHelp.Location = new Point(this.buttonReset.Location.X - this.buttonVideoHelp.Width - this.comboBoxCauDaDanhDau.Width - this.labelCauHienTai.Width - this.buttonZoom.Width * 2 - this.button2.Width, y);
                this.richTextBox1.Location = new Point(0, this.comboBoxCauNext.Height);
                this.richTextBox1.Width = this.Width;
                this.richTextBox1.Height = this.Height - this.comboBoxCauNext.Height;
                this.panel1.Location = new Point(0, this.comboBoxCauNext.Height);
                this.panel1.Width = this.Width;
                this.panel1.Height = this.Height - this.comboBoxCauNext.Height;
                this.pictureBox1.Location = new Point(0, 0);
                this.pictureBox1.Width = this.Width - this.buttonEV.Width / 3;
                this.comboBoxCauNext.Text = (this.cau_User + 1).ToString();
                this.labelCauHienTai.Text = "Câu: " + this.cau_User.ToString();
                for (int index = 0; index < this.tong_so_cau; ++index)
                    this.comboBoxCauNext.Items.Add((object)(index + 1));
                this.Copy(Path.Combine(System.Windows.Forms.Application.StartupPath, "data\\Doc"), Environment.GetFolderPath(Environment.SpecialFolder.Personal));
                this.a.Visible = MsoTriState.msoTrue;
                this.a.WindowState = PpWindowState.ppWindowNormal;
                this.a.Top = 0.0f;
                this.a.Left = 0.0f;
                this.a.Width = (float)this.screen_width;
                this.a.Height = (float)(this.screen_height * 3 / 5);
                this.TopMost = true;
                this.currenTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
        }

        private void Copy(string sourceDir, string targetDir)
        {
            foreach (string file in Directory.GetFiles(sourceDir))
                File.Copy(file, Path.Combine(targetDir, Path.GetFileName(file)), true);
        }

        private void buttonCheck_Click(object sender, EventArgs e)
        {
            this.labelKQ.Text = this.CheckResult(this.paramater);
            if (this.labelKQ.Text == "True")
                this.Diem[this.cau_User - 1] = 1;
            this.check = false;
        }

        private enviroment Getparmater(int cauUser)
        {
            enviroment enviroment = new enviroment();
            int num1 = 0;
            for (int index = 0; index < School.Mn.Length; ++index)
            {
                int num2 = num1 + School.Mn[index];
                if (cauUser <= num2)
                {
                    enviroment.quesion = cauUser - num1;
                    enviroment.section = index;
                    enviroment.dirPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "data\\sec_" + index.ToString());
                    enviroment.Source_de_En = Path.Combine(enviroment.dirPath, "hinh\\E\\" + enviroment.quesion.ToString());
                    enviroment.Source_de_Vn = Path.Combine(enviroment.dirPath, "hinh\\V\\" + enviroment.quesion.ToString());
                    enviroment.Source_file_word_path = Path.Combine(enviroment.dirPath, "file\\" + enviroment.quesion.ToString());
                    enviroment.Source_file_help_path = Path.Combine(enviroment.dirPath, "help\\" + enviroment.quesion.ToString());
                    enviroment.Source_file_help_video_path = Path.Combine(enviroment.dirPath, "Vhelp\\" + enviroment.quesion.ToString());
                    break;
                }
                num1 = num2;
            }
            return enviroment;
        }

        private void load_cau_hoi(int cau_hoi_so)
        {
            this.paramater = this.Getparmater(cau_hoi_so);
            this.setDefalt(this.paramater);
            this.paramater.Dest_file_Word_Name = Path.Combine(System.Windows.Forms.Application.StartupPath, $"Word\\{this.paramater.section.ToString()}_{this.paramater.quesion.ToString()}.docx");
            this.paramater.DeTiengAnh = Path.Combine(System.Windows.Forms.Application.StartupPath, $"tam\\{cau_hoi_so.ToString()}E");
            this.paramater.DeTiengViet = Path.Combine(System.Windows.Forms.Application.StartupPath, $"tam\\{cau_hoi_so.ToString()}V");
            this.paramater.Dest_file_help_Name = Path.Combine(System.Windows.Forms.Application.StartupPath, "tam\\help");
            this.paramater.Dest_file_help_video_Name = Path.Combine(System.Windows.Forms.Application.StartupPath, "tam\\Vhelp");
            Home.DecryptFile(this.paramater.Source_file_word_path, this.paramater.Dest_file_Word_Name.ToString());
            Home.DecryptFile(this.paramater.Source_de_En, this.paramater.DeTiengAnh);
            Home.DecryptFile(this.paramater.Source_de_Vn, this.paramater.DeTiengViet);
            if (File.Exists(this.paramater.Source_file_help_video_path))
                Home.DecryptFile(this.paramater.Source_file_help_video_path, this.paramater.Dest_file_help_video_Name);
            this.pictureBox1.Image = Image.FromFile(this.paramater.DeTiengAnh);
            this.pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
            this.pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            this.pictureBox1.Height = this.pictureBox1.Height * this.Width / this.pictureBox1.Width;
            this.pictureBox1.Width = this.Width - this.buttonCheck.Width / 2;
            this.pictureBox1.Refresh();
            // ISSUE: reference to a compiler-generated method
            this.d = this.a.Presentations.Open(this.paramater.Dest_file_Word_Name);
            this.h.refresh();
        }

        private void setDefalt(enviroment par)
        {
        }

        private string CheckResult(enviroment par) => this.Check(par);

        private string Check(enviroment par)
        {
            string str = "";
            try
            {
                switch (par.section)
                {
                    case 0:
                        str = S0_SlideMaster.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 1:
                        str = S1_Illustrator.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 2:
                        str = S2_Chart.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 3:
                        str = S3_AddSlide.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 4:
                        str = S4_Layout.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 5:
                        str = S5_Text.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 6:
                        str = S6_VideoAudio.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 7:
                        str = S7_Animation.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 8:
                        str = S8_Transition.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 9:
                        str = S9_Section.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 10:
                        str = S10_Final_Steps.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 11:
                        str = S11.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 12:
                        str = S12.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 13:
                        str = S13.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 14:
                        str = S14.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 15:
                        str = S15.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 16 /*0x10*/:
                        str = S16.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 17:
                        str = S17.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    case 18:
                        str = CheckWord.CheckCau(par.quesion, this.a, this.d).ToString();
                        break;
                    default:
                        str = "have problem";
                        break;
                }
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Đống tất cả các hộp thoại đang mở trước");
            }
            return str;
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                this.a.WindowState = PpWindowState.ppWindowNormal;
                this.a.Top = 0.0f;
                this.a.Left = 0.0f;
                this.a.Width = (float)this.screen_width;
                this.a.Height = (float)(this.screen_height * 3 / 5);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Bạn đã tắt MS Word! Tắt chương trình và làm lại");
            }
        }

        private void buttonReset_Click(object sender, EventArgs e)
        {
            int num1 = 0;
            try
            {
                this.turnOff();
            }
            catch (Exception ex)
            {
                int num2 = (int)MessageBox.Show("close các hộp thoại trước khi reset");
                num1 = 1;
            }
            if (num1 == 1)
                return;
            this.load_cau_hoi(this.cau_User);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            TimeSpan timeSpan = new TimeSpan(0, 0, 0) + (DateTime.Now - this.currenTime);
        }

        private void buttonEV_Click(object sender, EventArgs e)
        {
            if (this.buttonEV.Text == "Tiếng Việt")
            {
                this.buttonEV.Text = "Tiếng Anh";
                this.pictureBox1.Image.Dispose();
                this.pictureBox1.Image = Image.FromFile(this.paramater.DeTiengViet);
            }
            else
            {
                this.buttonEV.Text = "Tiếng Việt";
                this.pictureBox1.Image.Dispose();
                this.pictureBox1.Image = Image.FromFile(this.paramater.DeTiengAnh);
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
        }

        private void pictureBox1_MouseHover(object sender, EventArgs e) => this.panel1.Focus();

        private void buttonHelp_Click(object sender, EventArgs e)
        {
            this.h.refresh();
            this.h.Show();
        }

        private void buttonZoom_Click(object sender, EventArgs e)
        {
            this.comboBoxCauDaDanhDau.Items.Add((object)this.cau_User);
        }

        private void buttonSummary_Click(object sender, EventArgs e)
        {
            int selectedIndex = this.comboBoxCauDaDanhDau.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < this.comboBoxCauDaDanhDau.Items.Count)
            {
                this.comboBoxCauDaDanhDau.Items.RemoveAt(selectedIndex);
            }
            else
            {
                int num = (int)MessageBox.Show("Chọn câu đã đánh dấu cần xóa");
            }
        }

        private void buttonVideoHelp_Click(object sender, EventArgs e)
        {
            if (this.buttonVideoHelp.Text == "Hướng dẫn sử dụng")
            {
                this.buttonVideoHelp.Text = "Về chế độ làm bài";
                this.TopMost = false;
                try
                {
                    Process.Start(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\hdh.mp4"));
                }
                catch (Exception ex)
                {
                    int num = (int)MessageBox.Show("đổi chương trinh mặt định xem Video khác" + ex.Message);
                }
            }
            else
            {
                this.buttonVideoHelp.Text = "Hướng dẫn sử dụng";
                this.TopMost = true;
            }
        }

        private void buttonTopMost_Click(object sender, EventArgs e)
        {
            if (this.TopMost)
            {
                this.TopMost = false;
                this.button2.Text = "Tự Làm";
                Process.Start("https://meet.google.com/uay-ywnp-xzs");
            }
            else
            {
                this.TopMost = true;
                this.button2.Text = "Gọi Thầy";
            }
        }

        private void buttonHelpVideo_Click(object sender, EventArgs e)
        {
            string str = Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\1.mp4");
            string fileHelpVideoPath = this.paramater.Source_file_help_video_path;
            try
            {
                Home.DecryptFile(fileHelpVideoPath, str);
                Process.Start(str);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Tắt Video help trước khi mở help mới");
            }
        }

        private void buttonHelp_Click_1(object sender, EventArgs e)
        {
            if (this.TopMost)
            {
                this.buttonHelp.Text = "Về chế độ làm bài";
                this.TopMost = false;
            }
            else
            {
                this.buttonHelp.Text = "Về chế độ thường";
                this.TopMost = true;
            }
        }
    }
}