using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Point = System.Drawing.Point;

namespace MOS_PPT_LEARN
{
    public partial class Form1 : Form
    {
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

        public Form1()
        {
            this.h = new Help();
            InitializeComponent();
            this.a = (Microsoft.Office.Interop.PowerPoint.Application)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("91493441-5A91-11CF-8700-00AA0060263B")));
            this.load_cau_hoi(this.cau_User);
            this.a.PresentationBeforeClose += this.a_PresentationBeforeClose;
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

        private void buttonZoom_Click(object sender, EventArgs e)
        {
            this.comboBoxCauDaDanhDau.Items.Add((object)this.cau_User);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
        }

        private void pictureBox1_MouseHover(object sender, EventArgs e) => this.panel1.Focus();

        private void buttonCheck_Click(object sender, EventArgs e)
        {
            this.labelKQ.Text = this.CheckResult(this.paramater);
            if (this.labelKQ.Text == "True")
                this.Diem[this.cau_User - 1] = 1;
            this.check = false;
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

        private void load_cau_hoi(int cau_hoi_so)
        {
            this.paramater = this.Getparmater(cau_hoi_so);
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

        private void Copy(string sourceDir, string targetDir)
        {
            foreach (string file in Directory.GetFiles(sourceDir))
                File.Copy(file, Path.Combine(targetDir, Path.GetFileName(file)), true);
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
    }
}

    
