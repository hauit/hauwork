using Microsoft.Office.Interop.Word;
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
using WindowsFormsApplication1.Base;
using Point = System.Drawing.Point;

namespace WindowsFormsApplication1
{
    public partial class Form1_Dup : Form
    {
        private int zoom = 0;
        private bool check = true;
        private int so_cau_dung = 0;
        private int so_cau_sai = 0;
        private DateTime currenTime;
        private int screen_height;
        private int screen_width;
        private Microsoft.Office.Interop.Word.Application a;
        private Microsoft.Office.Interop.Word._Document d;
        private _Document d1;
        private int tong_so_cau = School.Tong();
        private int cau_User = 1;
        private bool chotat = false;
        public object readOnly = (object)false;
        public object isVisible = (object)true;
        public object missing = (object)Missing.Value;
        private enviroment paramater;
        private Help_cu h;
        private int[] Diem;
        public Form1_Dup()
        {
            this.h = new Help_cu();
            InitializeComponent();
            this.FormClosed += new FormClosedEventHandler(this.Form1_Dup_FormClosed);
            this.a = (Microsoft.Office.Interop.Word.Application)Activator.CreateInstance(System.Type.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
            this.load_cau_hoi(this.cau_User);
            // ISSUE: method pointer
            // ISSUE: object of a compiler-generated type is created
            ////TODO: need to check this decode code
            //new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "DocumentBeforeClose").AddEventHandler((object)this.a, (Delegate)new ApplicationEvents4_DocumentBeforeCloseEventHandler((object)this, (UIntPtr)__methodptr(a_DocumentBeforeClose)));
        }

        private void Form1_Dup_Load(object sender, EventArgs e)
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
                this.buttonHelp.Location = new Point(this.buttonExit.Width + this.buttonRefresh.Width, y);
                this.buttonHelpVideo.Location = new Point(this.buttonExit.Width + this.buttonRefresh.Width + this.buttonHelp.Width, y);
                this.buttonCheck.Location = new Point(this.buttonExit.Width + this.buttonRefresh.Width + this.buttonHelp.Width + this.buttonHelpVideo.Width, y);
                this.labelKQ.Location = new Point(this.buttonExit.Width + this.buttonRefresh.Width + this.buttonHelp.Width + this.buttonHelpVideo.Width + this.buttonCheck.Width, y);
                this.buttonEV.Location = new Point(this.Width - this.buttonEV.Width, y);
                this.comboBoxCauNext.Location = new Point(this.Width - this.buttonEV.Width - this.comboBoxCauNext.Width, 0);
                this.buttonNext.Location = new Point(this.Width - this.buttonEV.Width - this.comboBoxCauNext.Width - this.buttonNext.Width, y);
                this.buttonReset.Location = new Point(this.buttonNext.Location.X - this.buttonReset.Width, y);
                this.labelCauHienTai.Location = new Point(this.buttonReset.Location.X - this.labelCauHienTai.Width, y);
                this.buttonSummary.Location = new Point(this.labelCauHienTai.Location.X - this.buttonSummary.Width, y);
                this.comboBoxCauDaDanhDau.Location = new Point(this.buttonSummary.Location.X - this.comboBoxCauDaDanhDau.Width, 0);
                this.buttonZoom.Location = new Point(this.comboBoxCauDaDanhDau.Location.X - this.buttonZoom.Width, y);
                this.button2.Location = new Point(this.buttonZoom.Location.X - this.button2.Width, y);
                this.buttonVideoHelp.Location = new Point(this.button2.Location.X - this.buttonVideoHelp.Width, y);
                this.richTextBox1.Location = new Point(0, this.comboBoxCauNext.Height);
                this.richTextBox1.Width = this.Width;
                this.richTextBox1.Height = this.Height - this.comboBoxCauNext.Height;
                this.panel1.Location = new Point(0, this.comboBoxCauNext.Height);
                this.panel1.Width = this.Width;
                this.panel1.Height = this.Height - this.comboBoxCauNext.Height;
                //this.pictureBox1.Location = new Point(0, 0);
                //this.pictureBox1.Width = this.Width - this.buttonEV.Width / 3;
                this.richTextQuestion.Location = new Point(0, 0);
                this.richTextQuestion.Width = this.Width - this.buttonEV.Width / 3;
                this.comboBoxCauNext.Text = (this.cau_User + 1).ToString();
                this.labelCauHienTai.Text = "Câu: " + this.cau_User.ToString();
                for (int index = 0; index < this.tong_so_cau; ++index)
                    this.comboBoxCauNext.Items.Add((object)(index + 1));
                this.Copy(Path.Combine(System.Windows.Forms.Application.StartupPath, "data\\Doc"), Environment.GetFolderPath(Environment.SpecialFolder.Personal));
                this.a.Visible = true;
                this.a.WindowState = WdWindowState.wdWindowStateNormal;
                this.a.Top = 0;
                this.a.Left = -10;
                // ISSUE: reference to a compiler-generated method
                this.a.Resize(this.screen_width, this.screen_height * 3 / 5);
                this.TopMost = true;
                this.currenTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
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
                    str1 = str1 + (object)(index + 1) + ",";
            }
            string str2 = num.ToString() + "Câu," + str1;
            string str3 = "<p>" + DateTime.Now.ToShortDateString() + "," + str2 + "</p>";
            if (num <= 1)
                return;
            string str4 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Diem");
            if (!Directory.Exists(str4))
                Directory.CreateDirectory(str4);
            string path = Path.Combine(str4, "hoc_Word.html");
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

        private void a_DocumentBeforeClose(Document Doc, ref bool Cancel)
        {
            if (this.chotat)
                return;
            Cancel = true;
        }

        private void Form1_Dup_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            this.LuuDiem();
            this.LuuDanhDau();
            object SaveChanges1 = (object)WdSaveOptions.wdDoNotSaveChanges;
            try
            {
                this.chotat = true;
                while (this.a.Windows.Count >= 1)
                {
                    object Index = (object)1;
                    // ISSUE: reference to a compiler-generated method
                    this.a.Windows[ref Index].Close(ref SaveChanges1, ref this.missing);
                }
                if (this.a != null)
                {
                    object SaveChanges2 = (object)Missing.Value;
                    object OriginalFormat = (object)Missing.Value;
                    object RouteDocument = (object)Missing.Value;
                    // ISSUE: reference to a compiler-generated method
                    this.a.Quit(ref SaveChanges2, ref OriginalFormat, ref RouteDocument);
                }
                this.chotat = false;
            }
            catch (Exception ex)
            {
            }
            foreach (string file in Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "Word")))
                File.Delete(file);
            //this.pictureBox1.Image.Dispose();
            foreach (string file in Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "tam")))
                File.Delete(file);
            this.Close();
        }

        private void turnOff()
        {
            object SaveChanges = (object)WdSaveOptions.wdDoNotSaveChanges;
            this.chotat = true;
            while (this.a.Windows.Count >= 1)
            {
                object Index = (object)1;
                // ISSUE: reference to a compiler-generated method
                this.a.Windows[ref Index].Close(ref SaveChanges, ref this.missing);
            }
            this.chotat = false;
            //this.pictureBox1.Image.Dispose();
            foreach (string file in Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "tam")))
                File.Delete(file);
            foreach (string file in Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "Word")))
                File.Delete(file);
        }

        private void loadTongSoCau()
        {
            for (int index = 0; index < School.Mn.Length; ++index)
            {
                string[] files = Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "Data\\Sec_" + (object)index + "\\hinh\\E"));
                School.Mn[index] = files.Length;
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
                    enviroment.Source_de_En = ClsListQuestion.GetEngQuestion(cauUser);
                    enviroment.Source_de_Vn = ClsListQuestion.GetVNQuestion(cauUser);
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
            this.paramater.Dest_file_Word_Name = (object)Path.Combine(System.Windows.Forms.Application.StartupPath, "Word\\" + this.paramater.section.ToString() + "_" + this.paramater.quesion.ToString());
            ////TODO: sua doi thanh text
            this.paramater.DeTiengAnh = this.paramater.Source_de_En;
            this.paramater.DeTiengViet = this.paramater.Source_de_Vn;
            this.paramater.Dest_file_help_Name = Path.Combine(System.Windows.Forms.Application.StartupPath, "tam\\help");
            this.paramater.Dest_file_help_video_Name = Path.Combine(System.Windows.Forms.Application.StartupPath, "tam\\Vhelp");
            Home.DecryptFile(this.paramater.Source_file_word_path, this.paramater.Dest_file_Word_Name.ToString());
            ////TODO: họ tạo file tạm cho câu hỏi hiện tại trong thư mục tam\\ rồi copy đề + video vào đấy
            ///         hiện tại đã bỏ việc tạo file của 2 câu hỏi TA và việt(mình tạo file text câu hỏi rồi đọc luôn vào biến)
            ///         Với file video thì tạm thời đề lại(file của họ là mã hóa. cũng cần thay đổi - bỏ mã hóa)
            //Home.DecryptFile(this.paramater.Source_de_En, this.paramater.DeTiengAnh);
            //Home.DecryptFile(this.paramater.Source_de_Vn, this.paramater.DeTiengViet);
            if (File.Exists(this.paramater.Source_file_help_video_path))
                Home.DecryptFile(this.paramater.Source_file_help_video_path, this.paramater.Dest_file_help_video_Name);
            //this.pictureBox1.Image = Image.FromFile(this.paramater.DeTiengAnh);
            //if (this.panel1.Width > this.pictureBox1.Width)
            //    this.pictureBox1.Left = (this.panel1.Width - this.pictureBox1.Width) / 2;
            //else
            //    this.pictureBox1.Left = 0;
            this.richTextQuestion.Text = this.paramater.DeTiengAnh;
            if (this.panel1.Width > this.richTextQuestion.Width)
                this.richTextQuestion.Left = (this.panel1.Width - this.richTextQuestion.Width) / 2;
            else
                this.richTextQuestion.Left = 0;
            object PasswordDocument = (object)"271565234";
            // ISSUE: reference to a compiler-generated method
            this.d = (_Document)this.a.Documents.Open(ref this.paramater.Dest_file_Word_Name, ref this.missing, ref this.readOnly, ref this.missing, ref PasswordDocument, ref this.missing, ref this.missing, ref this.missing, ref this.missing, ref this.missing, ref this.missing, ref this.isVisible, ref this.missing, ref this.missing, ref this.missing, ref this.missing);
            this.a.Top = 0;
            this.a.Left = -10;
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
                        str = PageSetup.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 1:
                        ////TODO:
                        str = Design.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 2:
                        str = Header_Footer.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 3:
                        ////TODO:
                        //str = Style.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 4:
                        str = Paragaph.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 5:
                        str = Illustrator.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 6:
                        str = Table.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 7:
                        str = DocText.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 8:
                        str = Links.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 9:
                        str = Reference.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 10:
                        str = Final_Steps.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 11:
                        str = CheckWord.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 12:
                        ////TODO:
                        //str = OnTap.CheckCau(par.quesion, this.a, this.d);
                        break;
                    default:
                        str = "Chua cham";
                        break;
                }
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Đóng tất cả các hộp thoại đang mở trước");
            }
            return str;
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                this.a.WindowState = WdWindowState.wdWindowStateNormal;
                this.a.Top = 0;
                this.a.Left = -10;
                // ISSUE: reference to a compiler-generated method
                this.a.Resize(this.screen_width, this.screen_height * 3 / 5);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Bạn đã tắt MS Word! Tắt chương trình và làm lại");
            }
        }

        private void buttonReset_Click(object sender, EventArgs e)
        {
            this.labelKQ.Text = "";
            int num1 = 0;
            try
            {
                this.turnOff();
            }
            catch (Exception ex)
            {
                int num2 = (int)MessageBox.Show("Đóng các hộp thoại trước khi reset");
                num1 = 1;
            }
            if (num1 == 1)
                return;
            this.load_cau_hoi(this.cau_User);
        }

        private void buttonEV_Click(object sender, EventArgs e)
        {
            if (this.buttonEV.Text == "V")
            {
                this.buttonEV.Text = "E";
                this.buttonEV.BackgroundImage = WindowsFormsApplication1.Properties.Resources.en;
                this.richTextQuestion.Text = this.paramater.DeTiengAnh;
                //this.pictureBox1.Image.Dispose();
                //this.pictureBox1.Image = Image.FromFile(this.paramater.DeTiengViet);
            }
            else
            {
                this.buttonEV.Text = "V";
                this.buttonEV.BackgroundImage = WindowsFormsApplication1.Properties.Resources.vi;
                this.richTextQuestion.Text = this.paramater.DeTiengViet;
                //this.pictureBox1.Image.Dispose();
                //this.pictureBox1.Image = Image.FromFile(this.paramater.DeTiengAnh);
            }
        }

        private void buttonHelp_Click(object sender, EventArgs e) => this.h.Show();

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

        private void buttonZoom_Click(object sender, EventArgs e) => this.comboBoxCauDaDanhDau.Items.Add((object)this.cau_User);

        private void buttonVideoHelp_Click(object sender, EventArgs e)
        {
            if (this.buttonVideoHelp.Text == "Hướng dẫn phần mềm")
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
                this.buttonVideoHelp.Text = "Hướng dẫn phần mềm";
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
                this.button2.Text = "Gọi Cô";
            }
        }

        private void buttonHelpVideo_Click(object sender, EventArgs e)
        {
            string str = Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\1.avi");
            string fileHelpVideoPath = this.paramater.Source_file_help_video_path;
            try
            {
                //Home.DecryptFile(fileHelpVideoPath, str);
                str = "https://www.youtube.com/watch?v=fT_3Bsb9hbA";
                Process.Start(str);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Tắt Video hướng dẫn làm bài trước khi mở hướng dẫn mới");
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

        private void buttonNext_Click(object sender, EventArgs e)
        {
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
            this.buttonEV.Text = "V";
        }
    }
}
