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

namespace MOS_EXCEL_LEARN
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
        private Microsoft.Office.Interop.Excel.Application a;
        private Microsoft.Office.Interop.Excel.Workbook d;
        private Microsoft.Office.Interop.Excel.Workbook d1;
        private int tong_so_cau = ClsListQuestion.GetQuestionNumber();
        private int cau_User = 1;
        private bool chotat = false;
        public object readOnly = (object)false;
        public object isVisible = (object)true;
        public object missing = (object)Missing.Value;
        private enviroment paramater;
        private Help h;
        private int[] Diem;
        private ClsQuestion questionObj;
        public Form1()
        {
            //this.h = new Help();   TODO: uncomment after finish Help form
            this.InitializeComponent();
            this.FormClosed += new FormClosedEventHandler(this.Form1_FormClosed);
            try
            {
                this.a = (Microsoft.Office.Interop.Excel.Application)Activator.CreateInstance(System.Type.GetTypeFromCLSID(new Guid("00024500-0000-0000-C000-000000000046")));
                this.load_cau_hoi(this.cau_User);
                // ISSUE: method pointer
                // ISSUE: object of a compiler-generated type is created
                //((Microsoft.Office.Interop.Excel.AppEvents_Event)this.a)
                //.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(this.a_WorkbookBeforeClose);

                // Initially show the Excel application
                this.a.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
                System.Threading.Thread.Sleep(200); // cho Office kịp cập nhật
                this.a.Top = 0.0;
                this.a.Left = 0.0;
                this.a.Width = (double)this.screen_width;
                this.a.Height = (double)(this.screen_height * 3 / 5);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
        }

        private void a_WorkbookBeforeClose(Microsoft.Office.Interop.Excel.Workbook Wb, ref bool Cancel)
        {
            if (this.chotat)
                return;
            Cancel = true;
        }

        private void Form1_New_Load(object sender, EventArgs e)
        {

            this.loadDanhDau();
            try
            {
                //this.loadTongSoCau();
                //this.tong_so_cau = School.Tong();
                this.Diem = new int[this.tong_so_cau];
                //int y = this.comboBoxCauNext.Height <= this.buttonExit.Height ? (this.buttonExit.Height - this.comboBoxCauNext.Height) / 2 : (this.comboBoxCauNext.Height - this.buttonExit.Height) / 2;
                int y = 5;
                int leftPadRight = 5;
                int rightPadLeft = 5;
                this.screen_height = this.Height;
                this.screen_width = this.Width;
                this.WindowState = FormWindowState.Normal;
                this.Width = this.screen_width;
                this.Height = this.screen_height / 5;
                this.Location = new Point(0, this.screen_height * 4 / 5);
                this.buttonExit.Location = new Point(0, y);
                this.buttonVideoHelp.Location = new Point(this.buttonExit.Width + leftPadRight, y);
                this.buttonRefresh.Location = new Point(this.buttonExit.Width + this.buttonVideoHelp.Width + leftPadRight + 5, y);

                this.buttonHelp.Location = new Point(this.buttonExit.Width + this.buttonVideoHelp.Width + this.buttonRefresh.Width + leftPadRight + 10, y);
                this.buttonHelpVideo.Location = new Point(this.buttonExit.Width + this.buttonVideoHelp.Width + this.buttonRefresh.Width + this.buttonHelp.Width + leftPadRight + 15, y);
                this.buttonCheck.Location = new Point(this.buttonExit.Width + this.buttonVideoHelp.Width + this.buttonRefresh.Width + this.buttonHelp.Width + this.buttonHelpVideo.Width + leftPadRight + 20, y);
                this.labelKQ.Location = new Point(this.buttonExit.Width + this.buttonVideoHelp.Width + this.buttonRefresh.Width + this.buttonHelp.Width + this.buttonHelpVideo.Width + this.buttonCheck.Width + leftPadRight + 25, y);


                this.buttonEV.Location = new Point(this.Width - this.buttonEV.Width, y);
                this.comboBoxCauNext.Location = new Point(this.Width - this.buttonEV.Width - this.comboBoxCauNext.Width - rightPadLeft, y);
                this.buttonNext.Location = new Point(this.Width - this.buttonEV.Width - this.comboBoxCauNext.Width - this.buttonNext.Width - rightPadLeft - 5, y);
                this.buttonReset.Location = new Point(this.buttonNext.Location.X - this.buttonReset.Width - rightPadLeft, y);
                this.labelCauHienTai.Location = new Point(this.buttonReset.Location.X - this.labelCauHienTai.Width - rightPadLeft, y);
                this.buttonSummary.Location = new Point(this.labelCauHienTai.Location.X - this.buttonSummary.Width - rightPadLeft, y);
                this.comboBoxCauDaDanhDau.Location = new Point(this.buttonSummary.Location.X - this.comboBoxCauDaDanhDau.Width - rightPadLeft, y);
                this.buttonZoom.Location = new Point(this.comboBoxCauDaDanhDau.Location.X - this.buttonZoom.Width - rightPadLeft, y);

                //this.button2.Location = new Point(this.buttonZoom.Location.X - this.button2.Width, y); //Gọi cô

                this.richTextBox1.Location = new Point(0, this.comboBoxCauNext.Height + 10);
                this.richTextBox1.Width = this.Width;
                this.richTextBox1.Height = this.Height - this.comboBoxCauNext.Height;
                this.panel1.Location = new Point(0, this.comboBoxCauNext.Height + 10);
                this.panel1.Width = this.Width;
                this.panel1.Height = this.Height - this.comboBoxCauNext.Height;
                this.richTextQuestion.Location = new Point(0, 0);
                this.richTextQuestion.Width = this.Width - this.buttonEV.Width / 3;
                this.comboBoxCauNext.Text = (this.cau_User + 1).ToString();
                this.labelCauHienTai.Text = "Câu " + this.cau_User.ToString();
                for (int index = 0; index < this.tong_so_cau; ++index)
                    this.comboBoxCauNext.Items.Add((object)(index + 1));
                this.Copy(Path.Combine(System.Windows.Forms.Application.StartupPath, "data\\Doc"), Environment.GetFolderPath(Environment.SpecialFolder.Personal));

                this.a.Visible = true;
                this.a.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
                this.a.Top = 0.0;
                this.a.Left = 0.0;
                this.a.Width = (double)this.screen_width;
                this.a.Height = (double)(this.screen_height * 3 / 5);
                this.TopMost = true;
                this.currenTime = DateTime.Now;
                this.setDiem();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
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

        private void setDiem()
        {
            for (int index = 0; index < this.Diem.Length; ++index)
                this.Diem[index] = 0;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        private void load_cau_hoi(int cau_hoi_so)
        {
            questionObj = ClsListQuestion.GetQuestion(cau_hoi_so);
            if (!questionObj.Status)
            {
                throw new ArgumentException("Câu hỏi chưa sẵn sàng để học");
            }
            this.paramater = this.GetparmaterNew(questionObj.CorrectIndex);
            //this.paramater = this.Getparmater(cau_hoi_so);
            this.setDefalt(this.paramater);
            this.paramater.Dest_file_Word_Name = Path.Combine(System.Windows.Forms.Application.StartupPath, "Word\\" + this.questionObj.MaskIndex.ToString() + ".docx");
            this.paramater.DeTiengAnh = this.paramater.Source_de_En;
            this.paramater.DeTiengViet = this.paramater.Source_de_Vn;
            Home.DecryptFile(this.paramater.Source_file_word_path, this.paramater.Dest_file_Word_Name.ToString());
            this.richTextQuestion.Text = this.paramater.DeTiengViet;
            if (this.panel1.Width > this.richTextQuestion.Width)
                this.richTextQuestion.Left = (this.panel1.Width - this.richTextQuestion.Width) / 2;
            else
                this.richTextQuestion.Left = 0;
            this.d = this.a.Workbooks.Open(this.paramater.Dest_file_Word_Name, System.Type.Missing, System.Type.Missing, System.Type.Missing, (object)"271565234", System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
        }

        private enviroment GetparmaterNew(int cauUser)
        {
            enviroment enviroment = new enviroment();
            int a = 0;
            enviroment.quesion = cauUser;
            for (int index = 0; index < School.Mn.Length; ++index)
            {
                a = a + School.Mn[index];
                if (cauUser <= a)
                {
                    enviroment.section = index;
                    break;
                }
                enviroment.quesion = cauUser - a;
                enviroment.section = index;
            }

            enviroment.dirPath = Path.Combine(System.Windows.Forms.Application.StartupPath);
            enviroment.Source_de_En = ClsListQuestion.GetEngQuestion(cauUser);
            enviroment.Source_de_Vn = ClsListQuestion.GetVNQuestion(cauUser);
            enviroment.Source_file_word_path = Path.Combine(enviroment.dirPath, "data\\" + questionObj.MaskIndex.ToString());

            return enviroment;
        }

        private void setDefalt(enviroment par)
        {
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

        private void button1_Click(object sender, EventArgs e)
        {
            this.LuuDiem();
            this.LuuDanhDau();
            object SaveChanges = (object)Microsoft.Office.Interop.Excel.XlSaveAction.xlDoNotSaveChanges;
            try
            {
                this.chotat = true;
                while (this.a.Windows.Count >= 1)
                {
                    // ISSUE: reference to a compiler-generated method
                    this.a.Windows[(object)1].Close(SaveChanges, System.Type.Missing, System.Type.Missing);
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
            //this.pictureBox1.Image.Dispose();
            foreach (string file in Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "tam")))
                File.Delete(file);

            Process[] excelProcesses = Process.GetProcessesByName("EXCEL");

            if (excelProcesses.Length > 0)
            {
                foreach (var process in excelProcesses)
                {
                    try
                    {
                        process.Kill(); // buộc dừng tiến trình
                        process.WaitForExit(); // đợi đến khi tiến trình thật sự đóng
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Không thể tắt Excel: " + ex.Message);
                    }
                }
            }

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
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
            this.buttonEV.Text = "Tiếng Việt";
            //this.h.refresh(); TODO: uncomment after finish Help form
            this.a.Visible = true;
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
                int num2 = (int)MessageBox.Show("Chọn file excel nhấn nút ESC trước khi reset");
                num1 = 1;
            }
            if (num1 == 1)
                return;
            this.load_cau_hoi(this.cau_User);
            this.a.Visible = true;
        }

        private void buttonEV_Click(object sender, EventArgs e)
        {
            if (this.buttonEV.Text == "Tiếng Việt")
            {
                this.buttonEV.Text = "Tiếng Anh";
                this.richTextQuestion.Text = this.paramater.DeTiengAnh;
            }
            else
            {
                this.buttonEV.Text = "Tiếng Việt";
                this.richTextQuestion.Text = this.paramater.DeTiengViet;
            }
        }

        private void buttonHelp_Click(object sender, EventArgs e)
        {
            if (this.TopMost)
            {
                //this.buttonHelp.Text = "Về chế độ làm bài";
                this.buttonHelp.Text = "Ẩn Taskbar";
                this.TopMost = false;
            }
            else
            {
                //this.buttonHelp.Text = "Về chế độ thường";
                this.buttonHelp.Text = "Hiện Taskbar";
                this.TopMost = true;
            }
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                this.a.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
                this.a.Top = 0.0;
                this.a.Left = 0.0;
                this.a.Width = (double)this.screen_width;
                this.a.Height = (double)(this.screen_height * 3 / 5);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Bạn đã tắt MS Excel! Tắt chương trình và làm lại");
            }
        }

        private void buttonCheck_Click(object sender, EventArgs e)
        {
            this.labelKQ.Text = this.CheckResult(this.paramater);
            if (this.labelKQ.Text == "True")
                this.Diem[this.cau_User - 1] = 1;
            this.check = false;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
        }

        private void pictureBox1_MouseHover(object sender, EventArgs e)
        {
        }

        private void buttonZoom_Click(object sender, EventArgs e)
        {
            this.comboBoxCauDaDanhDau.Items.Add((object)this.cau_User);
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
            //this.h.refresh();    TODO: uncomment after finish Help form
            this.a.Visible = true;
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
            if (this.buttonVideoHelp.Text == "HDSD")
            {
                //this.buttonVideoHelp.Text = "Về chế độ làm bài";
                this.buttonHelp.Text = "Ẩn Taskbar";
                this.TopMost = false;
                try
                {
                    //Process.Start(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\hdh.mp4"));
                    Process.Start("https://go.mos360.vn/mosexcelhdsd");
                }
                catch (Exception ex)
                {
                    int num = (int)MessageBox.Show("Vui lòng kết nối mạng để xem Hướng dẫn sử dụng phần mềm " + ex.Message);
                }
            }
            else
            {
                this.buttonVideoHelp.Text = "HDSD";
                this.TopMost = true;
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (this.TopMost)
            {
                this.TopMost = false;
                this.button2.Text = "Tự Làm";
                Process.Start("https://google.com");
            }
            else
            {
                this.TopMost = true;
                this.button2.Text = "Gọi Thầy";
            }
        }

        private void buttonHelpVideo_Click(object sender, EventArgs e)
        {
            try
            {
                if (!questionObj.Status)
                {
                    throw new ArgumentException("Câu hỏi chưa sẵn sàng để học");
                }
                Process.Start(questionObj.Url);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
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
            string str4 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "MOS360");
            if (!Directory.Exists(str4))
                Directory.CreateDirectory(str4);
            string path = Path.Combine(str4, "MOS360 TIẾN ĐỘ LUYỆN TẬP - EXCEL.html");
            if (File.Exists(path))
            {
                TextWriter textWriter = (TextWriter)new StreamWriter(path, true);
                textWriter.WriteLine(str3);
                textWriter.Close();
            }
            else
            {
                TextWriter textWriter = (TextWriter)new StreamWriter(path, true);
                string str5 = "<h2 style=\"text-align: center;\"><span style=\"color: #ff0000;\"><strong>MOS360 - X&oacute;a tan nỗi lo CHUẨN ĐẦU RA cho sinh vi&ecirc;n.</strong></span></h2>\r\n<p style=\"padding-left: 40px;\">Kh&ocirc;ng cần ho&agrave;n hảo - chỉ cần bắt đầu. Gửi b&agrave;i cho MOS360 để giữ vững phong độ nh&eacute; 🚀</p>\r\n<p style=\"padding-left: 40px;\">Zalo: <a href=\"https://zalo.me/0912888360\" target=\"_blank\" rel=\"noopener\">0912.888.360</a></p>\r\n<p style=\"padding-left: 40px;\">Fanpage: <a href=\"https://www.facebook.com/mos360.vn\" target=\"_blank\">https://www.facebook.com/mos360.vn</a></p>\r\n<p style=\"padding-left: 40px;\">Website: <a href=\"https://mos360.vn\" target=\"_blank\">https://mos360.vn</a></p>\r\n<p style=\"padding-left: 40px;\">Ch&uacute;c bạn học MOS thật vui, l&agrave;m b&agrave;i thật 'phi&ecirc;u' v&agrave; điểm cao v&egrave;o v&egrave;o nh&eacute; 🚀🎯</p>\r\n<p style=\"text-align: center;\"><strong>KẾT QUẢ LUYỆN TẬP EXCEL</strong></p>\r\n<p style=\"padding-left: 40px;\">&nbsp;</p>";
                textWriter.WriteLine(str5);
                textWriter.WriteLine(str3);
                textWriter.Close();
            }
        }

        private void turnOff()
        {
            object SaveChanges = (object)Microsoft.Office.Interop.Excel.XlSaveAction.xlDoNotSaveChanges;
            this.chotat = true;
            while (this.a.Windows.Count >= 1)
            {
                // ISSUE: reference to a compiler-generated method
                this.a.Windows[(object)1].Close(SaveChanges, System.Type.Missing, System.Type.Missing);
            }
            this.chotat = false;
            //this.pictureBox1.Image.Dispose();
            foreach (string file in Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "tam")))
                File.Delete(file);
            foreach (string file in Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "Word")))
                File.Delete(file);
        }

        private string CheckResult(enviroment par)
        {
            string str = "press ESC and close all dialogBox to contuine";
            try
            {
                str = this.Check(par);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Nhấn nút ESC và đóng các hộp thoại để tiếp tục");
            }
            return str;
        }

        private string Check(enviroment par)
        {
            //MessageBox.Show(par.section.ToString() + " - " + par.quesion.ToString());

            string str = "";
            try
            {
                switch (par.section)
                {
                    case 0:
                        str = Sec0.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 1:
                        str = Sec1.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 2:
                        str = Sec2.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 3:
                        str = Sec3.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 4:
                        str = Sec4.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 5:
                        str = Sec5.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 6:
                        str = Sec6.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 7:
                        str = Sec7.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 8:
                        str = Sec8.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 9:
                        str = Sec9.CheckCau(par.quesion, this.a, this.d);
                        break;
                    case 10:
                        str = CheckWord.CheckCau(par.quesion, (Microsoft.Office.Interop.Excel._Application)this.a, (Microsoft.Office.Interop.Excel._Workbook)this.d).ToString();
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

        private void LuuDanhDau()
        {
            TextWriter textWriter = (TextWriter)new StreamWriter(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip//luudanhdau"));
            for (int index = 0; index < this.comboBoxCauDaDanhDau.Items.Count; ++index)
                textWriter.WriteLine(this.comboBoxCauDaDanhDau.Items[index].ToString());
            textWriter.Close();
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
            this.buttonEV.Text = "Tiếng Việt";
            //this.h.refresh(); TODO: uncomment after finish Help form
            this.a.Visible = true;
        }
    }
}
