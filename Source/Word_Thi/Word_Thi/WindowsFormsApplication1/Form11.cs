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
using WindowsFormsApplication1.Properties;
using File = System.IO.File;
using Point = System.Drawing.Point;

namespace WindowsFormsApplication1
{
    public partial class Form11 : Form
    {
        private string pathWork;
        private string pathRun;
        private string pathFileOfficeMaHoa;
        private string pathFileOffice;
        private string pathReset;
        private int works;
        private int workIndex;
        private Microsoft.Office.Interop.Word.Application a;
        private string[] ImageFile;
        private int screen_height;
        private int screen_width;
        private Document d;
        private DateTime timeStrart;
        private int Cau_So;
        private int next = 0;
        private int soLanReSet = 0;
        private double Diem = 0.0;
        private bool[] CacCauDaCheck;
        private int So_Cau_Dung = 0;
        private int So_Cau_Sai = 0;
        private int Tong_So_Cau;
        private List<int> lsViTri;
        private bool chotat = false;
        private int currentest = 0;
        public Form11()
        {
            InitializeComponent();
            this.textBox1.MouseClick += new MouseEventHandler(this.textBox1_MouseClick);
            this.getEnviroment();
        }

        private void a_DocumentBeforeClose(Document Doc, ref bool Cancel)
        {
            if (this.chotat)
                return;
            Cancel = true;
        }

        private void textBox1_MouseClick(object sender, MouseEventArgs e)
        {
            this.textBox1.Text = this.panel2.VerticalScroll.Value.ToString();
            if (this.lsViTri.Count > this.Cau_So)
                this.lsViTri[this.Cau_So] = this.panel2.VerticalScroll.Value;
            else
                this.lsViTri.Add(this.panel2.VerticalScroll.Value);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Bạn có muốn reset không? (reset sẽ làm lại từ đầu)", "Cảnh Báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) != DialogResult.Yes)
                    return;
                this.chotat = true;
                while (this.a.Documents.Count >= 1)
                {
                    object SaveChanges = (object)false;
                    object missing1 = System.Type.Missing;
                    object missing2 = System.Type.Missing;
                    object Index = (object)1;
                    // ISSUE: reference to a compiler-generated method
                    this.a.Documents[ref Index].Close(ref SaveChanges, ref missing1, ref missing2);
                }
                this.chotat = false;
                Home.DecryptFile(this.pathFileOfficeMaHoa, this.pathFileOffice);
                object pathFileOffice = (object)this.pathFileOffice;
                object missing3 = System.Type.Missing;
                object missing4 = System.Type.Missing;
                object missing5 = System.Type.Missing;
                object PasswordDocument = (object)"271565234";
                object missing6 = System.Type.Missing;
                object missing7 = System.Type.Missing;
                object missing8 = System.Type.Missing;
                object missing9 = System.Type.Missing;
                object missing10 = System.Type.Missing;
                object missing11 = System.Type.Missing;
                object missing12 = System.Type.Missing;
                object missing13 = System.Type.Missing;
                object missing14 = System.Type.Missing;
                object missing15 = System.Type.Missing;
                object XMLTransform = (object)Missing.Value;
                // ISSUE: reference to a compiler-generated method
                this.d = this.a.Documents.Open(ref pathFileOffice, ref missing3, ref missing4, ref missing5, ref PasswordDocument, ref missing6, ref missing7, ref missing8, ref missing9, ref missing10, ref missing11, ref missing12, ref missing13, ref missing14, ref missing15, ref XMLTransform);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Đống các cửa sổ thông báo của Word trước khi Reset");
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Cau_So = this.checkedListBox1.SelectedIndex;
            if (this.lsViTri.Count < this.checkedListBox1.Items.Count)
            {
                this.Cau_So = this.checkedListBox1.SelectedIndex;
                this.panel2.VerticalScroll.Value = this.Cau_So * this.panel2.VerticalScroll.Maximum / this.Tong_So_Cau + this.panel2.VerticalScroll.Minimum;
                this.richTextBox1.Text = "KQ";
            }
            else
            {
                this.Cau_So = this.checkedListBox1.SelectedIndex;
                this.panel2.VerticalScroll.Value = this.lsViTri[this.Cau_So];
                this.richTextBox1.Text = "KQ";
            }
        }

        private void getEnviroment()
        {
            int length = Directory.GetDirectories(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\MaHoa")).Length;
            Random r = new Random();
            if (Program.TypeOfTest == 0)
            {
                this.button1.Visible = true;
                this.textBox2.Visible = true;
                Program.Tong = length;
                Program.Lessons = new int[length];
                for (int index = 0; index < length; ++index)
                    Program.Lessons[index] = index;
            }
            else
            {
                this.button1.Visible = false;
                this.textBox2.Visible = false;
                Program.Tong = 7;
                Program.Lessons = new int[length];
                for (int i = 0; i < length; ++i)
                    Program.Lessons[i] = this.chonBai(0, length, i, r);
            }
        }

        private int chonBai(int min, int max, int i, Random r)
        {
            int kq;
            do
            {
                kq = r.Next(min, max);
            }
            while (this.tontai(kq, i));
            return kq;
        }

        private bool tontai(int kq, int c)
        {
            for (int index = 0; index < c; ++index)
            {
                if (Program.Lessons[index] == kq)
                    return true;
            }
            return false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = true;
            int num = this.Height / 5;
            int width = this.Width;
            this.screen_height = this.Height;
            this.screen_width = this.Width;
            this.WindowState = FormWindowState.Normal;
            this.Height = num;
            this.Width = width;
            this.Location = new Point(0, num * 4);
            this.checkedListBox1.Location = new Point(0, 0);
            this.checkedListBox1.Height = this.Height;
            this.tabControl1.Location = new Point(this.checkedListBox1.Width, 0);
            this.tabControl1.Height = this.Height;
            this.tabControl1.Width = this.Width - this.checkedListBox1.Width - this.buttonReset.Width;
            this.panel1.Location = new Point(0, 0);
            this.panel1.Height = this.tabPage2.Height;
            this.panel1.Width = this.tabPage2.Width;
            this.panel2.Location = new Point(0, 0);
            this.panel2.Height = this.tabPage2.Height;
            this.panel2.Width = this.tabPage2.Width;
            this.buttonExit.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width, 0);
            this.buttonReset.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width, 0);
            this.buttonSubmit.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width, 0);
            this.label1.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width, 0);
            this.button2.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.button2.Width, 0);
            this.comboBox1.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.button2.Width - this.comboBox1.Width, 0);
            this.buttonSave.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.comboBox1.Width - this.button2.Width, 0);
            this.textBox1.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.comboBox1.Width - this.textBox1.Width, 0);
            this.button1.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.textBox1.Width - this.comboBox1.Width - this.button1.Width, 0);
            this.textBox2.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.textBox1.Width - this.button1.Width - this.comboBox1.Width - this.textBox2.Width, 0);
            this.buttonRefresh.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.buttonRefresh.Width - this.comboBox1.Width - this.button2.Width, 0);
            this.buttonShowHide.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.textBox1.Width - this.button1.Width - this.textBox2.Width - this.comboBox1.Width - this.buttonShowHide.Width, 0);
            this.buttonHelp.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.textBox1.Width - this.button1.Width - this.textBox2.Width - this.comboBox1.Width - this.buttonShowHide.Width - this.buttonHelp.Width, 0);
            this.buttonxhdh.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.textBox1.Width - this.button1.Width - this.textBox2.Width - this.comboBox1.Width - this.buttonShowHide.Width - this.buttonHelp.Width - this.buttonxhdh.Width, 0);
            this.panel4.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width, this.Height - this.buttonCheck.Height - this.panel4.Height);
            this.buttonCheck.Location = new Point(this.checkedListBox1.Width + this.tabControl1.Width, this.Height - this.buttonCheck.Height);
            this.loadcaucanhoi();
            this.a = (Microsoft.Office.Interop.Word.Application)Activator.CreateInstance(System.Type.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
            this.a.Visible = true;
            this.a.WindowState = WdWindowState.wdWindowStateNormal;
            this.a.Top = 0;
            this.a.Left = -10;
            this.a.Width = width;
            this.a.Height = num * 3;
            // ISSUE: method pointer
            // ISSUE: object of a compiler-generated type is created
            ((ApplicationEvents4_Event)this.a).DocumentBeforeClose += this.a_DocumentBeforeClose;
            this.SetUp(Program.Lessons[this.currentest]);
            this.timeStrart = DateTime.Now;
            this.timeStrart = this.timeStrart.AddMinutes(50.0);
            this.timer1.Start();
        }

        private void loadcaucanhoi()
        {
            this.comboBox1.Items.Clear();
            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "zip//luu");
            if (!File.Exists(path))
                return;
            TextReader textReader = (TextReader)new StreamReader(path);
            string str;
            do
            {
                str = textReader.ReadLine();
                if (str != null)
                    this.comboBox1.Items.Add((object)str);
            }
            while (str != null);
            textReader.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string str = this.CheckCauLon(this.Cau_So);
            if (!this.CacCauDaCheck[this.Cau_So])
            {
                if (str == "True")
                    ++this.So_Cau_Dung;
                else
                    ++this.So_Cau_Sai;
                this.CacCauDaCheck[this.Cau_So] = true;
            }
            this.richTextBox1.Text = str;
        }

        private void button3_Click(object sender, EventArgs e) => this.submit();

        private void submit()
        {
            if (MessageBox.Show("Bạn có chắc nộp bài?", "Cảnh Báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
            {
                this.ChamDiem();
                if (this.next >= Program.Tong - 1)
                {
                    this.next = 0;
                    string str = "Kết quả của " + Home.thongtin + " = Đúng " + (object)this.So_Cau_Dung + "câu; Sai " + (object)this.So_Cau_Sai + " câu => Điểm: " + (1000.0 / (double)(this.So_Cau_Dung + this.So_Cau_Sai) * (double)this.So_Cau_Dung).ToString("#,##0.0");
                    this.luuDiem(str);
                    int num = (int)MessageBox.Show(str);
                    this.So_Cau_Dung = 0;
                    this.So_Cau_Sai = 0;
                }
                else
                    ++this.next;
                this.SetUp(Program.Lessons[this.next]);
            }
            this.chotat = false;
        }

        private void luuDiem(string diem)
        {
            TextWriter textWriter1 = (TextWriter)new StreamWriter(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip//diem.html"), true);
            string shortDateString = DateTime.Now.ToShortDateString();
            textWriter1.WriteLine("<p style=\"text-align:center;\">" + shortDateString + ":" + diem + "</p>");
            textWriter1.Close();
            TextWriter textWriter2 = (TextWriter)new StreamWriter(Path.Combine(System.Windows.Forms.Application.StartupPath, "diem.html"), true);
            textWriter2.WriteLine("<p style=\"text-align:center;\">" + shortDateString + ":" + diem + "</p>");
            textWriter2.Close();
            string str1 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), nameof(diem));
            if (!Directory.Exists(str1))
                Directory.CreateDirectory(str1);
            string path = Path.Combine(str1, "Word.html");
            if (File.Exists(path))
            {
                TextWriter textWriter3 = (TextWriter)new StreamWriter(path, true);
                textWriter3.WriteLine("<p style=\"text-align:center;\">" + shortDateString + ":" + diem + "</p>");
                textWriter3.Close();
            }
            else
            {
                TextWriter textWriter4 = (TextWriter)new StreamWriter(path, true);
                string str2 = "<P style=\"text-align:center;color:red;font-size:160%\"> Điều Kiện để được thi lại miễn phí là phải học hành chăm chỉ, gửi điểm thường xuyên cho thầy</p><P style=\"text-align:center;\">Khi Có Điểm bạn copy Điểm này gửi cho thầy qua 1 trong các face sau: </p> <P style=\"text-align:center;\">================================================================== </p> <P style=\"text-align:center;\"><a href=\"https://www.facebook.com/tinhoc.hoangkha.17\" target=\"_blank\">1. Minh Quang Vũ </a>  </p><P style=\"text-align:center;\"><a href=\"https://www.facebook.com/truongthinh.19\" target=\"_blank\">2. Trường Thịnh</a>  </p> <P style=\"text-align:center;\"><a href=\"https://www.facebook.com/tinhocmos900\" target=\"_blank\">3. Nguyễn Cữu Đàm</a>  </p><P style=\"text-align:center;\">================================================================== </p> ";
                textWriter4.WriteLine(str2);
                textWriter4.WriteLine("<p style=\"text-align:center;\">" + shortDateString + ":" + diem + "</p>");
                textWriter4.Close();
            }
        }

        private void ChamDiem()
        {
            for (int cau = 0; cau < this.Tong_So_Cau; ++cau)
            {
                if (!this.CacCauDaCheck[cau])
                {
                    if (this.CheckCauLon(cau) == "True")
                        ++this.So_Cau_Dung;
                    else
                        ++this.So_Cau_Sai;
                }
            }
            this.chotat = true;
            object obj = (object)WdSaveOptions.wdDoNotSaveChanges;
            while (this.a.Documents.Count >= 1)
            {
                object SaveChanges = obj;
                object missing1 = System.Type.Missing;
                object missing2 = System.Type.Missing;
                object Index = (object)1;
                // ISSUE: reference to a compiler-generated method
                this.a.Documents[ref Index].Close(ref SaveChanges, ref missing1, ref missing2);
            }
            this.chotat = false;
        }

        private void SetUp(int Index)
        {
            string[] directories = Directory.GetDirectories(Path.Combine(System.Windows.Forms.Application.StartupPath, "Zip\\MaHoa"));
            this.works = directories.Length;
            Array.Sort<string>(directories, (IComparer<string>)StringComparer.InvariantCulture);
            this.workIndex = Index;
            this.pathWork = directories[this.workIndex];
            this.pathRun = Path.Combine(System.Windows.Forms.Application.StartupPath, "Zip\\Tam");
            this.ImageFile = Directory.GetFiles(Path.Combine(this.pathWork, "Test"));
            this.checkedListBox1.Items.Clear();
            for (int index = 0; index < this.ImageFile.Length; ++index)
                this.checkedListBox1.Items.Add((object)(index + 1).ToString());
            this.pathFileOfficeMaHoa = Path.Combine(this.pathWork, "Source\\Main");
            this.pathFileOfficeMaHoa = Directory.GetFiles(this.pathFileOfficeMaHoa)[0];
            this.pathFileOffice = Path.Combine(System.Windows.Forms.Application.StartupPath, "Zip\\Tam\\A\\");
            this.pathFileOffice = Path.Combine(this.pathFileOffice, Path.GetFileName(this.pathFileOfficeMaHoa));
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            string path = Path.Combine(this.pathWork, "Source\\Data");
            try
            {
                foreach (string file in Directory.GetFiles(path))
                    File.Copy(file, Path.Combine(folderPath, Path.GetFileName(file)), true);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Thư mục Document của bạn không truy cập được, các file dữ liệu cần cho bài làm sẽ ở thư mục Data" + ex.Message);
                foreach (string file in Directory.GetFiles(path))
                    File.Copy(file, Path.Combine(System.Windows.Forms.Application.StartupPath + "\\DATA", Path.GetFileName(file)), true);
            }
            this.Tong_So_Cau = this.ImageFile.Length;
            this.CacCauDaCheck = new bool[this.Tong_So_Cau];
            for (int index = 0; index < this.Tong_So_Cau; ++index)
                this.CacCauDaCheck[index] = false;
            switch (this.workIndex)
            {
                case 0:
                    this.pictureBox2.Image = (Image)Resources.E01;
                    this.pictureBox1.Image = (Image)Resources.V01;
                    break;
                case 1:
                    this.pictureBox2.Image = (Image)Resources.E02;
                    this.pictureBox1.Image = (Image)Resources.V02;
                    break;
                case 2:
                    this.pictureBox2.Image = (Image)Resources.E03;
                    this.pictureBox1.Image = (Image)Resources.V03;
                    break;
                case 3:
                    this.pictureBox2.Image = (Image)Resources.E04;
                    this.pictureBox1.Image = (Image)Resources.V04;
                    break;
                case 4:
                    this.pictureBox2.Image = (Image)Resources.E05;
                    this.pictureBox1.Image = (Image)Resources.V05;
                    break;
                case 5:
                    this.pictureBox2.Image = (Image)Resources.E06;
                    this.pictureBox1.Image = (Image)Resources.V06;
                    break;
                case 6:
                    this.pictureBox2.Image = (Image)Resources.E07;
                    this.pictureBox1.Image = (Image)Resources.V07;
                    break;
                case 7:
                    this.pictureBox2.Image = (Image)Resources.E08;
                    this.pictureBox1.Image = (Image)Resources.V08;
                    break;
                case 8:
                    this.pictureBox2.Image = (Image)Resources.E09;
                    this.pictureBox1.Image = (Image)Resources.V09;
                    break;
                case 9:
                    this.pictureBox2.Image = (Image)Resources.E10;
                    this.pictureBox1.Image = (Image)Resources.V10;
                    break;
                case 10:
                    this.pictureBox2.Image = (Image)Resources.E11;
                    this.pictureBox1.Image = (Image)Resources.V11;
                    break;
                case 11:
                    this.pictureBox2.Image = (Image)Resources.E12;
                    this.pictureBox1.Image = (Image)Resources.V12;
                    break;
                case 12:
                    this.pictureBox2.Image = (Image)Resources.E13;
                    this.pictureBox1.Image = (Image)Resources.V13;
                    break;
                case 13:
                    this.pictureBox2.Image = (Image)Resources.E14;
                    this.pictureBox1.Image = (Image)Resources.V14;
                    break;
            }
            TextReader textReader = (TextReader)new StreamReader(Path.Combine(this.pathWork, "a.txt"));
            this.lsViTri = new List<int>();
            for (string s = textReader.ReadLine(); s != null; s = textReader.ReadLine())
                this.lsViTri.Add(int.Parse(s));
            textReader.Close();
            this.Cau_So = 0;
            this.checkedListBox1.SelectedIndex = this.Cau_So;
            Home.DecryptFile(this.pathFileOfficeMaHoa, this.pathFileOffice);
            object pathFileOffice = (object)this.pathFileOffice;
            object missing1 = System.Type.Missing;
            object missing2 = System.Type.Missing;
            object missing3 = System.Type.Missing;
            object PasswordDocument = (object)"271565234";
            object missing4 = System.Type.Missing;
            object missing5 = System.Type.Missing;
            object missing6 = System.Type.Missing;
            object missing7 = System.Type.Missing;
            object missing8 = System.Type.Missing;
            object missing9 = System.Type.Missing;
            object missing10 = System.Type.Missing;
            object missing11 = System.Type.Missing;
            object missing12 = System.Type.Missing;
            object missing13 = System.Type.Missing;
            object XMLTransform = (object)Missing.Value;
            // ISSUE: reference to a compiler-generated method
            this.d = this.a.Documents.Open(ref pathFileOffice, ref missing1, ref missing2, ref missing3, ref PasswordDocument, ref missing4, ref missing5, ref missing6, ref missing7, ref missing8, ref missing9, ref missing10, ref missing11, ref missing12, ref missing13, ref XMLTransform);
            this.soLanReSet = 0;
            for (int index = 0; index < this.CacCauDaCheck.Length; ++index)
                this.CacCauDaCheck[index] = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.chotat = true;
            while (this.a.Documents.Count >= 1)
            {
                object SaveChanges = (object)false;
                object missing1 = System.Type.Missing;
                object missing2 = System.Type.Missing;
                object Index = (object)1;
                // ISSUE: reference to a compiler-generated method
                this.a.Documents[ref Index].Close(ref SaveChanges, ref missing1, ref missing2);
            }
            object SaveChanges1 = (object)Missing.Value;
            object OriginalFormat = (object)Missing.Value;
            object RouteDocument = (object)Missing.Value;
            // ISSUE: reference to a compiler-generated method
            this.a.Quit(ref SaveChanges1, ref OriginalFormat, ref RouteDocument);
            this.chotat = false;
            this.Close();
        }

        private void timer1_Tick(object sender, EventArgs e) => this.label1.Text = (this.timeStrart - DateTime.Now).ToString().Substring(3, 5);

        private void LoadCau(int Cau)
        {
            this.checkedListBox1.SelectedIndex = Cau;
            foreach (string file in Directory.GetFiles(this.pathRun))
                File.Delete(file);
            this.ImageFile[this.checkedListBox1.SelectedIndex].Replace(Path.Combine(this.pathWork, "Test"), this.pathRun);
            Home.DecryptFile(this.pathFileOfficeMaHoa, this.pathFileOffice);
            object pathFileOffice = (object)this.pathFileOffice;
            object missing1 = System.Type.Missing;
            object missing2 = System.Type.Missing;
            object missing3 = System.Type.Missing;
            object PasswordDocument = (object)"271565234";
            object missing4 = System.Type.Missing;
            object missing5 = System.Type.Missing;
            object missing6 = System.Type.Missing;
            object missing7 = System.Type.Missing;
            object missing8 = System.Type.Missing;
            object missing9 = System.Type.Missing;
            object missing10 = System.Type.Missing;
            object missing11 = System.Type.Missing;
            object missing12 = System.Type.Missing;
            object missing13 = System.Type.Missing;
            object XMLTransform = (object)Missing.Value;
            // ISSUE: reference to a compiler-generated method
            this.d = this.a.Documents.Open(ref pathFileOffice, ref missing1, ref missing2, ref missing3, ref PasswordDocument, ref missing4, ref missing5, ref missing6, ref missing7, ref missing8, ref missing9, ref missing10, ref missing11, ref missing12, ref missing13, ref XMLTransform);
        }

        private string CheckCauLon(int cau)
        {
            try
            {
                switch (this.workIndex)
                {
                    //case 0:
                    //    return CheckWork1.Check(cau, this.a, this.d);
                    //case 1:
                    //    return CheckWork2.Check(cau, this.a, this.d);
                    //case 2:
                    //    return CheckWork3.Check(cau, this.a, this.d);
                    //case 3:
                    //    return CheckWork4.Check(cau, this.a, this.d);
                    //case 4:
                    //    return CheckWork5.Check(cau, this.a, this.d);
                    //case 5:
                    //    return CheckWork6.Check(cau, this.a, this.d);
                    //case 6:
                    //    return CheckWork7.Check(cau, this.a, this.d);
                    //case 7:
                    //    return CheckWork8.Check(cau, this.a, this.d);
                    //case 8:
                    //    return CheckWork9.Check(cau, this.a, this.d);
                    //case 9:
                    //    return CheckWork10.Check(cau, this.a, this.d);
                    //case 10:
                    //    return CheckWork11.Check(cau, this.a, this.d);
                    //case 11:
                    //    return CheckWork12.Check(cau, this.a, this.d);
                    //case 12:
                    //    return CheckWork13.Check(cau, this.a, this.d);
                    //case 13:
                    //    return CheckWork14.Check(cau, this.a, this.d);
                    //case 14:
                    //    return CheckWork15.Check(cau, this.a, this.d);
                    //case 15:
                    //    return CheckWork16.Check(cau, this.a, this.d);
                    //case 16:
                    //    return CheckWork17.Check(cau, this.a, this.d);
                    //case 17:
                    //    return CheckWork18.Check(cau, this.a, this.d);
                    //case 18:
                    //    return CheckWork19.Check(cau, this.a, this.d);
                    //case 19:
                    //    return CheckWork20.Check(cau, this.a, this.d);
                    //case 20:
                    //    return CheckWork21.Check(cau, this.a, this.d);
                    //case 21:
                    //    return CheckWork22.Check(cau, this.a, this.d);
                    //case 22:
                    //    return CheckWork23.Check(cau, this.a, this.d);
                    //case 23:
                    //    return CheckWork24.Check(cau, this.a, this.d);
                    //case 24:
                    //    return CheckWork25.Check(cau, this.a, this.d);
                    //case 25:
                    //    return CheckWork26.Check(cau, this.a, this.d);
                    //case 26:
                    //    return CheckWork27.Check(cau, this.a, this.d);
                    //case 27:
                    //    return CheckWork28.Check(cau, this.a, this.d);
                    default:
                        return "đề tham khảo";
                }
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.Items.Count > 10)
            {
                int num = (int)MessageBox.Show("Quá 10 câu không biết rồi bạn gọi thầy đi");
            }
            else
            {
                TextWriter textWriter = (TextWriter)new StreamWriter(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\luu"), true);
                textWriter.WriteLine((Program.Lessons[this.next] + 1).ToString() + "|" + (object)(this.checkedListBox1.SelectedIndex + 1));
                textWriter.Close();
                this.loadcaucanhoi();
            }
        }

        private void pictureBox2_MouseHover(object sender, EventArgs e) => this.panel2.Focus();

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                this.next = int.Parse(this.textBox2.Text) - 2;
                if (this.next >= -1)
                {
                    this.submit();
                }
                else
                {
                    int num = (int)MessageBox.Show("so lon hon 0");
                    this.textBox2.Focus();
                }
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("nhap so");
                this.textBox2.Focus();
            }
        }

        private void buttonHelp_Click(object sender, EventArgs e)
        {
            string str = Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\tam\\1.avi");
            string fileNameSource = Path.Combine(this.pathWork, "help\\" + (object)(this.checkedListBox1.SelectedIndex + 1));
            try
            {
                Home.DecryptFile(fileNameSource, str);
                Process.Start(str);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Tắt Video help trước khi mở help mới |" + ex.Message);
            }
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

        private void buttonxhdh_Click(object sender, EventArgs e)
        {
            if (this.buttonxhdh.Text == "Hướng dẫn sử dụng")
            {
                this.buttonxhdh.Text = "Về chế độ làm bài";
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
                this.buttonxhdh.Text = "Hướng dẫn sử dụng";
                this.TopMost = true;
            }
        }

        private void buttonShowHide_Click(object sender, EventArgs e)
        {
            if (this.TopMost)
            {
                this.TopMost = false;
                this.buttonShowHide.Text = "Tự Làm";
                Process.Start("https://meet.google.com/uay-ywnp-xzs");
            }
            else
            {
                this.TopMost = true;
                this.buttonShowHide.Text = "Gọi Thầy";
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (this.comboBox1.Items.Count == 0)
            {
                int num1 = (int)MessageBox.Show("Bạn không có câu cần hỏi");
            }
            else if (this.comboBox1.Text == "")
            {
                int num2 = (int)MessageBox.Show("Chọn Project");
            }
            else
            {
                string text = this.comboBox1.Text;
                this.comboBox1.Items.RemoveAt(this.comboBox1.SelectedIndex);
                TextWriter textWriter = (TextWriter)new StreamWriter(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\Luu"), false);
                for (int index = 0; index < this.comboBox1.Items.Count; ++index)
                    textWriter.WriteLine(this.comboBox1.Items[index].ToString());
                textWriter.Close();
                this.submit(int.Parse(text.Split('|')[0]) - 1);
            }
        }

        public void submit(int pro)
        {
            if (MessageBox.Show("Bạn có chắc nộp bài?", "Cảnh Báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
            {
                for (int cau = 0; cau < this.Tong_So_Cau; ++cau)
                {
                    if (!this.CacCauDaCheck[cau])
                    {
                        if (this.CheckCauLon(cau) == "True")
                            ++this.So_Cau_Dung;
                        else
                            ++this.So_Cau_Sai;
                    }
                }
                this.Cau_So = 0;
                this.checkedListBox1.SelectedIndex = 0;
                this.pathReset = Path.Combine(this.pathWork, "Source\\Sub\\" + this.Cau_So.ToString());
                this.chotat = true;
                while (this.a.Documents.Count >= 1)
                {
                    object SaveChanges = (object)false;
                    object missing1 = System.Type.Missing;
                    object missing2 = System.Type.Missing;
                    object Index = (object)1;
                    // ISSUE: reference to a compiler-generated method
                    this.a.Documents[ref Index].Close(ref SaveChanges, ref missing1, ref missing2);
                }
                this.chotat = false;
                this.SetUp(pro);
                this.soLanReSet = 0;
                for (int index = 0; index < this.CacCauDaCheck.Length; ++index)
                    this.CacCauDaCheck[index] = false;
            }
            this.chotat = false;
            this.Cau_So = 0;
            this.checkedListBox1.SelectedIndex = 0;
        }

        private void pictureBox1_MouseHover(object sender, EventArgs e) => this.panel1.Focus();
    }
}
