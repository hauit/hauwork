// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Form1
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

//using Microsoft.Office.Interop.Excel;
using MOS_EXCEL_LEARN;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
  public class Form1 : Form
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
    private int tong_so_cau = School.Tong();
    private int cau_User = 1;
    private bool chotat = false;
    public object readOnly = (object) false;
    public object isVisible = (object) true;
    public object missing = (object) Missing.Value;
    private enviroment paramater;
    private Help h;
    private int[] Diem;
    private IContainer components = (IContainer) null;
    private Button buttonExit;
    private Button buttonNext;
    private Button buttonReset;
    private Button buttonEV;
    private Button buttonHelp;
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
          str1 = str1 + (object) (index + 1) + ",";
      }
      string str2 = num.ToString() + "Câu," + str1;
      string str3 = "<p>" + DateTime.Now.ToShortDateString() + "," + str2 + "</p>";
      if (num <= 1)
        return;
      string str4 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Diem");
      if (!Directory.Exists(str4))
        Directory.CreateDirectory(str4);
      string path = Path.Combine(str4, "hoc_Excel.html");
      if (File.Exists(path))
      {
        TextWriter textWriter = (TextWriter) new StreamWriter(path, true);
        textWriter.WriteLine(str3);
        textWriter.Close();
      }
      else
      {
        TextWriter textWriter = (TextWriter) new StreamWriter(path, true);
        string str5 = "<P style=\"text-align:center;color:red;font-size:160%\"> Điều Kiện để được thi lại miễn phí là phải học hành chăm chỉ, gửi điểm thường xuyên cho thầy</p><P style=\"text-align:center;\">Khi Có Điểm bạn copy Điểm này gửi cho thầy qua 1 trong các face sau: </p> <P style=\"text-align:center;\">================================================================== </p> <P style=\"text-align:center;\"><a href=\"https://www.facebook.com/tinhoc.hoangkha.17\" target=\"_blank\">1. Minh Quang Vũ </a>  </p><P style=\"text-align:center;\"><a href=\"https://www.facebook.com/truongthinh.19\" target=\"_blank\">2. Trường Thịnh</a>  </p> <P style=\"text-align:center;\"><a href=\"https://www.facebook.com/tinhocmos900\" target=\"_blank\">3. Nguyễn Cữu Đàm</a>  </p><P style=\"text-align:center;\">================================================================== </p> ";
        textWriter.WriteLine(str5);
        textWriter.WriteLine(str3);
        textWriter.Close();
      }
    }

    private void LuuDanhDau()
    {
      TextWriter textWriter = (TextWriter) new StreamWriter(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip//luudanhdau"));
      for (int index = 0; index < this.comboBoxCauDaDanhDau.Items.Count; ++index)
        textWriter.WriteLine(this.comboBoxCauDaDanhDau.Items[index].ToString());
      textWriter.Close();
    }

    private void loadDanhDau()
    {
      string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "zip//luudanhdau");
      if (!File.Exists(path))
        return;
      TextReader textReader = (TextReader) new StreamReader(path);
      for (string str = textReader.ReadLine(); str != null; str = textReader.ReadLine())
        this.comboBoxCauDaDanhDau.Items.Add((object) str);
      textReader.Close();
    }

    public Form1()
    {
      //this.h = new Help(); TODO: uncomment after finish Help form
      this.InitializeComponent();
      this.FormClosed += new FormClosedEventHandler(this.Form1_FormClosed);
      try
      {
        this.a = (Microsoft.Office.Interop.Excel.Application) Activator.CreateInstance(System.Type.GetTypeFromCLSID(new Guid("00024500-0000-0000-C000-000000000046")));
        this.load_cau_hoi(this.cau_User);
        // ISSUE: method pointer
        // ISSUE: object of a compiler-generated type is created
        ((Microsoft.Office.Interop.Excel.AppEvents_Event)this.a)
        .WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(this.a_WorkbookBeforeClose);
        }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    private void a_WorkbookBeforeClose(Microsoft.Office.Interop.Excel.Workbook Wb, ref bool Cancel)
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
      object SaveChanges = (object)Microsoft.Office.Interop.Excel.XlSaveAction.xlDoNotSaveChanges;
      try
      {
        this.chotat = true;
        while (this.a.Windows.Count >= 1)
        {
          // ISSUE: reference to a compiler-generated method
          this.a.Windows[(object) 1].Close(SaveChanges, System.Type.Missing, System.Type.Missing);
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
      object SaveChanges = (object)Microsoft.Office.Interop.Excel.XlSaveAction.xlDoNotSaveChanges;
      this.chotat = true;
      while (this.a.Windows.Count >= 1)
      {
        // ISSUE: reference to a compiler-generated method
        this.a.Windows[(object) 1].Close(SaveChanges, System.Type.Missing, System.Type.Missing);
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
      //this.h.refresh();  TODO: uncomment after finish Help form
      this.a.Visible = true;
    }

    private void loadTongSoCau()
    {
      for (int index = 0; index < School.Mn.Length; ++index)
      {
        string[] files = Directory.GetFiles(Path.Combine(System.Windows.Forms.Application.StartupPath, "Data\\Sec_" + (object) index + "\\hinh\\E"));
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
        this.pictureBox1.Location = new Point(0, 0);
        this.pictureBox1.Width = this.Width - this.buttonEV.Width / 3;
        this.comboBoxCauNext.Text = (this.cau_User + 1).ToString();
        this.labelCauHienTai.Text = "Câu " + this.cau_User.ToString();
        for (int index = 0; index < this.tong_so_cau; ++index)
          this.comboBoxCauNext.Items.Add((object) (index + 1));
        this.Copy(Path.Combine(System.Windows.Forms.Application.StartupPath, "data\\Doc"), Environment.GetFolderPath(Environment.SpecialFolder.Personal));
        this.a.Visible = true;
        this.a.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
        this.a.Top = 0.0;
        this.a.Left = 0.0;
        this.a.Width = (double) this.screen_width;
        this.a.Height = (double) (this.screen_height * 3 / 5);
        this.TopMost = true;
        this.currenTime = DateTime.Now;
        this.setDiem();
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
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
      this.paramater.Dest_file_Word_Name = Path.Combine(System.Windows.Forms.Application.StartupPath, "Word\\" + this.paramater.section.ToString() + "_" + this.paramater.quesion.ToString() + ".docx");
      this.paramater.DeTiengAnh = Path.Combine(System.Windows.Forms.Application.StartupPath, "tam\\" + cau_hoi_so.ToString() + "E");
      this.paramater.DeTiengViet = Path.Combine(System.Windows.Forms.Application.StartupPath, "tam\\" + cau_hoi_so.ToString() + "V");
      this.paramater.Dest_file_help_Name = Path.Combine(System.Windows.Forms.Application.StartupPath, "tam\\help");
      this.paramater.Dest_file_help_video_Name = Path.Combine(System.Windows.Forms.Application.StartupPath, "tam\\Vhelp");
      Home.DecryptFile(this.paramater.Source_file_word_path, this.paramater.Dest_file_Word_Name.ToString());
      Home.DecryptFile(this.paramater.Source_de_En, this.paramater.DeTiengAnh);
      Home.DecryptFile(this.paramater.Source_de_Vn, this.paramater.DeTiengViet);
      if (File.Exists(this.paramater.Source_file_help_video_path))
      {
        Home.DecryptFile(this.paramater.Source_file_help_video_path, this.paramater.Dest_file_help_video_Name);
      }
      else
      {
        int num = (int) MessageBox.Show("Chua co Video help");
      }
      this.pictureBox1.Image = Image.FromFile(this.paramater.DeTiengAnh);
      this.pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
      this.pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
      this.pictureBox1.Height = this.pictureBox1.Height * this.Width / this.pictureBox1.Width;
      this.pictureBox1.Width = this.Width - this.buttonCheck.Width / 2;
      this.pictureBox1.Refresh();
      // ISSUE: reference to a compiler-generated method
      this.d = this.a.Workbooks.Open(this.paramater.Dest_file_Word_Name, System.Type.Missing, System.Type.Missing, System.Type.Missing, (object) "271565234", System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
    }

    private void setDefalt(enviroment par)
    {
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
        int num = (int) MessageBox.Show("Nhấn nút ESC và đóng các hộp thoại để tiếp tục");
      }
      return str;
    }

    private string Check(enviroment par)
    {
      string str = "";
      try
      {
        switch (par.section)
        {
          case 0:
            //str = Sec0.CheckCau(par.quesion, this.a, this.d);
            break;
          case 1:
            //str = Sec1.CheckCau(par.quesion, this.a, this.d);
            break;
          case 2:
            //str = Sec2.CheckCau(par.quesion, this.a, this.d);
            break;
          case 3:
            //str = Sec3.CheckCau(par.quesion, this.a, this.d);
            break;
          case 4:
            //str = Sec4.CheckCau(par.quesion, this.a, this.d);
            break;
          case 5:
            //str = Sec5.CheckCau(par.quesion, this.a, this.d);
            break;
          case 6:
            //str = Sec6.CheckCau(par.quesion, this.a, this.d);
            break;
          case 7:
            //str = Sec7.CheckCau(par.quesion, this.a, this.d);
            break;
          case 8:
            //str = Sec8.CheckCau(par.quesion, this.a, this.d);
            break;
          case 9:
            //str = Sec9.CheckCau(par.quesion, this.a, this.d);
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
        int num = (int) MessageBox.Show("Đống tất cả các hộp thoại đang mở trước");
      }
      return str;
    }

    private void buttonRefresh_Click(object sender, EventArgs e)
    {
      try
      {
        this.a.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
        this.a.Top = 0.0;
        this.a.Left = 0.0;
        this.a.Width = (double) this.screen_width;
        this.a.Height = (double) (this.screen_height * 3 / 5);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show("Bạn đã tắt MS Word! Tắt chương trình và làm lại");
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
        int num2 = (int) MessageBox.Show("Chọn file excel nhấn nút ESC trước khi reset");
        num1 = 1;
      }
      if (num1 == 1)
        return;
      this.load_cau_hoi(this.cau_User);
      this.a.Visible = true;
    }

    private void timer1_Tick(object sender, EventArgs e)
    {
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

    private void pictureBox1_MouseHover(object sender, EventArgs e)
    {
    }

    private void buttonHelp_Click(object sender, EventArgs e)
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

    private void buttonSummary_Click(object sender, EventArgs e)
    {
      int selectedIndex = this.comboBoxCauDaDanhDau.SelectedIndex;
      if (selectedIndex >= 0 && selectedIndex < this.comboBoxCauDaDanhDau.Items.Count)
      {
        this.comboBoxCauDaDanhDau.Items.RemoveAt(selectedIndex);
      }
      else
      {
        int num = (int) MessageBox.Show("Chọn câu đã đánh dấu cần xóa");
      }
    }

    private void buttonZoom_Click(object sender, EventArgs e)
    {
      this.comboBoxCauDaDanhDau.Items.Add((object) this.cau_User);
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
      //this.h.refresh();  TODO: uncomment after finish Help form
      this.a.Visible = true;
    }

    private void buttonVideoHelp_Click(object sender, EventArgs e)
    {
      if (this.buttonVideoHelp.Text == "HDSD phần mềm")
      {
        this.buttonVideoHelp.Text = "Về chế độ làm bài";
        this.TopMost = false;
        try
        {
          Process.Start(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\hdh.mp4"));
        }
        catch (Exception ex)
        {
          int num = (int) MessageBox.Show("đổi chương trinh mặt định xem Video khác" + ex.Message);
        }
      }
      else
      {
        this.buttonVideoHelp.Text = "HDSD phần mềm";
        this.TopMost = true;
      }
    }

    private void button2_Click_1(object sender, EventArgs e)
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
      string str = Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\1.avi");
      string fileHelpVideoPath = this.paramater.Source_file_help_video_path;
      try
      {
        Home.DecryptFile(fileHelpVideoPath, str);
        Process.Start(str);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show("Tắt Video help trước khi mở help mới");
      }
    }

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
      this.buttonHelp = new Button();
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
      ((ISupportInitialize) this.pictureBox1).BeginInit();
      this.panel1.SuspendLayout();
      this.SuspendLayout();
      this.buttonExit.BackColor = Color.Transparent;
      this.buttonExit.Location = new Point(100, 68);
      this.buttonExit.Name = "buttonExit";
      this.buttonExit.Size = new Size(75, 23);
      this.buttonExit.TabIndex = 0;
      this.buttonExit.Text = "Exit";
      this.buttonExit.UseVisualStyleBackColor = false;
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
      this.buttonHelp.Location = new Point(365, 234);
      this.buttonHelp.Name = "buttonHelp";
      this.buttonHelp.Size = new Size(114, 23);
      this.buttonHelp.TabIndex = 4;
      this.buttonHelp.Text = "Về chế độ thường";
      this.buttonHelp.UseVisualStyleBackColor = true;
      this.buttonHelp.Click += new EventHandler(this.buttonHelp_Click);
      this.buttonRefresh.Location = new Point(365, 263);
      this.buttonRefresh.Name = "buttonRefresh";
      this.buttonRefresh.Size = new Size(87, 23);
      this.buttonRefresh.TabIndex = 5;
      this.buttonRefresh.Text = "Phục hồi Excel";
      this.buttonRefresh.UseVisualStyleBackColor = true;
      this.buttonRefresh.Click += new EventHandler(this.buttonRefresh_Click);
      this.richTextBox1.Font = new System.Drawing.Font("Times New Roman", 18f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
      this.richTextBox1.Location = new Point(61, 296);
      this.richTextBox1.Name = "richTextBox1";
      this.richTextBox1.ReadOnly = true;
      this.richTextBox1.Size = new Size(103, 82);
      this.richTextBox1.TabIndex = 6;
      this.richTextBox1.Text = "";
      this.buttonCheck.BackColor = Color.FromArgb((int) byte.MaxValue, 192, 128);
      this.buttonCheck.Location = new Point(285, 184);
      this.buttonCheck.Name = "buttonCheck";
      this.buttonCheck.Size = new Size(97, 23);
      this.buttonCheck.TabIndex = 7;
      this.buttonCheck.Text = "Check đúng/sai";
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
      this.labelKQ.Font = new System.Drawing.Font("Times New Roman", 14.25f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
      this.labelKQ.ForeColor = SystemColors.Highlight;
      this.labelKQ.Location = new Point(516, 220);
      this.labelKQ.Name = "labelKQ";
      this.labelKQ.Size = new Size(0, 21);
      this.labelKQ.TabIndex = 10;
      this.labelCauHienTai.AutoSize = true;
      this.labelCauHienTai.Font = new System.Drawing.Font("Times New Roman", 14.25f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
      this.labelCauHienTai.ForeColor = Color.DeepPink;
      this.labelCauHienTai.Location = new Point(569, 274);
      this.labelCauHienTai.Name = "labelCauHienTai";
      this.labelCauHienTai.Size = new Size(72, 21);
      this.labelCauHienTai.TabIndex = 12;
      this.labelCauHienTai.Text = "Câu 150";
      this.comboBoxCauNext.BackColor = SystemColors.MenuText;
      this.comboBoxCauNext.Font = new System.Drawing.Font("Times New Roman", 14.25f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
      this.comboBoxCauNext.ForeColor = Color.Coral;
      this.comboBoxCauNext.FormattingEnabled = true;
      this.comboBoxCauNext.Location = new Point(720, 369);
      this.comboBoxCauNext.Name = "comboBoxCauNext";
      this.comboBoxCauNext.Size = new Size(52, 29);
      this.comboBoxCauNext.TabIndex = 13;
      this.comboBoxCauNext.Text = "1";
      this.panel1.AutoScroll = true;
      this.panel1.Controls.Add((Control) this.pictureBox1);
      this.panel1.Location = new Point(573, 35);
      this.panel1.Name = "panel1";
      this.panel1.Size = new Size(282, 172);
      this.panel1.TabIndex = 15;
      this.buttonZoom.BackColor = Color.FromArgb(192, (int) byte.MaxValue, (int) byte.MaxValue);
      this.buttonZoom.Location = new Point(487, 288);
      this.buttonZoom.Name = "buttonZoom";
      this.buttonZoom.Size = new Size(76, 23);
      this.buttonZoom.TabIndex = 21;
      this.buttonZoom.Text = "Đánh dấu";
      this.buttonZoom.UseVisualStyleBackColor = false;
      this.buttonZoom.Click += new EventHandler(this.buttonZoom_Click);
      this.buttonSummary.Location = new Point(285, 263);
      this.buttonSummary.Name = "buttonSummary";
      this.buttonSummary.Size = new Size(42, 23);
      this.buttonSummary.TabIndex = 23;
      this.buttonSummary.Text = "Xóa";
      this.buttonSummary.UseVisualStyleBackColor = true;
      this.buttonSummary.Click += new EventHandler(this.buttonSummary_Click);
      this.buttonVideoHelp.BackColor = Color.FromArgb((int) byte.MaxValue, (int) byte.MaxValue, 192);
      this.buttonVideoHelp.Location = new Point(545, 218);
      this.buttonVideoHelp.Name = "buttonVideoHelp";
      this.buttonVideoHelp.Size = new Size(108, 23);
      this.buttonVideoHelp.TabIndex = 24;
      this.buttonVideoHelp.Text = "HDSD phần mềm";
      this.buttonVideoHelp.UseVisualStyleBackColor = false;
      this.buttonVideoHelp.Click += new EventHandler(this.buttonVideoHelp_Click);
      this.button2.BackColor = Color.Transparent;
      this.button2.ForeColor = Color.Black;
      this.button2.Location = new Point(389, 374);
      this.button2.Name = "button2";
      this.button2.Size = new Size(77, 23);
      this.button2.TabIndex = 25;
      this.button2.Text = "Gọi Thầy";
      this.button2.UseVisualStyleBackColor = false;
      this.button2.Click += new EventHandler(this.button2_Click_1);
      this.comboBoxCauDaDanhDau.BackColor = Color.White;
      this.comboBoxCauDaDanhDau.Font = new System.Drawing.Font("Times New Roman", 12f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
      this.comboBoxCauDaDanhDau.ForeColor = Color.Red;
      this.comboBoxCauDaDanhDau.FormattingEnabled = true;
      this.comboBoxCauDaDanhDau.Location = new Point(835, 298);
      this.comboBoxCauDaDanhDau.Name = "comboBoxCauDaDanhDau";
      this.comboBoxCauDaDanhDau.Size = new Size(43, 27);
      this.comboBoxCauDaDanhDau.TabIndex = 26;
      this.buttonHelpVideo.BackColor = Color.FromArgb((int) byte.MaxValue, 192, 192);
      this.buttonHelpVideo.Location = new Point(365, 333);
      this.buttonHelpVideo.Name = "buttonHelpVideo";
      this.buttonHelpVideo.Size = new Size(75, 23);
      this.buttonHelpVideo.TabIndex = 30;
      this.buttonHelpVideo.Text = "Chỉ làm bài";
      this.buttonHelpVideo.UseVisualStyleBackColor = false;
      this.buttonHelpVideo.Click += new EventHandler(this.buttonHelpVideo_Click);
      this.AutoScaleDimensions = new SizeF(6f, 13f);
      this.AutoScaleMode = AutoScaleMode.Font;
      this.BackColor = Color.Black;
      this.ClientSize = new Size(1016, 599);
      this.Controls.Add((Control) this.buttonHelpVideo);
      this.Controls.Add((Control) this.labelKQ);
      this.Controls.Add((Control) this.comboBoxCauDaDanhDau);
      this.Controls.Add((Control) this.button2);
      this.Controls.Add((Control) this.buttonZoom);
      this.Controls.Add((Control) this.panel1);
      this.Controls.Add((Control) this.comboBoxCauNext);
      this.Controls.Add((Control) this.labelCauHienTai);
      this.Controls.Add((Control) this.buttonCheck);
      this.Controls.Add((Control) this.richTextBox1);
      this.Controls.Add((Control) this.buttonRefresh);
      this.Controls.Add((Control) this.buttonHelp);
      this.Controls.Add((Control) this.buttonEV);
      this.Controls.Add((Control) this.buttonReset);
      this.Controls.Add((Control) this.buttonNext);
      this.Controls.Add((Control) this.buttonExit);
      this.Controls.Add((Control) this.buttonSummary);
      this.Controls.Add((Control) this.buttonVideoHelp);
      this.FormBorderStyle = FormBorderStyle.None;
      this.Name = nameof (Form1);
      this.Text = nameof (Form1);
      this.TopMost = true;
      this.WindowState = FormWindowState.Maximized;
      this.Load += new EventHandler(this.Form1_Load);
      ((ISupportInitialize) this.pictureBox1).EndInit();
      this.panel1.ResumeLayout(false);
      this.panel1.PerformLayout();
      this.ResumeLayout(false);
      this.PerformLayout();
    }
  }
}
