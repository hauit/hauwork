// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Form2
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

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

namespace WindowsFormsApplication1
{
  public class Form2_cu : Form
  {
    private IContainer components = (IContainer) null;
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

    public Form2_cu()
    {
      this.InitializeComponent();
      this.textBoxUser.LostFocus += new EventHandler(this.textBoxUser_LostFocus);
    }

    private void textBoxUser_LostFocus(object sender, EventArgs e)
    {
    }

    private void buttonOK_Click(object sender, EventArgs e)
    {
      if (this.comboBox1.Visible)
      {
        if (this.comboBox1.Text == "")
        {
          int num = (int) MessageBox.Show("Bạn phải chọn bài test");
        }
        else
        {
          Program.user = "271565";
          Program.pass = this.textBoxPass.Text;
          Program.status = 1;
          this.Close();
        }
      }
      else
      {
        Program.user = "271565";
        Program.pass = this.textBoxPass.Text;
        Program.status = 1;
        this.Close();
      }
    }

    private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) => Program.TypeOfTest = this.comboBox1.SelectedIndex;

    private void Form2_Load(object sender, EventArgs e)
    {
      string str = Path.Combine(Application.StartupPath, "zip\\b");
      try
      {
        if (!System.IO.File.Exists(str))
        {
          int num1 = (int) MessageBox.Show("Chọn file key để active phần mềm này");
          if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
          {
            if (System.IO.File.Exists(this.openFileDialog1.FileName))
            {
              System.IO.File.Copy(this.openFileDialog1.FileName, str);
              System.IO.File.Copy("C:\\MOS\\GOC HO TRO\\NHAC HOC.lnk", Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\nhac hoc.lnk", true);
              this.check(str);
            }
            else
            {
              int num2 = (int) MessageBox.Show("Không tiềm thấy file");
              this.Close();
            }
          }
          else
            this.Close();
        }
        else
          this.check(str);
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show("Không thể copy file, click chuột phải chọn chạy bằng administrator |" + ex.Message);
        this.Close();
      }
    }

    public static DateTime GetDateFromeInternet()
    {
      DateTime dateFromeInternet;
      try
      {
        dateFromeInternet = Form2_cu.GetMicrosoftNistTime();
      }
      catch (Exception ex1)
      {
        try
        {
          dateFromeInternet = Form2_cu.GetDateTime();
        }
        catch (Exception ex2)
        {
          dateFromeInternet = Form2_cu.GetFastestNISTDate();
        }
      }
      return dateFromeInternet;
    }

    public static DateTime GetMicrosoftNistTime() => DateTime.ParseExact(WebRequest.Create("https://www.google.com/search?q=Iraq time").GetResponse().Headers["date"], "ddd, dd MMM yyyy HH:mm:ss 'GMT'", (IFormatProvider) CultureInfo.InvariantCulture.DateTimeFormat, DateTimeStyles.AssumeUniversal);

    public static DateTime GetDateTime()
    {
      DateTime dateTime = DateTime.MinValue;
      HttpWebRequest httpWebRequest = (HttpWebRequest) WebRequest.Create("https://www.microsoft.com/vi-vn");
      httpWebRequest.Method = "GET";
      httpWebRequest.Accept = "text/html, application/xhtml+xml, */*";
      httpWebRequest.UserAgent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)";
      httpWebRequest.ContentType = "application/x-www-form-urlencoded";
      httpWebRequest.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
      HttpWebResponse response = (HttpWebResponse) httpWebRequest.GetResponse();
      if (response.StatusCode == HttpStatusCode.OK)
        dateTime = DateTime.ParseExact(response.Headers["date"], "ddd, dd MMM yyyy HH:mm:ss 'GMT'", (IFormatProvider) CultureInfo.InvariantCulture.DateTimeFormat, DateTimeStyles.AssumeUniversal);
      return dateTime;
    }

    public static DateTime GetFastestNISTDate()
    {
      DateTime fastestNistDate = DateTime.MinValue;
      string[] source = new string[10]
      {
        "nist1-ny.ustiming.org",
        "nist1-nj.ustiming.org",
        "nist1-pa.ustiming.org",
        "time-a.nist.gov",
        "time-b.nist.gov",
        "nist1.aol-va.symmetricom.com",
        "nist1.columbiacountyga.gov",
        "nist1-chi.ustiming.org",
        "nist.expertsmi.com",
        "nist.netservicesgroup.com"
      };
      Random rnd = new Random();
      foreach (string hostname in ((IEnumerable<string>) source).OrderBy<string, double>((Func<string, double>) (s => rnd.NextDouble())).Take<string>(5))
      {
        try
        {
          string str = string.Empty;
          using (StreamReader streamReader = new StreamReader((Stream) new TcpClient(hostname, 13).GetStream()))
            str = streamReader.ReadToEnd();
          if (!string.IsNullOrEmpty(str))
          {
            string[] strArray1 = str.Split(' ');
            if (strArray1.Length >= 6 && strArray1[5] == "0")
            {
              string[] strArray2 = strArray1[1].Split('-');
              string[] strArray3 = strArray1[2].Split(':');
              fastestNistDate = new DateTime(Convert.ToInt32(strArray2[0]) + 2000, Convert.ToInt32(strArray2[1]), Convert.ToInt32(strArray2[2]), Convert.ToInt32(strArray3[0]), Convert.ToInt32(strArray3[1]), Convert.ToInt32(strArray3[2])).ToLocalTime();
              return fastestNistDate;
            }
          }
        }
        catch
        {
        }
      }
      return fastestNistDate;
    }

    public bool GetMacAddress(string mac)
    {
      foreach (NetworkInterface networkInterface in NetworkInterface.GetAllNetworkInterfaces())
      {
        string str = networkInterface.GetPhysicalAddress().ToString();
        if (str != "" && mac.Contains(str))
          return true;
      }
      return false;
    }

    public string getMac() => ((IEnumerable<NetworkInterface>) NetworkInterface.GetAllNetworkInterfaces()).Where<NetworkInterface>((Func<NetworkInterface, bool>) (nic => nic.OperationalStatus == OperationalStatus.Up)).Select<NetworkInterface, string>((Func<NetworkInterface, string>) (nic => nic.GetPhysicalAddress().ToString())).FirstOrDefault<string>().ToString();

    public bool CheckForInternetConnection(int timeoutMs = 10000, string url = null)
    {
      try
      {
        HttpWebRequest httpWebRequest = (HttpWebRequest) WebRequest.Create(url);
        httpWebRequest.KeepAlive = false;
        httpWebRequest.Timeout = timeoutMs;
        using ((HttpWebResponse) httpWebRequest.GetResponse())
          return true;
      }
      catch
      {
        return false;
      }
    }

    private void check(string KeyPath)
    {
      string str1 = Path.Combine(Application.StartupPath, "Zip\\t");
      if (System.IO.File.Exists(KeyPath))
      {
        try
        {
          Home.DecryptFileT(KeyPath, str1);
        }
        catch (Exception ex)
        {
          int num = (int) MessageBox.Show("Bạn phải chọn đúng file Key của nhà cung cấp |lỗi: " + ex.Message);
          System.IO.File.Delete(KeyPath);
          this.Close();
        }
        try
        {
          TextReader textReader = (TextReader) new StreamReader(str1);
          string[] strArray = textReader.ReadLine().Split('|');
          DateTime dateTime = new DateTime(int.Parse(strArray[2]), int.Parse(strArray[1]), int.Parse(strArray[3]));
          string mac = textReader.ReadLine().Trim();
          string str2 = textReader.ReadLine().Trim();
          Home.thongtin = textReader.ReadLine().Trim();
          string str3 = textReader.ReadLine().Trim();
          textReader.Close();
          System.IO.File.Delete(str1);
          if (str3 != "V1.0")
          {
            int num = (int) MessageBox.Show("Key của bạn sai Version, vào web thimos1000.wordpress.com để tìm hổ trợ");
            System.IO.File.Delete(KeyPath);
            this.Close();
          }
          try
          {
            DateTime dateFromeInternet = Form2_cu.GetDateFromeInternet();
            if (dateFromeInternet < new DateTime(2022, 5, 24))
            {
              if (this.CheckForInternetConnection(url: "https://www.google.com.vn/?hl=vi"))
              {
                DateTime now = DateTime.Now;
                if (now > dateTime || (dateTime - now).Days > 70)
                {
                  int num = (int) MessageBox.Show("Phần mềm hết hạng");
                  System.IO.File.Delete(KeyPath);
                  this.Close();
                }
                else
                {
                  Program.TypeOfTest = !(now.AddDays(2.0) > dateTime) ? 1 : 0;
                  TextWriter textWriter = (TextWriter) new StreamWriter("C:\\MOS\\GOC HO TRO\\ngay_thi_W");
                  textWriter.WriteLine(strArray[3] + "|" + strArray[1] + "|" + strArray[2]);
                  textWriter.Close();
                }
              }
              else
              {
                int num = (int) MessageBox.Show("không có mạng hoặc sai giờ hệ thống hoặc bị chặn bởi tường lửa");
                this.Close();
              }
            }
            else if (dateFromeInternet > dateTime || (dateTime - dateFromeInternet).Days > 70)
            {
              int num = (int) MessageBox.Show("Phần mềm hết hạng");
              System.IO.File.Delete(KeyPath);
              this.Close();
            }
            else
            {
              Program.TypeOfTest = !(dateFromeInternet.AddDays(2.0) > dateTime) ? 1 : 0;
              TextWriter textWriter = (TextWriter) new StreamWriter("C:\\MOS\\GOC HO TRO\\ngay_thi_W");
              textWriter.WriteLine(strArray[3] + "|" + strArray[1] + "|" + strArray[2]);
              textWriter.Close();
            }
          }
          catch (Exception ex)
          {
            int num = (int) MessageBox.Show("không kết nói được với server vào web thimos1000.wordpress.com để tìm hổ trợ|" + ex.Message);
            this.Close();
          }
          if (str2 != "W2019")
          {
            int num = (int) MessageBox.Show("Phần mềm W2019 mà key la " + str2);
            System.IO.File.Delete(KeyPath);
            this.Close();
          }
          try
          {
            if (this.GetMacAddress(mac))
              return;
            int num = (int) MessageBox.Show("Key này không dành cho máy này; vào web thimos1000.wordpress.com để tìm hổ trợ");
            System.IO.File.Delete(KeyPath);
            this.Close();
          }
          catch (Exception ex)
          {
            int num = (int) MessageBox.Show("không truy cập được ID máy |" + ex.Message);
            this.Close();
          }
        }
        catch (Exception ex)
        {
          int num = (int) MessageBox.Show("Không phải file Key |" + ex.Message);
          System.IO.File.Delete(KeyPath);
          this.Close();
        }
      }
      else
      {
        int num1 = (int) MessageBox.Show("không tìm thây file");
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
      ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof (Form2_cu));
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
      this.textBoxUser.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.textBoxUser.Location = new Point(140, 384);
      this.textBoxUser.Name = "textBoxUser";
      this.textBoxUser.Size = new Size(398, 38);
      this.textBoxUser.TabIndex = 1;
      this.textBoxUser.Text = "hocmosonline";
      this.textBoxPass.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
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
      this.label1.Font = new Font("Times New Roman", 15.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label1.ForeColor = Color.Black;
      this.label1.Location = new Point(37, 384);
      this.label1.Name = "label1";
      this.label1.Size = new Size(51, 24);
      this.label1.TabIndex = 5;
      this.label1.Text = "User";
      this.label2.AutoSize = true;
      this.label2.Font = new Font("Times New Roman", 15.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label2.ForeColor = Color.Black;
      this.label2.Location = new Point(35, 437);
      this.label2.Name = "label2";
      this.label2.Size = new Size(100, 24);
      this.label2.TabIndex = 6;
      this.label2.Text = "PassWord";
      this.comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
      this.comboBox1.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
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
      this.label3.Font = new Font("Times New Roman", 15.75f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label3.Location = new Point(35, 490);
      this.label3.Name = "label3";
      this.label3.Size = new Size(79, 24);
      this.label3.TabIndex = 4;
      this.label3.Text = "Lesson:";
      this.label4.AutoSize = true;
      this.label4.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Bold, GraphicsUnit.Point, (byte) 0);
      this.label4.ForeColor = Color.White;
      this.label4.Location = new Point(120, 53);
      this.label4.Name = "label4";
      this.label4.Size = new Size(412, 31);
      this.label4.TabIndex = 7;
      this.label4.Text = "PHẦN LÀM TEST WORD 2019";
      this.label5.AutoSize = true;
      this.label5.Font = new Font("Microsoft Sans Serif", 15.75f, FontStyle.Regular, GraphicsUnit.Point, (byte) 0);
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
      this.BackgroundImage = (Image) componentResourceManager.GetObject("$this.BackgroundImage");
      this.BackgroundImageLayout = ImageLayout.Center;
      this.ClientSize = new Size(698, 586);
      this.Controls.Add((Control) this.label5);
      this.Controls.Add((Control) this.label4);
      this.Controls.Add((Control) this.label3);
      this.Controls.Add((Control) this.comboBox1);
      this.Controls.Add((Control) this.label2);
      this.Controls.Add((Control) this.label1);
      this.Controls.Add((Control) this.buttonOK);
      this.Controls.Add((Control) this.textBoxPass);
      this.Controls.Add((Control) this.textBoxUser);
      this.DoubleBuffered = true;
      this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
      this.Name = nameof (Form2_cu);
      this.StartPosition = FormStartPosition.CenterParent;
      this.Text = "Login";
      this.Load += new EventHandler(this.Form2_Load);
      this.ResumeLayout(false);
      this.PerformLayout();
    }
  }
}
