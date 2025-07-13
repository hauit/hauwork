// Decompiled with JetBrains decompiler
// Type: NHACHOC.Form1
// Assembly: NHACHOC, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 203B0FF7-2323-40ED-93E4-470E5FE69D30
// Assembly location: C:\Users\Admin\Downloads\WE_19\GOC HO TRO\NHAC HOC\NHACHOC.exe

using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using NHACHOC.Properties;

namespace NHACHOC
{
  public class Form1_cu : Form
  {
    private IContainer components = (IContainer) null;
    private Timer timer1;
    private TextBox textBox2;
    private TextBox textBox1;
    private PictureBox pictureBox1;
    private string thongBao = "KIỂM TRA TRÊN WEB XEM LỊCH THI CỦA BẠN CÓ ĐÚNG VỚI LỊCH NÀY KHÔNG?:\r\n";
    private string nhacnho = "Thường Xuyên Ôn bài để kịp thi nha các bạn cố gắn mỗi môn 3 lần trên 950, IIG sẽ không chấp nhận lý do học không kịp nha các bạn\r\n";

    protected override void Dispose(bool disposing)
    {
      if (disposing && this.components != null)
        this.components.Dispose();
      base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
      this.components = (IContainer) new System.ComponentModel.Container();
      ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof (Form1_cu));
      this.timer1 = new Timer(this.components);
      this.textBox2 = new TextBox();
      this.textBox1 = new TextBox();
      this.pictureBox1 = new PictureBox();
      ((ISupportInitialize) this.pictureBox1).BeginInit();
      this.SuspendLayout();
      this.timer1.Interval = 30000;
      this.timer1.Tick += new EventHandler(this.timer1_Tick);
      this.textBox2.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Regular, GraphicsUnit.Point, (byte) 163);
      this.textBox2.Location = new Point(318, 221);
      this.textBox2.Multiline = true;
      this.textBox2.Name = "textBox2";
      this.textBox2.ReadOnly = true;
      this.textBox2.ScrollBars = ScrollBars.Both;
      this.textBox2.Size = new Size(673, 194);
      this.textBox2.TabIndex = 1;
      this.textBox2.Text = "Thường Xuyên Ôn bài để kịp thi nha các bạn cố gắn mỗi môn 3 lần trên 950, IIG sẽ không chấp nhận lý do học không kịp nha các bạn";
      this.textBox1.Font = new Font("Microsoft Sans Serif", 20.25f, FontStyle.Bold, GraphicsUnit.Point, (byte) 163);
      this.textBox1.ForeColor = Color.Red;
      this.textBox1.Location = new Point(12, 21);
      this.textBox1.Multiline = true;
      this.textBox1.Name = "textBox1";
      this.textBox1.ReadOnly = true;
      this.textBox1.Size = new Size(979, 194);
      this.textBox1.TabIndex = 2;
      this.pictureBox1.Image = (Image) Resources._2022_08_25;
      this.pictureBox1.Location = new Point(12, 221);
      this.pictureBox1.Name = "pictureBox1";
      this.pictureBox1.Size = new Size(300, 194);
      this.pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
      this.pictureBox1.TabIndex = 3;
      this.pictureBox1.TabStop = false;
      this.AutoScaleDimensions = new SizeF(6f, 13f);
      this.AutoScaleMode = AutoScaleMode.Font;
      this.ClientSize = new Size(1021, 448);
      this.Controls.Add((Control) this.pictureBox1);
      this.Controls.Add((Control) this.textBox1);
      this.Controls.Add((Control) this.textBox2);
      this.Icon = (Icon) componentResourceManager.GetObject("$this.Icon");
      this.Name = nameof (Form1_cu);
      this.Text = "NHẮC HỌC HÀNH";
      this.Load += new EventHandler(this.Form1_Load);
      ((ISupportInitialize) this.pictureBox1).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();
    }

    public Form1_cu() => this.InitializeComponent();

    private void Form1_Load(object sender, EventArgs e)
    {
      string str1 = "C:\\MOS\\GOC HO TRO\\ngay_thi_W";
      string str2 = "C:\\MOS\\GOC HO TRO\\ngay_thi_E";
      string str3 = "C:\\MOS\\GOC HO TRO\\ngay_thi_P";
      string path = Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\nhac hoc.lnk";
      if (File.Exists(str1) || File.Exists(str2) || File.Exists(str3))
      {
        if (DateTime.Now >= this.TimNgayThiCuoi(str1, str2, str3))
        {
          File.Delete(path);
        }
        else
        {
          this.textBox1.Text = this.thongBao;
          this.textBox2.Text = this.nhacnho;
        }
      }
      else
      {
        DateTime dateTime = this.TimNgayHocGanNhat();
        if (dateTime > DateTime.Now)
        {
          this.textBox1.Text = "Học MOS hôm nay đi bạn";
          this.textBox2.Text = "Khi nào học ok thì vào gg meet đk thi và cài phần Mềm TEST\r\ngg Meet và các hỗ trợ khác năm trong shortcut HUONG DAN HOC trên desktop bạn";
        }
        else
        {
          int days = (DateTime.Now - dateTime).Days;
          if (days > 3 && days < 30)
          {
            this.textBox1.Text = "ĐÃ " + (object) days + " BẠN CHƯA HỌC MOS, TRANH THỦ HỌC NÓ ĐI BẠN";
            this.textBox2.Text = "Khi nào học ok thì vào gg meet đk thi và cài phần Mềm TEST\r\ngg Meet và các hỗ trợ khác năm trong shortcut HUONG DAN HOC trên desktop bạn";
          }
          else
            this.Close();
        }
      }
      this.timer1.Start();
    }

    private DateTime TimNgayThiCuoi(string w, string e, string p)
    {
      DateTime dateTime1 = DateTime.Now.AddDays(-2.0);
      if (File.Exists(w))
      {
        DateTime dateTime2 = this.check(w);
        this.thongBao = this.thongBao + " thi môn W: Ngày " + (object) (dateTime2.Day - 1) + " Tháng " + (object) dateTime2.Month + " năm " + (object) dateTime2.Year + "\r\n";
        if (dateTime2 > dateTime1)
          dateTime1 = dateTime2;
      }
      if (File.Exists(e))
      {
        DateTime dateTime3 = this.check(e);
        this.thongBao = this.thongBao + " thi môn E: Ngày " + (object) (dateTime3.Day - 1) + " Tháng " + (object) dateTime3.Month + " năm " + (object) dateTime3.Year + "\r\n";
        if (dateTime3 > dateTime1)
          dateTime1 = dateTime3;
      }
      if (File.Exists(p))
      {
        DateTime dateTime4 = this.check(p);
        this.thongBao = this.thongBao + " thi môn P: Ngày " + (object) (dateTime4.Day - 1) + " Tháng " + (object) dateTime4.Month + " năm " + (object) dateTime4.Year + "\r\n";
        if (dateTime4 > dateTime1)
          dateTime1 = dateTime4;
      }
      return dateTime1;
    }

    private DateTime check(string KeyPath)
    {
      DateTime dateTime = DateTime.Now.AddDays(-2.0);
      if (File.Exists(KeyPath))
      {
        try
        {
          TextReader textReader = (TextReader) new StreamReader(KeyPath);
          string[] strArray = textReader.ReadLine().Split('|');
          dateTime = new DateTime(int.Parse(strArray[2]), int.Parse(strArray[1]), int.Parse(strArray[0]));
          textReader.Close();
        }
        catch (Exception ex)
        {
          this.Close();
        }
      }
      return dateTime;
    }

    private DateTime TimNgayHocGanNhat()
    {
      DateTime dateTime = DateTime.Now.AddDays(1.0);
      string path = "C:\\MOS\\GOC HO TRO\\ngayhocgannhat";
      if (File.Exists(path))
      {
        try
        {
          TextReader textReader = (TextReader) new StreamReader(path);
          string str = textReader.ReadLine();
          textReader.Close();
          string[] strArray = str.Split('|');
          dateTime = new DateTime(int.Parse(strArray[2]), int.Parse(strArray[1]), int.Parse(strArray[0]));
        }
        catch (Exception ex)
        {
          dateTime = DateTime.Now.AddDays(1.0);
        }
      }
      return dateTime;
    }

    private void timer1_Tick(object sender, EventArgs e) => this.Close();
  }
}
