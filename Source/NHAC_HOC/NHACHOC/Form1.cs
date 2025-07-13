using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NHACHOC
{
    public partial class Form1 : Form
    {
        private string thongBao = "KIỂM TRA TRÊN WEB XEM LỊCH THI CỦA BẠN CÓ ĐÚNG VỚI LỊCH NÀY KHÔNG?:\r\n";
        private string nhacnho = "Thường Xuyên Ôn bài để kịp thi nha các bạn cố gắn mỗi môn 3 lần trên 950, IIG sẽ không chấp nhận lý do học không kịp nha các bạn\r\n";

        public Form1()
        {
            InitializeComponent();
        }

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
                        this.textBox1.Text = "ĐÃ " + (object)days + " BẠN CHƯA HỌC MOS, TRANH THỦ HỌC NÓ ĐI BẠN";
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
                this.thongBao = this.thongBao + " thi môn W: Ngày " + (object)(dateTime2.Day - 1) + " Tháng " + (object)dateTime2.Month + " năm " + (object)dateTime2.Year + "\r\n";
                if (dateTime2 > dateTime1)
                    dateTime1 = dateTime2;
            }
            if (File.Exists(e))
            {
                DateTime dateTime3 = this.check(e);
                this.thongBao = this.thongBao + " thi môn E: Ngày " + (object)(dateTime3.Day - 1) + " Tháng " + (object)dateTime3.Month + " năm " + (object)dateTime3.Year + "\r\n";
                if (dateTime3 > dateTime1)
                    dateTime1 = dateTime3;
            }
            if (File.Exists(p))
            {
                DateTime dateTime4 = this.check(p);
                this.thongBao = this.thongBao + " thi môn P: Ngày " + (object)(dateTime4.Day - 1) + " Tháng " + (object)dateTime4.Month + " năm " + (object)dateTime4.Year + "\r\n";
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
                    TextReader textReader = (TextReader)new StreamReader(KeyPath);
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
                    TextReader textReader = (TextReader)new StreamReader(path);
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
