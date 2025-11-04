using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Cache;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApplication1.Base;

namespace MOS_EXCEL_LEARN
{
    public partial class Form2 : Form
    {
        private DateTime dt;
        private DateTime ngayhethang;
        private DateTime ngayhethangCurrentDate;
        private string mac;
        public Form2()
        {
            InitializeComponent();
            this.textBoxUser.LostFocus += new EventHandler(this.textBoxUser_LostFocus);
            if (!string.IsNullOrEmpty(Properties.Settings.Default.PASS))
                this.textBoxPass.Text = Properties.Settings.Default.PASS;
        }

        private void textBoxUser_LostFocus(object sender, EventArgs e)
        {
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            //if (Process.GetProcessesByName("EXCEL").Length > 0)
            //{
            //    MessageBox.Show("Vui lòng đóng tất cả các file excel trước khi học để tránh mất dữ liệu");
            //    this.textBoxUser.Focus();
            //    return;
            //}

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

            if (this.textBoxUser.Text.ToUpper() != "MOS TRAINING")
            {
                int num = (int)MessageBox.Show("Sai tên đăng nhập");
                this.textBoxUser.Focus();
                return;
            }

            if (this.textBoxPass.Text == "")
            {
                int num = (int)MessageBox.Show("Nhập mật khẩu");
                this.textBoxPass.Focus();
                return;
            }

            Program.user = "271565";
            Program.pass = this.textBoxPass.Text;
            if (!this.check1())
            {
                return;
            }

            Program.status = 1;
            this.Close();
        }

        private bool check1()
        {
            string expire = GetExpireDate();
            if (!string.IsNullOrEmpty(expire))
            {
                DateTime.TryParseExact(expire, "yyyyMMdd", null, DateTimeStyles.None, out this.ngayhethang);
            }

            string pass = ToMD5(this.mac + this.ngayhethang.ToString("yyyyMMdd"));
            if (this.ngayhethang < this.dt)
            {
                MessageBox.Show("Phần mềm đã hết hạn");
                Properties.Settings.Default.PASS = string.Empty;
                Properties.Settings.Default.DATE = string.Empty;
                Properties.Settings.Default.Save();
                return false;
            }

            if (pass.ToLower() == textBoxPass.Text.ToLower())
            {
                Properties.Settings.Default.PASS = pass.ToLower();
                Properties.Settings.Default.DATE = int.Parse(this.ngayhethang.ToString("yyyMMdd")).ToString("X");
                Properties.Settings.Default.Save();
            }
            else
            {
                //
                this.ngayhethangCurrentDate = this.dt.AddDays(30.0);
                string passCurrentDate = ToMD5(this.mac + this.ngayhethangCurrentDate.ToString("yyyyMMdd"));
                if (passCurrentDate.ToLower() == textBoxPass.Text.ToLower())
                {
                    Properties.Settings.Default.PASS = passCurrentDate.ToLower();
                    Properties.Settings.Default.DATE = int.Parse(this.ngayhethangCurrentDate.ToString("yyyMMdd")).ToString("X");
                    Properties.Settings.Default.Save();
                }
                else
                {
                    Properties.Settings.Default.DATE = string.Empty;
                    Properties.Settings.Default.Save();
                    MessageBox.Show("Mật khẩu không đúng hoặc Phần mềm đã hết hạn. Vui lòng liên hệ Admin để cấp lại");
                    return false;
                }
            }
            return true;
        }

        private string GetExpireDate()
        {
            string data = Properties.Settings.Default.DATE;
            if (!string.IsNullOrEmpty(data))
            {
                data = Convert.ToInt32(data, 16).ToString();
            }

            return data;
        }

        public string ToMD5(string input)
        {
            input += "FJKSFxnEO7EUKIK9KFWT";
            using (MD5 md5 = MD5.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(input);
                byte[] hashBytes = md5.ComputeHash(inputBytes);
                StringBuilder sb = new StringBuilder();
                foreach (byte b in hashBytes)
                {
                    sb.Append(b.ToString("x2"));
                }

                string result = sb.ToString();
                return result.Substring(result.Length - 10);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            try
            {
                this.check();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Không kết nói được với server vào website để hỗ trợ: https://mos360.vn" + ex.Message);
                this.Close();
            }
        }

        private void check()
        {
            this.dt = Form2.GetDateFromeInternet();
            if (this.dt < new DateTime(2022, 1, 20))
            {
                int num = (int)MessageBox.Show("Chưa kết nối mạng");
                this.Close();
                return;
            }
            this.mac = this.GetMotherboardID();
            if (this.mac.Length < 2)
            {
                int num = (int)MessageBox.Show("Tạo ID bị lỗi. Vui lòng liên hệ Admin");
                this.Close();
                return;
            }
            this.ngayhethang = this.dt.AddDays(30.0);
            string randomID = Base64Encode(this.mac + this.dt.ToString("yyyyMMdd"));
            this.richTextBox1.Text = $"Nếu bạn chưa có mật khẩu để đăng nhập, vui lòng gửi ID bên dưới cho Admin để được cấp:\n\n{randomID}\n\nWebsite hỗ trợ: https://mos360.vn";
        }

        private string Base64Encode(string data)
        {
            byte[] textBytes = Encoding.UTF8.GetBytes(data);
            var result = Convert.ToBase64String(textBytes);

            return result;
        }

        private string GetMotherboardID()
        {
            string motherboardID = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard");
            foreach (ManagementObject obj in searcher.Get())
            {
                motherboardID = obj["SerialNumber"]?.ToString()?.Trim();
                break;
            }
            return motherboardID;
        }

        public static DateTime GetDateFromeInternet()
        {
            DateTime dateFromeInternet;
            try
            {
                dateFromeInternet = Form2.GetMicrosoftNistTime();
            }
            catch (Exception ex1)
            {
                try
                {
                    dateFromeInternet = Form2.GetFastestNISTDate();
                }
                catch (Exception ex2)
                {
                    dateFromeInternet = Form2.GetDateTime();
                }
            }
            return dateFromeInternet;
        }

        public static DateTime GetMicrosoftNistTime() => DateTime.ParseExact(WebRequest.Create("https://www.google.com/search?q=Iraq time").GetResponse().Headers["date"], "ddd, dd MMM yyyy HH:mm:ss 'GMT'", (IFormatProvider)CultureInfo.InvariantCulture.DateTimeFormat, DateTimeStyles.AssumeUniversal);

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
            foreach (string hostname in ((IEnumerable<string>)source).OrderBy<string, double>((Func<string, double>)(s => rnd.NextDouble())).Take<string>(5))
            {
                try
                {
                    string str = string.Empty;
                    using (StreamReader streamReader = new StreamReader((Stream)new TcpClient(hostname, 13).GetStream()))
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

        public static DateTime GetDateTime()
        {
            DateTime dateTime = DateTime.MinValue;
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("https://www.microsoft.com");
            httpWebRequest.Method = "GET";
            httpWebRequest.Accept = "text/html, application/xhtml+xml, */*";
            httpWebRequest.UserAgent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)";
            httpWebRequest.ContentType = "application/x-www-form-urlencoded";
            httpWebRequest.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
            HttpWebResponse response = (HttpWebResponse)httpWebRequest.GetResponse();
            if (response.StatusCode == HttpStatusCode.OK)
                dateTime = DateTime.ParseExact(response.Headers["date"], "ddd, dd MMM yyyy HH:mm:ss 'GMT'", (IFormatProvider)CultureInfo.InvariantCulture.DateTimeFormat, DateTimeStyles.AssumeUniversal);
            return dateTime;
        }
    }
}
