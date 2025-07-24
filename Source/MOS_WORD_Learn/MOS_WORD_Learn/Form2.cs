using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApplication1.Base;

namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        private DateTime dt;
        private DateTime ngayhethang;
        private string mac;
        public Form2()
        {
            InitializeComponent();
            this.textBoxUser.LostFocus += new EventHandler(this.textBoxUser_LostFocus);
        }

        private void textBoxUser_LostFocus(object sender, EventArgs e)
        {
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.textBoxUser.Text.ToUpper() != "ON LUYEN MOS")
                {
                    int num = (int)MessageBox.Show("sai user name");
                    this.textBoxUser.Focus();
                    return;
                }

                if (this.textBoxPass.Text == "")
                {
                    int num = (int)MessageBox.Show("Nhập Pass");
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
                ClsSession.Language = radioTV.Checked ? Language.Vietnamese : Language.English;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void WritrMac(string mac, DateTime ngay, string path)
        {
            TextWriter textWriter = (TextWriter)new StreamWriter(path);
            textWriter.WriteLine(mac);
            textWriter.WriteLine(ngay.Year.ToString() + "|" + (object)ngay.Month + "|" + (object)ngay.Day);
            textWriter.WriteLine("To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries");
            textWriter.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            try
            {
                this.check();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("không kết nói được với server vào google meet để tìm hổ trợ|" + ex.Message);
                this.Close();
            }
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

        public string getMac() => ((IEnumerable<NetworkInterface>)NetworkInterface.GetAllNetworkInterfaces()).Where<NetworkInterface>((Func<NetworkInterface, bool>)(nic => nic.OperationalStatus == OperationalStatus.Up)).Select<NetworkInterface, string>((Func<NetworkInterface, string>)(nic => nic.GetPhysicalAddress().ToString())).FirstOrDefault<string>().ToString();

        private void check()
        {
            ////TODO:
            //return;
            this.dt = Form2.GetDateFromeInternet();
            if (this.dt < new DateTime(2022, 1, 20))
            {
                int num = (int)MessageBox.Show("Chưa kết nối mạng");
                this.Close();
                return;
            }
            this.mac = this.getMac();
            if (this.mac.Length < 10)
            {
                int num = (int)MessageBox.Show("Tạo randomID bị lỗi");
                this.Close();
                return;
            }
            this.ngayhethang = this.dt.AddDays(30.0);
            string randomID = Base64Encode(this.mac + this.dt.ToString("yyyyMMdd"));
            this.richTextBox1.Text = $"Nếu bạn chưa có MK đăng nhập thì vui lòng gửi ID này \"{randomID}\" cho admin để lấy mật khẩu";
            //var a = Properties.Settings.Default.PASS;
            //var b = Properties.Settings.Default.DATE;
            //var c = Convert.ToInt32(b, 16);
            //this.ngayhethang = this.dt.AddDays(30.0);
            //Properties.Settings.Default.PASS = ToMD5(this.mac + this.ngayhethang.ToString("yyyyMMdd"));
            //Properties.Settings.Default.DATE = int.Parse(this.ngayhethang.ToString("yyyMMdd")).ToString("X");
            //Properties.Settings.Default.Save();
            //this.dt = this.randomday(this.dt);
            //this.textBoxPassSo.Text = this.dt.ToShortDateString();
            //this.WritrMac(this.mac, this.dt, Path.Combine(Application.StartupPath, "zip\\a"));
            //Home.EncryptFileT(Path.Combine(Application.StartupPath, "zip\\a"), Path.Combine(Application.StartupPath, "zip\\c"));
        }

        private string Base64Encode(string data)
        {
            byte[] textBytes = Encoding.UTF8.GetBytes(data);
            var result = Convert.ToBase64String(textBytes);

            return result;
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

        private DateTime randomday(DateTime dt)
        {
            DateTime dateTime;
            if (System.IO.File.Exists(Path.Combine(Application.StartupPath, "zip\\c")))
            {
                Home.DecryptFileT(Path.Combine(Application.StartupPath, "zip\\c"), Path.Combine(Application.StartupPath, "zip\\a"));
                TextReader textReader = (TextReader)new StreamReader(Path.Combine(Application.StartupPath, "zip\\a"));
                if (this.mac == textReader.ReadLine())
                {
                    string[] strArray = textReader.ReadLine().Split('|');
                    dateTime = new DateTime(int.Parse(strArray[0]), int.Parse(strArray[1]), int.Parse(strArray[2]));
                    if (dateTime.AddDays(120.0) < dt)
                    {
                        Random random = new Random();
                        dateTime = dt.AddDays((double)random.Next(-30, 30));
                    }
                }
                else
                {
                    Random random = new Random();
                    dateTime = dt.AddDays((double)random.Next(-30, 30));
                }
                textReader.Close();
            }
            else
            {
                Random random = new Random();
                dateTime = dt.AddDays((double)random.Next(-30, 30));
            }
            return dateTime;
        }

        private bool check1()
        {
            ////TODO:
            //return true;
            string expire = GetExpireDate();
            if (!string.IsNullOrEmpty(expire))
            {
                DateTime.TryParseExact(expire, "yyyyMMdd", null, DateTimeStyles.None, out this.ngayhethang);
            }
            if (this.ngayhethang < this.dt)
            {
                MessageBox.Show("Phần mềm hết hạn");
                return false;
            }

            string pass = ToMD5(this.mac + this.ngayhethang.ToString("yyyyMMdd"));
            if (pass.ToLower() == textBoxPass.Text.ToLower())
            {
                Properties.Settings.Default.PASS = pass.ToLower();
                Properties.Settings.Default.DATE = int.Parse(this.ngayhethang.ToString("yyyMMdd")).ToString("X");
                Properties.Settings.Default.Save();
            }
            else
            {
                Properties.Settings.Default.DATE = string.Empty;
                Properties.Settings.Default.Save();
                MessageBox.Show("Phần Sai mật khẩu. vui lòng liên hệ Admin để lấy lại");
                return false;
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

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
