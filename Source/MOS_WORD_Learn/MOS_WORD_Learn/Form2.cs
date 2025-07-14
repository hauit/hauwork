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
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            if (this.textBoxUser.Text.ToUpper() != "ON LUYEN MOS")
            {
                int num = (int)MessageBox.Show("sai user name");
                this.textBoxUser.Focus();
            }
            else if (this.textBoxPass.Text == "")
            {
                int num = (int)MessageBox.Show("Nhập Pass");
                this.textBoxPass.Focus();
            }
            else
            {
                Program.user = "271565";
                Program.pass = this.textBoxPass.Text;
                if (this.check1())
                {
                    Program.status = 1;
                    this.WritrMac(this.mac, this.ngayhethang, Path.Combine(Application.StartupPath, "zip\\a"));
                    Home.EncryptFileT(Path.Combine(Application.StartupPath, "zip\\a"), Path.Combine(Application.StartupPath, "zip\\b"));
                    System.IO.File.Delete(Path.Combine(Application.StartupPath, "zip\\c"));
                    string str = "C:\\MOS\\GOC HO TRO\\NHAC HOC.lnk";
                    string destFileName = Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\nhac hoc.lnk";
                    if (System.IO.File.Exists(str))
                        System.IO.File.Copy(str, destFileName, true);
                    this.Close();
                }
                else
                {
                    int num = (int)MessageBox.Show("Sai Pass, liên hệ google meet để lấy pass đúng");
                }
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

        private void Form2_Load(object sender, EventArgs e) => this.check();

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
            return;
            try
            {
                this.dt = Form2.GetDateFromeInternet();
                if (this.dt < new DateTime(2022, 1, 20))
                {
                    int num = (int)MessageBox.Show("Chưa kết nối mạng");
                    this.Close();
                }
                else if (System.IO.File.Exists(Path.Combine(Application.StartupPath, "zip\\b")))
                {
                    try
                    {
                        Home.DecryptFileT(Path.Combine(Application.StartupPath, "zip\\b"), Path.Combine(Application.StartupPath, "zip\\a"));
                        TextReader textReader = (TextReader)new StreamReader(Path.Combine(Application.StartupPath, "zip\\a"));
                        string mac = textReader.ReadLine();
                        string[] strArray = textReader.ReadLine().Split('|');
                        DateTime dateTime = new DateTime(int.Parse(strArray[0]), int.Parse(strArray[1]), int.Parse(strArray[2]));
                        textReader.Close();
                        this.mac = this.getMac();
                        if (this.GetMacAddress(mac))
                        {
                            if (this.dt < dateTime)
                            {
                                Program.status = 1;
                                this.Close();
                            }
                            else
                            {
                                int num = (int)MessageBox.Show("Phần mềm hết hạng");
                                System.IO.File.Delete(Path.Combine(Application.StartupPath, "zip\\b"));
                                System.IO.File.Delete(Path.Combine(Application.StartupPath, "zip\\c"));
                                this.Close();
                            }
                        }
                        else
                        {
                            int num = (int)MessageBox.Show("sai key");
                            this.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        int num = (int)MessageBox.Show("Sai key" + ex.Message);
                        this.Close();
                    }
                }
                else
                {
                    this.mac = this.getMac();
                    if (this.mac.Length < 10)
                    {
                        int num = (int)MessageBox.Show("Địa chỉ MAC có vấn đề");
                        this.Close();
                    }
                    else
                    {
                        this.ngayhethang = this.dt.AddDays(30.0);
                        this.dt = this.randomday(this.dt);
                        this.textBoxPassSo.Text = this.dt.ToShortDateString();
                        this.WritrMac(this.mac, this.dt, Path.Combine(Application.StartupPath, "zip\\a"));
                        Home.EncryptFileT(Path.Combine(Application.StartupPath, "zip\\a"), Path.Combine(Application.StartupPath, "zip\\c"));
                    }
                }
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("không kết nói được với server vào google meet để tìm hổ trợ|" + ex.Message);
                this.Close();
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
            return true;
            string s = "";
            double num1 = ((double)(int)(this.dt - new DateTime(1900, 1, 1)).TotalDays + 2.0) * 271565.0;
            try
            {
                for (int startIndex = 0; startIndex < Program.user.Length; ++startIndex)
                {
                    int num2 = ((startIndex + 1) * (this.dt.Month + this.dt.Day + this.dt.Year) - int.Parse(Program.user.Substring(startIndex, 1)) - (startIndex + 1)) % 10;
                    s = s + num2.ToString() + (((this.dt.Day + this.dt.Month + this.dt.Year) * int.Parse(Program.user.Substring(startIndex, 1)) + (startIndex + 1) + int.Parse(Program.user.Substring(startIndex, 1))) % 10).ToString();
                }
                if ((double.Parse(s) + num1).ToString() != Program.pass)
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
