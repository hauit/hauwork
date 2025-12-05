using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Cache;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace MOS_PPT_TEST
{
    public partial class Form2 : Form
    {
        private DateTime dt;
        private DateTime ngayhethang;
        private DateTime ngayhethangCurrentDate;
        private string mac;
        private string user;

        public Form2()
        {
            Form1.listExam = GetQuestion(System.Windows.Forms.Application.StartupPath + "\\zip\\Exam\\ExamList");
            CheckNetFramework();
            InitializeComponent();
            for(int i = 0; i < Form1.listExam.Count; i++)
            {
                this.cbExamList.DataSource = Form1.listExam;
                this.cbExamList.DisplayMember = "ExamIndex";
                this.cbExamList.ValueMember = "ExamIndex";
            }
            this.textBoxUser.LostFocus += new EventHandler(this.textBoxUser_LostFocus);
            if (!string.IsNullOrEmpty(Properties.Settings.Default.USER))
                this.textBoxUser.Text = Properties.Settings.Default.USER;
            if (!string.IsNullOrEmpty(Properties.Settings.Default.PASS))
                this.textBoxPass.Text = Properties.Settings.Default.PASS;
        }

        private List<Exam> GetQuestion(string path)
        {
            byte[] buffer = Home.DecryptFile(path);
            string jsonString = Encoding.UTF8.GetString(buffer);

            // Giải mã JSON thành object
            var listExam = JsonConvert.DeserializeObject<List<Exam>>(jsonString);
            return listExam;
        }

        private void textBoxUser_LostFocus(object sender, EventArgs e)
        {
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            try
            {
                //if (Process.GetProcessesByName("WINWORD").Length > 0)
                //{
                //    MessageBox.Show("Vui lòng đóng tất cả các file word trước khi thi để tránh mất dữ liệu");
                //    this.textBoxUser.Focus();
                //    return;
                //}

                Process[] pptProcesses = Process.GetProcessesByName("POWERPNT");

                if (pptProcesses.Length > 0)
                {
                    foreach (var process in pptProcesses)
                    {
                        try
                        {
                            process.Kill(); // buộc dừng tiến trình
                            process.WaitForExit(); // đợi đến khi tiến trình thật sự đóng
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Không thể tắt PowerPoint: " + ex.Message);
                        }
                    }
                }

                if (this.textBoxUser.Text == "")
                {
                    int num = (int)MessageBox.Show("Nhập tài khoản");
                    this.textBoxUser.Focus();
                    return;
                }

                if (this.textBoxPass.Text == "")
                {
                    int num = (int)MessageBox.Show("Nhập mật khẩu");
                    this.textBoxPass.Focus();
                    return;
                }

                //if (this.comboBox1.Visible)
                //{
                //    if (this.comboBox1.Text == "")
                //    {
                //        int num = (int)MessageBox.Show("Bạn phải chọn đề thi");
                //    }
                //    else
                //    {
                //        Program.user = "271565";
                //        Program.pass = this.textBoxPass.Text;
                //        Program.status = 1;
                //        this.Close();
                //    }
                //}
                //else
                //{
                //    Program.user = "271565";
                //    Program.pass = this.textBoxPass.Text;
                //    Program.status = 1;
                //    this.Close();
                //}

                Program.user = this.textBoxUser.Text;
                Program.pass = this.textBoxPass.Text;
                if (!this.check1())
                {
                    return;
                }

                Program.status = 1;
                Form1.currentExam = (Exam)this.cbExamList.SelectedItem;
                Form1.Language = this.radioEN.Checked ? "EN" : "VI";
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) => Program.TypeOfTest = this.cbExamList.SelectedIndex;

        private void Form2_Load(object sender, EventArgs e)
        {
            //string str = Path.Combine(Application.StartupPath, "zip\\b");
            //try
            //{
            //    if (!System.IO.File.Exists(str))
            //    {
            //        int num1 = (int)MessageBox.Show("Chọn file key để active phần mềm này");
            //        if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            //        {
            //            if (System.IO.File.Exists(this.openFileDialog1.FileName))
            //            {
            //                System.IO.File.Copy(this.openFileDialog1.FileName, str);
            //                System.IO.File.Copy("C:\\MOS\\GOC HO TRO\\NHAC HOC.lnk", Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\nhac hoc.lnk", true);
            //                this.check(str);
            //            }
            //            else
            //            {
            //                int num2 = (int)MessageBox.Show("Không tiềm thấy file");
            //                this.Close();
            //            }
            //        }
            //        else
            //            this.Close();
            //    }
            //    else
            //        this.check(str);
            //}
            //catch (Exception ex)
            //{
            //    int num = (int)MessageBox.Show("Không thể copy file, click chuột phải chọn chạy bằng administrator |" + ex.Message);
            //    this.Close();
            //}

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
                    dateFromeInternet = Form2.GetDateTime();
                }
                catch (Exception ex2)
                {
                    dateFromeInternet = Form2.GetFastestNISTDate();
                }
            }
            return dateFromeInternet;
        }

        public static DateTime GetMicrosoftNistTime() => DateTime.ParseExact(WebRequest.Create("https://www.google.com/search?q=Iraq time").GetResponse().Headers["date"], "ddd, dd MMM yyyy HH:mm:ss 'GMT'", (IFormatProvider)CultureInfo.InvariantCulture.DateTimeFormat, DateTimeStyles.AssumeUniversal);

        public static DateTime GetDateTime()
        {
            DateTime dateTime = DateTime.MinValue;
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("https://www.microsoft.com/vi-vn");
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

        public bool CheckForInternetConnection(int timeoutMs = 10000, string url = null)
        {
            try
            {
                HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.KeepAlive = false;
                httpWebRequest.Timeout = timeoutMs;
                using ((HttpWebResponse)httpWebRequest.GetResponse())
                    return true;
            }
            catch
            {
                return false;
            }
        }

        private void checkOld(string KeyPath)
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
                    int num = (int)MessageBox.Show("Bạn phải chọn đúng file Key của nhà cung cấp |lỗi: " + ex.Message);
                    System.IO.File.Delete(KeyPath);
                    this.Close();
                }
                try
                {
                    TextReader textReader = (TextReader)new StreamReader(str1);
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
                        int num = (int)MessageBox.Show("Key của bạn sai Version, vào web thimos1000.wordpress.com để tìm hổ trợ");
                        System.IO.File.Delete(KeyPath);
                        this.Close();
                    }
                    try
                    {
                        DateTime dateFromeInternet = Form2.GetDateFromeInternet();
                        if (dateFromeInternet < new DateTime(2022, 5, 24))
                        {
                            if (this.CheckForInternetConnection(url: "https://www.google.com.vn/?hl=vi"))
                            {
                                DateTime now = DateTime.Now;
                                if (now > dateTime || (dateTime - now).Days > 70)
                                {
                                    int num = (int)MessageBox.Show("Phần mềm hết hạng");
                                    System.IO.File.Delete(KeyPath);
                                    this.Close();
                                }
                                else
                                {
                                    Program.TypeOfTest = !(now.AddDays(2.0) > dateTime) ? 1 : 0;
                                    TextWriter textWriter = (TextWriter)new StreamWriter("C:\\MOS\\GOC HO TRO\\ngay_thi_W");
                                    textWriter.WriteLine(strArray[3] + "|" + strArray[1] + "|" + strArray[2]);
                                    textWriter.Close();
                                }
                            }
                            else
                            {
                                int num = (int)MessageBox.Show("không có mạng hoặc sai giờ hệ thống hoặc bị chặn bởi tường lửa");
                                this.Close();
                            }
                        }
                        else if (dateFromeInternet > dateTime || (dateTime - dateFromeInternet).Days > 70)
                        {
                            int num = (int)MessageBox.Show("Phần mềm hết hạng");
                            System.IO.File.Delete(KeyPath);
                            this.Close();
                        }
                        else
                        {
                            Program.TypeOfTest = !(dateFromeInternet.AddDays(2.0) > dateTime) ? 1 : 0;
                            TextWriter textWriter = (TextWriter)new StreamWriter("C:\\MOS\\GOC HO TRO\\ngay_thi_W");
                            textWriter.WriteLine(strArray[3] + "|" + strArray[1] + "|" + strArray[2]);
                            textWriter.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        int num = (int)MessageBox.Show("không kết nói được với server vào web thimos1000.wordpress.com để tìm hổ trợ|" + ex.Message);
                        this.Close();
                    }
                    if (str2 != "W2019")
                    {
                        int num = (int)MessageBox.Show("Phần mềm W2019 mà key la " + str2);
                        System.IO.File.Delete(KeyPath);
                        this.Close();
                    }
                    try
                    {
                        if (this.GetMacAddress(mac))
                            return;
                        int num = (int)MessageBox.Show("Key này không dành cho máy này; vào web thimos1000.wordpress.com để tìm hổ trợ");
                        System.IO.File.Delete(KeyPath);
                        this.Close();
                    }
                    catch (Exception ex)
                    {
                        int num = (int)MessageBox.Show("không truy cập được ID máy |" + ex.Message);
                        this.Close();
                    }
                }
                catch (Exception ex)
                {
                    int num = (int)MessageBox.Show("Không phải file Key |" + ex.Message);
                    System.IO.File.Delete(KeyPath);
                    this.Close();
                }
            }
            else
            {
                int num1 = (int)MessageBox.Show("không tìm thây file");
            }
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
            this.mac = this.GetMotherboardID();
            if (this.mac.Length < 2)
            {
                int num = (int)MessageBox.Show("Tạo ID bị lỗi. Vui lòng liên hệ Admin");
                this.Close();
                return;
            }
            this.ngayhethang = this.dt.AddDays(60.0);
            string randomID = Base64Encode(this.mac + this.dt.ToString("yyyyMMdd"));
            this.richTextBox1.Text = $"Nếu bạn chưa có mật khẩu để đăng nhập, vui lòng gửi ID bên dưới cho Admin để được cấp:\n\n{randomID}\n\nWebsite hỗ trợ: https://mos360.vn";
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

        private bool check1()
        {
            ////TODO:
            //return true;
            string expire = GetExpireDate();
            if (!string.IsNullOrEmpty(expire))
            {
                DateTime.TryParseExact(expire, "yyyyMMdd", null, DateTimeStyles.None, out this.ngayhethang);
            }

            string user = this.textBoxUser.Text.Trim();
            string pass = ToMD5(user.ToLower() + this.mac + this.ngayhethang.ToString("yyyyMMdd"));

            if (this.ngayhethang < this.dt)
            {
                MessageBox.Show("Phần mềm đã hết hạn");
                Properties.Settings.Default.PASS = string.Empty;
                Properties.Settings.Default.DATE = string.Empty;
                Properties.Settings.Default.Save();
                return false;
            }
            if (pass.ToLower() == textBoxPass.Text.ToLower() && user.ToLower() == textBoxUser.Text.ToLower())
            {
                Properties.Settings.Default.PASS = pass.ToLower();
                Properties.Settings.Default.USER = user.ToLower();
                Properties.Settings.Default.DATE = int.Parse(this.ngayhethang.ToString("yyyMMdd")).ToString("X");
                Properties.Settings.Default.Save();
            }
            else
            {
                this.ngayhethangCurrentDate = this.dt.AddDays(60.0);
                string passCurrentDate = ToMD5(user.ToLower() + this.mac + this.ngayhethangCurrentDate.ToString("yyyyMMdd"));
                if (passCurrentDate.ToLower() == textBoxPass.Text.ToLower())
                {
                    Properties.Settings.Default.USER = user.ToLower();
                    Properties.Settings.Default.PASS = passCurrentDate.ToLower();
                    Properties.Settings.Default.DATE = int.Parse(this.ngayhethangCurrentDate.ToString("yyyMMdd")).ToString("X");
                    Properties.Settings.Default.Save();
                }
                else
                {
                    Properties.Settings.Default.DATE = string.Empty;
                    Properties.Settings.Default.Save();
                    MessageBox.Show("Tài khoản và Mật khẩu không đúng. Vui lòng liên hệ Admin để cấp lại");
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

        private void btnCancel_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private static void CheckNetFramework()
        {
            const int releaseKey471 = 461308; // .NET Framework 4.7.1
            int releaseKey = GetFrameworkReleaseKey();

            if (releaseKey < releaseKey471)
            {
                DialogResult result = MessageBox.Show(
                    "Ứng dụng cần .NET Framework 4.7.1.\nBạn có muốn tải và cài đặt ngay không?",
                    "Thiếu .NET Framework 4.7.1",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    //Process.Start(new ProcessStartInfo
                    //{
                    //    FileName = "https://go.microsoft.com/fwlink/?linkid=2088631",
                    //    UseShellExecute = true
                    //});
                    Process.Start("https://go.mos360.vn/net48");
                    //Environment.Exit(0);
                }
                else
                {
                    //MessageBox.Show("Bạn có thể tự tải và cài đặt .NET Framework 4.8 sau.", "Thông báo");
                    //close the application
                    //Environment.Exit(0);
                }

                Environment.Exit(0);
            }
        }

        private static int GetFrameworkReleaseKey()
        {
            try
            {
                using (RegistryKey ndpKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                    .OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"))
                {
                    if (ndpKey != null && ndpKey.GetValue("Release") != null)
                    {
                        return (int)ndpKey.GetValue("Release");
                        //return 1;
                    }
                }
            }
            catch { }
            return 0;
        }
    }
}
