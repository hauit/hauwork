// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.Program
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MOS_EXCEL_TEST
{
  internal static class Program
  {
    public static string user;
    public static string pass;
    public static string key;
    public static int[] Lessons;
    public static int Tong;
    public static int TypeOfTest = 0;
    public static DateTime dt;
    public static int status = 0;

    [STAThread]
    private static void Main()
    {
      Application.EnableVisualStyles();
      Application.SetCompatibleTextRenderingDefault(false);
      Program.run();
    }

    private static void start()
    {
      Application.Run((Form) new Form3());
      if (!(Program.key == "Next"))
        return;
      Application.Run((Form) new Form4());
      if (Program.key == "Next")
      {
        Application.Run((Form) new Form5());
        if (Program.key == "Next")
        {
          Application.Run((Form) new Form6());
          if (Program.key == "Next")
            Program.menu();
        }
      }
    }

    private static void menu()
    {
      bool flag = true;
      string str1 = Path.Combine(Application.StartupPath, "Zip\\a");
      if (System.IO.File.Exists("d:\\config\\a"))
        str1 = "d:\\config\\a";
      string str2 = Path.Combine(Application.StartupPath, "Zip\\tam.txt");
      if (System.IO.File.Exists(str1))
      {
        DateTime now = DateTime.Now;
        Home.DecryptFile(str1, str2);
        TextReader textReader = (TextReader) new StreamReader(str2);
        string[] strArray = textReader.ReadLine().Split('9');
        DateTime dateTime = new DateTime(int.Parse(strArray[2]), int.Parse(strArray[1]), int.Parse(strArray[3]));
        if (now > dateTime)
          flag = false;
        textReader.Close();
      }
      else
        flag = false;
      System.IO.File.Delete(str2);
      string path = Path.Combine(Application.StartupPath, "Zip\\b");
      if (System.IO.File.Exists(path))
      {
        TextReader textReader = (TextReader) new StreamReader(path);
        int num1 = int.Parse(textReader.ReadLine());
        textReader.Close();
        if (num1 <= 3)
        {
          int num2 = num1 + 1;
          TextWriter textWriter = (TextWriter) new StreamWriter(Path.Combine(Application.StartupPath, "Zip\\b"), false);
          textWriter.WriteLine(num2);
          textWriter.Close();
        }
        else
          flag = false;
      }
      else
        flag = false;
      if (!flag)
        return;
      Program.run();
    }

    private static void run()
    {
      Application.Run((Form) new Form2());
      switch (Program.status)
      {
        case 1:
          Application.Run((Form) new Form1());
          break;
        case 2:
          int num = (int) MessageBox.Show("Phần Mềm Đã Hết Hạn");
          break;
      }
    }

    public static bool check()
    {
      string s = "";
      if (Program.user.Length != 6 || Program.user != "271565")
        return false;
      DateTime dateTime = new DateTime(1900, 1, 1);
      double num1 = ((double) (int) (Program.dt - dateTime).TotalDays + 2.0) * 271565.0;
      try
      {
        for (int startIndex = 0; startIndex < Program.user.Length; ++startIndex)
        {
          int num2 = ((startIndex + 1) * (Program.dt.Month + Program.dt.Day + Program.dt.Year) - int.Parse(Program.user.Substring(startIndex, 1)) - (startIndex + 1)) % 10;
          string str = s + num2.ToString();
          num2 = ((Program.dt.Day + Program.dt.Month + Program.dt.Year) * int.Parse(Program.user.Substring(startIndex, 1)) + (startIndex + 1) + int.Parse(Program.user.Substring(startIndex, 1))) % 10;
          s = str + num2.ToString();
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

    public static DateTime GetNetworkTime()
    {
      byte[] buffer = new byte[48];
      buffer[0] = (byte) 27;
      IPEndPoint remoteEP = new IPEndPoint(Dns.GetHostEntry("mail.tdt.edu.vn").AddressList[0], 123);
      Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
      socket.Connect((EndPoint) remoteEP);
      socket.Send(buffer);
      socket.Receive(buffer);
      socket.Close();
      return new DateTime(1900, 1, 1).AddMilliseconds((double) (long) (((ulong) ((long) buffer[40] << 24 | (long) buffer[41] << 16 | (long) buffer[42] << 8) | (ulong) buffer[43]) * 1000UL + ((ulong) ((long) buffer[44] << 24 | (long) buffer[45] << 16 | (long) buffer[46] << 8) | (ulong) buffer[47]) * 1000UL / 4294967296UL));
    }

    public static DateTime GetNetworkTime1()
    {
      byte[] buffer = new byte[48];
      buffer[0] = (byte) 27;
      IPEndPoint remoteEP = new IPEndPoint(Dns.GetHostEntry("tdt.edu.vn").AddressList[0], 123);
      Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
      socket.Connect((EndPoint) remoteEP);
      socket.ReceiveTimeout = 3000;
      socket.Send(buffer);
      socket.Receive(buffer);
      socket.Close();
      ulong uint32_1 = (ulong) BitConverter.ToUInt32(buffer, 40);
      ulong uint32_2 = (ulong) BitConverter.ToUInt32(buffer, 44);
      return new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMilliseconds((double) (long) ((ulong) Program.SwapEndianness(uint32_1) * 1000UL + (ulong) Program.SwapEndianness(uint32_2) * 1000UL / 4294967296UL)).ToLocalTime();
    }

    private static uint SwapEndianness(ulong x) => (uint) ((ulong) ((((long) x & (long) byte.MaxValue) << 24) + (((long) x & 65280L) << 8)) + ((x & 16711680UL) >> 8) + ((x & 4278190080UL) >> 24));

    public static DateTime GetNistTime()
    {
      DateTime nistTime = DateTime.MinValue;
      HttpWebRequest httpWebRequest = (HttpWebRequest) WebRequest.Create("http://nist.time.gov/actualtime.cgi?lzbc=siqm9b");
      httpWebRequest.Method = "GET";
      httpWebRequest.Accept = "text/html, application/xhtml+xml, */*";
      httpWebRequest.UserAgent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)";
      httpWebRequest.ContentType = "application/x-www-form-urlencoded";
      httpWebRequest.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
      HttpWebResponse response = (HttpWebResponse) httpWebRequest.GetResponse();
      if (response.StatusCode == HttpStatusCode.OK)
        nistTime = new DateTime(1970, 1, 1).AddMilliseconds((double) Convert.ToInt64(Regex.Match(new StreamReader(response.GetResponseStream()).ReadToEnd(), "(?<=\\btime=\")[^\"]*").Value) / 1000.0).ToLocalTime();
      return nistTime;
    }

    public static bool CheckForInternetConnection()
    {
      try
      {
        using (WebClient webClient = new WebClient())
        {
          using (webClient.OpenRead("http://www.google.com"))
            return true;
        }
      }
      catch
      {
        return false;
      }
    }

    public static bool checkExitFile()
    {
      bool flag = true;
      string requestUriString = "https://docs.google.com/document/d/1r2KpTRPvfU9k08YMH2lyhKupnbGNNAw7vQRE22p8OW0/edit?usp=sharing";
      HttpWebResponse httpWebResponse = (HttpWebResponse) null;
      HttpWebRequest httpWebRequest = (HttpWebRequest) WebRequest.Create(requestUriString);
      httpWebRequest.Method = "HEAD";
      try
      {
        httpWebResponse = (HttpWebResponse) httpWebRequest.GetResponse();
      }
      catch (WebException ex)
      {
        flag = false;
      }
      finally
      {
        httpWebResponse?.Close();
      }
      return flag;
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

    public static DateTime GetMicrosoftNistTime() => DateTime.ParseExact(WebRequest.Create("http://www.microsoft.com").GetResponse().Headers["date"], "ddd, dd MMM yyyy HH:mm:ss 'GMT'", (IFormatProvider) CultureInfo.InvariantCulture.DateTimeFormat, DateTimeStyles.AssumeUniversal);

    public static bool GetMacAddress(string mac)
    {
      foreach (NetworkInterface networkInterface in NetworkInterface.GetAllNetworkInterfaces())
      {
        if (networkInterface.GetPhysicalAddress().ToString().Contains(mac))
          return true;
      }
      return false;
    }
  }
}
