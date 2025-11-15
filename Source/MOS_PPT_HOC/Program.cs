// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Program
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using System;
using System.Net.NetworkInformation;
using System.Windows.Forms;


namespace MOS_PPT_LEARN
{
    internal static class Program
    {
        public static string user;
        public static string pass;
        public static string key;
        public static DateTime dt;
        public static int status = 0;

        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Program.run();
        }

        public static bool GetMacAddress(string mac)
        {
            foreach (NetworkInterface networkInterface in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (networkInterface.GetPhysicalAddress().ToString().Contains(mac))
                    return true;
            }
            return false;
        }

        private static void run()
        {
            //Application.Run((Form)new Form1());
            //return;
            Application.Run((Form)new Form2());
            switch (Program.status)
            {
                case 1:
                    Application.Run((Form)new Form1());
                    break;
                case 2:
                    int num = (int)MessageBox.Show("Phần Mềm Đã Hết Hạn");
                    break;
            }
        }
    }
}