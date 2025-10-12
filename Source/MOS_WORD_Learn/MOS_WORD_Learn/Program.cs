// Decompiled with JetBrains decompiler
// Type: MOS_WORD_LEARN.Program
// Assembly: MOS_WORD_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using System;
using System.Net.NetworkInformation;
using System.Windows.Forms;

namespace MOS_WORD_LEARN
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

        private static void run()
        {
            // TODO: test form, should be remove
            //Application.Run((Form)new Form1_Dup());
            //return;
            Application.Run((Form)new Form2());
            switch (Program.status)
            {
                case 1:
                    Application.Run((Form)new Form1_Dup());
                    break;
                case 2:
                    int num = (int)MessageBox.Show("Phần Mềm Đã Hết Hạn");
                    break;
            }
        }

        public static bool GetMacAddress(string mac)
        {
            foreach (NetworkInterface networkInterface in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (networkInterface.GetPhysicalAddress().ToString() == mac)
                    return true;
            }
            return false;
        }

        public static bool GetMacAddress(string[] mac)
        {
            foreach (string mac1 in mac)
            {
                if (Program.GetMacAddress(mac1))
                    return true;
            }
            return false;
        }
    }
}
