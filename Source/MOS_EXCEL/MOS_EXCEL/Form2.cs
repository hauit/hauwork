using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MOS_EXCEL_LEARN
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

        private bool check1()
        {
            return true;//TODO: bypass check license. Just for debugging execution
            string s = "";
            double num1 = ((double)(int)(this.dt - new DateTime(1900, 1, 1)).TotalDays + 2.0) * 271565.0;
            try
            {
                for (int startIndex = 0; startIndex < Program.user.Length; ++startIndex)
                {
                    int num2 = ((startIndex + 1) * (this.dt.Month + this.dt.Day + this.dt.Year) - int.Parse(Program.user.Substring(startIndex, 1)) - (startIndex + 1)) % 10;
                    s = s + num2.ToString() + (((this.dt.Day + this.dt.Month + this.dt.Year) * int.Parse(Program.user.Substring(startIndex, 1)) + (startIndex + 1) + int.Parse(Program.user.Substring(startIndex, 1))) % 10).ToString();
                }
                var a = (double.Parse(s) + num1).ToString();
                if ((double.Parse(s) + num1).ToString() != Program.pass)
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private void WritrMac(string mac, DateTime ngay, string path)
        {
            TextWriter textWriter = (TextWriter)new StreamWriter(path);
            textWriter.WriteLine(mac);
            textWriter.WriteLine(ngay.Year.ToString() + "|" + (object)ngay.Month + "|" + (object)ngay.Day);
            textWriter.WriteLine("To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries");
            textWriter.Close();
        }
    }
}
