using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Security.AccessControl;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.textBox1.Text = this.GetMacAddress().ToString();
            if (!(this.textBox1.Text == ""))
                return;
            this.textBox1.Text = this.getMac();
        }

        public string getMac() => ((IEnumerable<NetworkInterface>)NetworkInterface.GetAllNetworkInterfaces()).Where<NetworkInterface>((Func<NetworkInterface, bool>)(nic => nic.OperationalStatus == OperationalStatus.Up)).Select<NetworkInterface, string>((Func<NetworkInterface, string>)(nic => nic.GetPhysicalAddress().ToString())).FirstOrDefault<string>().ToString();

        public string GetMacAddress()
        {
            string macAddress = "";
            foreach (NetworkInterface networkInterface in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (networkInterface.GetPhysicalAddress().ToString() != "")
                    macAddress = macAddress + networkInterface.GetPhysicalAddress().ToString() + "|";
            }
            return macAddress;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                this.textBox1.Text = this.GetMacAddress().ToString();
                if (!(this.textBox1.Text == ""))
                    return;
                this.textBox1.Text = this.getMac();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
        }

        private void GANQUYEN()
        {
            string path = Path.Combine(Environment.ExpandEnvironmentVariables("%ProgramFiles(x86)%"), "MOS");
            if (!Directory.Exists(path))
                return;
            FileSystemRights fileSystemRights = FileSystemRights.FullControl;
            FileSystemAccessRule rule1 = new FileSystemAccessRule("Users", fileSystemRights, InheritanceFlags.None, PropagationFlags.NoPropagateInherit, AccessControlType.Allow);
            DirectoryInfo directoryInfo = new DirectoryInfo(path);
            DirectorySecurity accessControl = directoryInfo.GetAccessControl(AccessControlSections.Access);
            bool modified1 = false;
            accessControl.ModifyAccessRule(AccessControlModification.Set, (AccessRule)rule1, out modified1);
            InheritanceFlags inheritanceFlags = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;
            FileSystemAccessRule rule2 = new FileSystemAccessRule("Users", fileSystemRights, inheritanceFlags, PropagationFlags.InheritOnly, AccessControlType.Allow);
            bool modified2 = false;
            accessControl.ModifyAccessRule(AccessControlModification.Add, (AccessRule)rule2, out modified2);
            directoryInfo.SetAccessControl(accessControl);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                this.GANQUYEN();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("click chuột phải Run As Administrator");
                this.Close();
            }
        }
    }
}
