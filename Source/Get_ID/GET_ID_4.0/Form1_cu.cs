// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Form1
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: F5185CDD-2BDC-4442-BA81-491144CA4DAC
// Assembly location: C:\Users\Admin\Downloads\WE_19\GOC HO TRO\GET_ID_4.0.exe

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Security.AccessControl;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
  public class Form1_cu : Form
  {
    private IContainer components = (IContainer) null;
    private Button button1;
    private TextBox textBox1;
    private Button button2;

    protected override void Dispose(bool disposing)
    {
      if (disposing && this.components != null)
        this.components.Dispose();
      base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
      this.button1 = new Button();
      this.textBox1 = new TextBox();
      this.button2 = new Button();
      this.SuspendLayout();
      this.button1.Location = new Point(178, 311);
      this.button1.Name = "button1";
      this.button1.Size = new Size(75, 23);
      this.button1.TabIndex = 0;
      this.button1.Text = "&OK";
      this.button1.UseVisualStyleBackColor = true;
      this.button1.Click += new EventHandler(this.button1_Click);
      this.textBox1.Font = new Font("Microsoft Sans Serif", 24f, FontStyle.Bold, GraphicsUnit.Point, (byte) 163);
      this.textBox1.Location = new Point(64, 52);
      this.textBox1.Multiline = true;
      this.textBox1.Name = "textBox1";
      this.textBox1.ReadOnly = true;
      this.textBox1.ScrollBars = ScrollBars.Both;
      this.textBox1.Size = new Size(344, 214);
      this.textBox1.TabIndex = 1;
      this.textBox1.Text = "No your Computer";
      this.button2.Location = new Point(351, 330);
      this.button2.Name = "button2";
      this.button2.Size = new Size(109, 34);
      this.button2.TabIndex = 2;
      this.button2.Text = "Cho phep Chay";
      this.button2.UseVisualStyleBackColor = true;
      this.button2.Visible = false;
      this.button2.Click += new EventHandler(this.button2_Click);
      this.AutoScaleDimensions = new SizeF(6f, 13f);
      this.AutoScaleMode = AutoScaleMode.Font;
      this.ClientSize = new Size(472, 366);
      this.Controls.Add((Control) this.button2);
      this.Controls.Add((Control) this.textBox1);
      this.Controls.Add((Control) this.button1);
      this.Name = nameof (Form1_cu);
      this.Text = nameof (Form1_cu);
      this.Load += new EventHandler(this.Form1_Load);
      this.ResumeLayout(false);
      this.PerformLayout();
    }

    public Form1_cu() => this.InitializeComponent();

    private void button1_Click(object sender, EventArgs e)
    {
      this.textBox1.Text = this.GetMacAddress().ToString();
      if (!(this.textBox1.Text == ""))
        return;
      this.textBox1.Text = this.getMac();
    }

    public string getMac() => ((IEnumerable<NetworkInterface>) NetworkInterface.GetAllNetworkInterfaces()).Where<NetworkInterface>((Func<NetworkInterface, bool>) (nic => nic.OperationalStatus == OperationalStatus.Up)).Select<NetworkInterface, string>((Func<NetworkInterface, string>) (nic => nic.GetPhysicalAddress().ToString())).FirstOrDefault<string>().ToString();

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
        int num = (int) MessageBox.Show(ex.Message);
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
      accessControl.ModifyAccessRule(AccessControlModification.Set, (AccessRule) rule1, out modified1);
      InheritanceFlags inheritanceFlags = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;
      FileSystemAccessRule rule2 = new FileSystemAccessRule("Users", fileSystemRights, inheritanceFlags, PropagationFlags.InheritOnly, AccessControlType.Allow);
      bool modified2 = false;
      accessControl.ModifyAccessRule(AccessControlModification.Add, (AccessRule) rule2, out modified2);
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
        int num = (int) MessageBox.Show("click chuột phải Run As Administrator");
        this.Close();
      }
    }
  }
}
