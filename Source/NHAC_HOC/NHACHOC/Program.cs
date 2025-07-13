// Decompiled with JetBrains decompiler
// Type: NHACHOC.Program
// Assembly: NHACHOC, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 203B0FF7-2323-40ED-93E4-470E5FE69D30
// Assembly location: C:\Users\Admin\Downloads\WE_19\GOC HO TRO\NHAC HOC\NHACHOC.exe

using System;
using System.Windows.Forms;

namespace NHACHOC
{
  internal static class Program
  {
    [STAThread]
    private static void Main()
    {
      Application.EnableVisualStyles();
      Application.SetCompatibleTextRenderingDefault(false);
      Application.Run((Form) new Form1());
    }
  }
}
