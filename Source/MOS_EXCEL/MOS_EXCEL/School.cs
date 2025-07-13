// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.School
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

namespace MOS_EXCEL_LEARN
{
  internal class School
  {
    public static int[] Mn = new int[10]
    {
      15,
      6,
      6,
      21,
      3,
      6,
      9,
      10,
      3,
      9
    };

    public static int Tong()
    {
      int num = 0;
      for (int index = 0; index < School.Mn.Length; ++index)
        num += School.Mn[index];
      return num;
    }
  }
}
