// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.School
// Assembly: MOS_EXCEL_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

namespace MOS_EXCEL_LEARN
{
  internal class School
  {
    public static int[] Mn = new int[10]
    {
      18,
      13,
      5,
      26,
      4,
      4,
      12,
      14,
      5,
      10
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
