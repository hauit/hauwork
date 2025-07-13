// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.School
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

namespace WindowsFormsApplication1
{
  internal class School
  {
    public static int[] Mn = new int[11]
    {
      9,
      8,
      3,
      9,
      5,
      4,
      5,
      7,
      4,
      5,
      7
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
