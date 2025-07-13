// Decompiled with JetBrains decompiler
// Type: NHACHOC.School
// Assembly: NHACHOC, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 203B0FF7-2323-40ED-93E4-470E5FE69D30
// Assembly location: C:\Users\Admin\Downloads\WE_19\GOC HO TRO\NHAC HOC\NHACHOC.exe

namespace NHACHOC
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
