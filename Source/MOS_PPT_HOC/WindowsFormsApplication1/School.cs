// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.School
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

#nullable disable
namespace WindowsFormsApplication1;

internal class School
{
  public static int[] Mn = new int[18]
  {
    5,
    14,
    3,
    6,
    2,
    9,
    4,
    6,
    4,
    1,
    15,
    5,
    5,
    7,
    5,
    5,
    4,
    4
  };

  public static int Tong()
  {
    int num = 0;
    for (int index = 0; index < School.Mn.Length; ++index)
      num += School.Mn[index];
    return num;
  }
}
