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
            12,
            9,
            6,
            7,
            13,
            23,
            11,
            19,
            3,
            8,
            11
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
