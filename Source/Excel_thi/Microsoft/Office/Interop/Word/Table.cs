// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Table
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [CompilerGenerated]
  [DefaultMember("Range")]
  [TypeIdentifier]
  [Guid("00020951-0000-0000-C000-000000000046")]
  [ComImport]
  public interface Table
  {
    Range Range { [DispId(0)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap1_3();

    Columns Columns { [DispId(100)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    Rows Rows { [DispId(101)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    Borders Borders { [DispId(1100)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(1100)] [param: MarshalAs(UnmanagedType.Interface), In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_2();

    int AutoFormatType { [DispId(106)] get; }

    [SpecialName]
    sealed extern void _VtblGap3_8();

    [DispId(17)]
    [return: MarshalAs(UnmanagedType.Interface)]
    Microsoft.Office.Interop.Word.Cell Cell([In] int Row, [In] int Column);

    [SpecialName]
    sealed extern void _VtblGap4_8();

    bool AllowAutoFit { [DispId(110)] get; [DispId(110)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap5_2();

    WdPreferredWidthType PreferredWidthType { [DispId(112)] get; [DispId(112)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap6_8();

    float Spacing { [DispId(117)] get; [DispId(117)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap7_4();

    object Style { [DispId(201)] [return: MarshalAs(UnmanagedType.Struct)] get; [DispId(201)] [param: MarshalAs(UnmanagedType.Struct), In] set; }

    [SpecialName]
    sealed extern void _VtblGap8_15();

    string Descr { [DispId(210)] [return: MarshalAs(UnmanagedType.BStr)] get; [DispId(210)] [param: MarshalAs(UnmanagedType.BStr), In] set; }
  }
}
