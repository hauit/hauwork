// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Paragraph
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [CompilerGenerated]
  [Guid("00020957-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [DefaultMember("Range")]
  [ComImport]
  public interface Paragraph
  {
    Range Range { [DispId(0)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap1_3();

    ParagraphFormat Format { [DispId(1102)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(1102)] [param: MarshalAs(UnmanagedType.Interface), In] set; }

    TabStops TabStops { [DispId(1103)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(1103)] [param: MarshalAs(UnmanagedType.Interface), In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_2();

    DropCap DropCap { [DispId(13)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    object Style { [DispId(100)] [return: MarshalAs(UnmanagedType.Struct)] get; [DispId(100)] [param: MarshalAs(UnmanagedType.Struct), In] set; }

    [SpecialName]
    sealed extern void _VtblGap3_12();

    float LeftIndent { [DispId(107)] get; [DispId(107)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap4_2();

    float LineSpacing { [DispId(109)] get; [DispId(109)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap5_2();

    float SpaceBefore { [DispId(111)] get; [DispId(111)] [param: In] set; }

    float SpaceAfter { [DispId(112)] get; [DispId(112)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap6_23();

    WdOutlineLevel OutlineLevel { [DispId(202)] get; [DispId(202)] [param: In] set; }
  }
}
