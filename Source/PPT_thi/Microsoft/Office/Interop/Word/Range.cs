// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Range
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("0002095E-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [CompilerGenerated]
  [DefaultMember("Text")]
  [ComImport]
  public interface Range
  {
    string Text { [DispId(0)] [return: MarshalAs(UnmanagedType.BStr)] get; [DispId(0)] [param: MarshalAs(UnmanagedType.BStr), In] set; }

    [SpecialName]
    sealed extern void _VtblGap1_6();

    Font Font { [DispId(5)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(5)] [param: MarshalAs(UnmanagedType.Interface), In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_2();

    Tables Tables { [DispId(50)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap3_2();

    Characters Characters { [DispId(53)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap4_14();

    ParagraphFormat ParagraphFormat { [DispId(1102)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(1102)] [param: MarshalAs(UnmanagedType.Interface), In] set; }

    ListFormat ListFormat { [DispId(68)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap5_4();

    int Bold { [DispId(130)] get; [DispId(130)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap6_2();

    WdUnderline Underline { [DispId(139)] get; [DispId(139)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap7_5();

    object Style { [DispId(151)] [return: MarshalAs(UnmanagedType.Struct)] get; [DispId(151)] [param: MarshalAs(UnmanagedType.Struct), In] set; }

    [SpecialName]
    sealed extern void _VtblGap8_32();

    InlineShapes InlineShapes { [DispId(319)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap9_103();

    FootnoteOptions FootnoteOptions { [DispId(410)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    EndnoteOptions EndnoteOptions { [DispId(411)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap10_13();

    object CharacterStyle { [DispId(420)] [return: MarshalAs(UnmanagedType.Struct)] get; }

    object ParagraphStyle { [DispId(421)] [return: MarshalAs(UnmanagedType.Struct)] get; }
  }
}
