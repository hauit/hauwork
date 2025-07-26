// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word._Document
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("0002096B-0000-0000-C000-000000000046")]
  [CompilerGenerated]
  [DefaultMember("Name")]
  [TypeIdentifier]
  [ComImport]
  public interface _Document
  {
    string Name { [DispId(0)] [return: MarshalAs(UnmanagedType.BStr)] get; }

    [SpecialName]
    sealed extern void _VtblGap1_3();

    object BuiltInDocumentProperties { [DispId(1000)] [return: MarshalAs(UnmanagedType.IDispatch)] get; }

    [SpecialName]
    sealed extern void _VtblGap2_2();

    Bookmarks Bookmarks { [DispId(4)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    Tables Tables { [DispId(6)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    Footnotes Footnotes { [DispId(7)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    Endnotes Endnotes { [DispId(8)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    Comments Comments { [DispId(9)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap3_9();

    Sections Sections { [DispId(15)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    Paragraphs Paragraphs { [DispId(16)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap4_3();

    Fields Fields { [DispId(20)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap5_1();

    Styles Styles { [DispId(22)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap6_9();

    PageSetup PageSetup { [DispId(1101)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(1101)] [param: MarshalAs(UnmanagedType.Interface), In] set; }

    [SpecialName]
    sealed extern void _VtblGap7_10();

    Window ActiveWindow { [DispId(42)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap8_7();

    bool EmbedTrueTypeFonts { [DispId(50)] get; [DispId(50)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap9_12();

    WdProtectionType ProtectionType { [DispId(60)] get; }

    Hyperlinks Hyperlinks { [DispId(61)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    Shapes Shapes { [DispId(62)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap10_6();

    InlineShapes InlineShapes { [DispId(68)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap11_6();

    bool ShowGrammaticalErrors { [DispId(72)] get; [DispId(72)] [param: In] set; }

    bool ShowSpellingErrors { [DispId(73)] get; [DispId(73)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap12_68();

    [DispId(1105)]
    void Close([MarshalAs(UnmanagedType.Struct), In, Optional] ref object SaveChanges, [MarshalAs(UnmanagedType.Struct), In, Optional] ref object OriginalFormat, [MarshalAs(UnmanagedType.Struct), In, Optional] ref object RouteDocument);

    [SpecialName]
    sealed extern void _VtblGap13_96();

    bool RemovePersonalInformation { [DispId(344)] get; [DispId(344)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap14_109();

    Bibliography Bibliography { [DispId(516)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap15_57();

    int CompatibilityMode { [DispId(567)] get; }
  }
}
