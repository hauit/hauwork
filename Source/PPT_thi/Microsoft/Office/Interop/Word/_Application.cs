// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word._Application
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("00020970-0000-0000-C000-000000000046")]
  [CompilerGenerated]
  [DefaultMember("Name")]
  [TypeIdentifier]
  [ComImport]
  public interface _Application
  {
    [SpecialName]
    sealed extern void _VtblGap1_4();

    Documents Documents { [DispId(6)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap2_2();

    Window ActiveWindow { [DispId(4)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap3_5();

    AutoCorrect AutoCorrect { [DispId(10)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap4_12();

    bool Visible { [DispId(23)] get; [DispId(23)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap5_49();

    int Left { [DispId(87)] get; [DispId(87)] [param: In] set; }

    int Top { [DispId(88)] get; [DispId(88)] [param: In] set; }

    int Width { [DispId(89)] get; [DispId(89)] [param: In] set; }

    int Height { [DispId(90)] get; [DispId(90)] [param: In] set; }

    WdWindowState WindowState { [DispId(91)] get; [DispId(91)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap6_2();

    Options Options { [DispId(93)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap7_23();

    [DispId(1105)]
    void Quit([MarshalAs(UnmanagedType.Struct), In, Optional] ref object SaveChanges, [MarshalAs(UnmanagedType.Struct), In, Optional] ref object OriginalFormat, [MarshalAs(UnmanagedType.Struct), In, Optional] ref object RouteDocument);

    [SpecialName]
    sealed extern void _VtblGap8_39();

    [DispId(361)]
    void Resize([In] int Width, [In] int Height);
  }
}
