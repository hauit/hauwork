// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word._Font
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("00020952-0000-0000-C000-000000000046")]
  [CompilerGenerated]
  [TypeIdentifier]
  [ComImport]
  public interface _Font
  {
    [SpecialName]
    sealed extern void _VtblGap1_4();

    int Bold { [DispId(130)] get; [DispId(130)] [param: In] set; }

    int Italic { [DispId(131)] get; [DispId(131)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_16();

    WdUnderline Underline { [DispId(140)] get; [DispId(140)] [param: In] set; }

    float Size { [DispId(141)] get; [DispId(141)] [param: In] set; }

    string Name { [DispId(142)] [return: MarshalAs(UnmanagedType.BStr)] get; [DispId(142)] [param: MarshalAs(UnmanagedType.BStr), In] set; }

    [SpecialName]
    sealed extern void _VtblGap3_49();

    WdColor UnderlineColor { [DispId(166)] get; [DispId(166)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap4_4();

    ShadowFormat TextShadow { [DispId(169)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(169)] [param: MarshalAs(UnmanagedType.Interface), In] set; }

    FillFormat Fill { [DispId(170)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(170)] [param: MarshalAs(UnmanagedType.Interface), In] set; }

    [SpecialName]
    sealed extern void _VtblGap5_4();

    ColorFormat TextColor { [DispId(173)] [return: MarshalAs(UnmanagedType.Interface)] get; }
  }
}
