// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.PageSetup
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [CompilerGenerated]
  [Guid("00020971-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [ComImport]
  public interface PageSetup
  {
    [SpecialName]
    sealed extern void _VtblGap1_3();

    float TopMargin { [DispId(100)] get; [DispId(100)] [param: In] set; }

    float BottomMargin { [DispId(101)] get; [DispId(101)] [param: In] set; }

    float LeftMargin { [DispId(102)] get; [DispId(102)] [param: In] set; }

    float RightMargin { [DispId(103)] get; [DispId(103)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_6();

    WdOrientation Orientation { [DispId(107)] get; [DispId(107)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap3_10();

    float FooterDistance { [DispId(113)] get; [DispId(113)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap4_4();

    int DifferentFirstPageHeaderFooter { [DispId(116)] get; [DispId(116)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap5_4();

    TextColumns TextColumns { [DispId(119)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(119)] [param: MarshalAs(UnmanagedType.Interface), In] set; }
  }
}
