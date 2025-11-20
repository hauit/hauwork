// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.InlineShape
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using Microsoft.Office.Core;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("000209A8-0000-0000-C000-000000000046")]
  [CompilerGenerated]
  [TypeIdentifier]
  [ComImport]
  public interface InlineShape
  {
    [SpecialName]
    sealed extern void _VtblGap1_9();

    WdInlineShapeType Type { [DispId(6)] get; }

    [SpecialName]
    sealed extern void _VtblGap2_1();

    float Height { [DispId(8)] get; [DispId(8)] [param: In] set; }

    float Width { [DispId(9)] get; [DispId(9)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap3_6();

    LineFormat Line { [DispId(112)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap4_13();

    string AlternativeText { [DispId(131)] [return: MarshalAs(UnmanagedType.BStr)] get; [DispId(131)] [param: MarshalAs(UnmanagedType.BStr), In] set; }

    bool IsPictureBullet { [DispId(132)] get; }

    [SpecialName]
    sealed extern void _VtblGap5_8();

    SmartArt SmartArt { [DispId(156)] [return: MarshalAs(UnmanagedType.Interface)] get; }
  }
}
