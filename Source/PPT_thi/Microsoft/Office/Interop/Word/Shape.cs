// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Shape
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using Microsoft.Office.Core;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("000209A0-0000-0000-C000-000000000046")]
  [CompilerGenerated]
  [TypeIdentifier]
  [ComImport]
  public interface Shape
  {
    [SpecialName]
    sealed extern void _VtblGap1_10();

    FillFormat Fill { [DispId(107)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    GroupShapes GroupItems { [DispId(108)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    float Height { [DispId(109)] get; [DispId(109)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_1();

    float Left { [DispId(111)] get; [DispId(111)] [param: In] set; }

    LineFormat Line { [DispId(112)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap3_2();

    string Name { [DispId(115)] [return: MarshalAs(UnmanagedType.BStr)] get; [DispId(115)] [param: MarshalAs(UnmanagedType.BStr), In] set; }

    [SpecialName]
    sealed extern void _VtblGap4_1();

    float Rotation { [DispId(117)] get; [DispId(117)] [param: In] set; }

    PictureFormat PictureFormat { [DispId(118)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap5_2();

    TextFrame TextFrame { [DispId(121)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap6_1();

    float Top { [DispId(123)] get; [DispId(123)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap7_5();

    float Width { [DispId(128)] get; [DispId(128)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap8_2();

    WdRelativeHorizontalPosition RelativeHorizontalPosition { [DispId(300)] get; [DispId(300)] [param: In] set; }

    WdRelativeVerticalPosition RelativeVerticalPosition { [DispId(301)] get; [DispId(301)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap9_2();

    WrapFormat WrapFormat { [DispId(303)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap10_31();

    int ID { [DispId(139)] get; }

    [SpecialName]
    sealed extern void _VtblGap11_9();

    float LeftRelative { [DispId(200)] get; [DispId(200)] [param: In] set; }

    float TopRelative { [DispId(201)] get; [DispId(201)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap12_4();

    WdRelativeHorizontalSize RelativeHorizontalSize { [DispId(204)] get; [DispId(204)] [param: In] set; }

    WdRelativeVerticalSize RelativeVerticalSize { [DispId(205)] get; [DispId(205)] [param: In] set; }

    SoftEdgeFormat SoftEdge { [DispId(152)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    GlowFormat Glow { [DispId(153)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap13_3();

    SmartArt SmartArt { [DispId(156)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    MsoShapeStyleIndex ShapeStyle { [DispId(150)] get; [DispId(150)] [param: In] set; }
  }
}
