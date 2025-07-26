// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Cell
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("0002094E-0000-0000-C000-000000000046")]
  [DefaultMember("Range")]
  [TypeIdentifier]
  [CompilerGenerated]
  [ComImport]
  public interface Cell
  {
    Range Range { [DispId(0)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap1_5();

    float Width { [DispId(6)] get; [DispId(6)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_4();

    WdCellVerticalAlignment VerticalAlignment { [DispId(1104)] get; [DispId(1104)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap3_4();

    Shading Shading { [DispId(105)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    Borders Borders { [DispId(1100)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(1100)] [param: MarshalAs(UnmanagedType.Interface), In] set; }
  }
}
