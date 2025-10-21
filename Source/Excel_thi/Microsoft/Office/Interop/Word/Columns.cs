// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Columns
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [CompilerGenerated]
  [Guid("0002094B-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [ComImport]
  public interface Columns : IEnumerable
  {
    [SpecialName]
    sealed extern void _VtblGap1_1();

    int Count { [DispId(2)] get; }

    [SpecialName]
    sealed extern void _VtblGap2_5();

    float Width { [DispId(3)] get; [DispId(3)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap3_3();

    Column this[[In] int Index] { [DispId(0)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap4_9();

    WdPreferredWidthType PreferredWidthType { [DispId(106)] get; [DispId(106)] [param: In] set; }
  }
}
