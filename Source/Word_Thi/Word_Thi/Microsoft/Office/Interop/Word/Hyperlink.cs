// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Hyperlink
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("0002099D-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [CompilerGenerated]
  [ComImport]
  public interface Hyperlink
  {
    [SpecialName]
    sealed extern void _VtblGap1_3();

    string Name { [DispId(1003)] [return: MarshalAs(UnmanagedType.BStr)] get; }

    [SpecialName]
    sealed extern void _VtblGap2_10();

    string Address { [DispId(1100)] [return: MarshalAs(UnmanagedType.BStr)] get; [DispId(1100)] [param: MarshalAs(UnmanagedType.BStr), In] set; }

    string SubAddress { [DispId(1101)] [return: MarshalAs(UnmanagedType.BStr)] get; [DispId(1101)] [param: MarshalAs(UnmanagedType.BStr), In] set; }

    [SpecialName]
    sealed extern void _VtblGap3_4();

    string TextToDisplay { [DispId(1012)] [return: MarshalAs(UnmanagedType.BStr)] get; [DispId(1012)] [param: MarshalAs(UnmanagedType.BStr), In] set; }
  }
}
