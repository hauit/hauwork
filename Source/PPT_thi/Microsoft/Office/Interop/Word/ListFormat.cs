// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.ListFormat
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("000209C0-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [CompilerGenerated]
  [ComImport]
  public interface ListFormat
  {
    int ListLevelNumber { [DispId(68)] get; [DispId(68)] [param: In] set; }

    List List { [DispId(69)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    ListTemplate ListTemplate { [DispId(70)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    int ListValue { [DispId(71)] get; }

    [SpecialName]
    sealed extern void _VtblGap1_2();

    WdListType ListType { [DispId(74)] get; }

    string ListString { [DispId(75)] [return: MarshalAs(UnmanagedType.BStr)] get; }

    [SpecialName]
    sealed extern void _VtblGap2_17();

    InlineShape ListPictureBullet { [DispId(76)] [return: MarshalAs(UnmanagedType.Interface)] get; }
  }
}
