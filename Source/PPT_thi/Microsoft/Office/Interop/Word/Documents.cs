// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Documents
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("0002096C-0000-0000-C000-000000000046")]
  [CompilerGenerated]
  [TypeIdentifier]
  [ComImport]
  public interface Documents : IEnumerable
  {
    [SpecialName]
    sealed extern void _VtblGap1_1();

    int Count { [DispId(2)] get; }

    [SpecialName]
    sealed extern void _VtblGap2_3();

    Document this[[MarshalAs(UnmanagedType.Struct), In] ref object Index] { [DispId(0)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap3_9();

    [DispId(19)]
    [return: MarshalAs(UnmanagedType.Interface)]
    Document Open(
      [MarshalAs(UnmanagedType.Struct), In] ref object FileName,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object ConfirmConversions,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object ReadOnly,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object AddToRecentFiles,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object PasswordDocument,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object PasswordTemplate,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object Revert,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object WritePasswordDocument,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object WritePasswordTemplate,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object Format,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object Encoding,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object Visible,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object OpenAndRepair,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object DocumentDirection,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object NoEncodingDialog,
      [MarshalAs(UnmanagedType.Struct), In, Optional] ref object XMLTransform);
  }
}
