// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Section
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("00020959-0000-0000-C000-000000000046")]
  [DefaultMember("Range")]
  [TypeIdentifier]
  [CompilerGenerated]
  [ComImport]
  public interface Section
  {
    Range Range { [DispId(0)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap1_3();

    PageSetup PageSetup { [DispId(1101)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(1101)] [param: MarshalAs(UnmanagedType.Interface), In] set; }

    HeadersFooters Headers { [DispId(121)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    HeadersFooters Footers { [DispId(122)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap2_3();

    Borders Borders { [DispId(1100)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(1100)] [param: MarshalAs(UnmanagedType.Interface), In] set; }
  }
}
