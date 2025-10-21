// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Style
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("0002092C-0000-0000-C000-000000000046")]
  [CompilerGenerated]
  [DefaultMember("NameLocal")]
  [TypeIdentifier]
  [ComImport]
  public interface Style
  {
    [SpecialName]
    sealed extern void _VtblGap1_3();

    string NameLocal { [DispId(0)] [return: MarshalAs(UnmanagedType.BStr)] get; [DispId(0)] [param: MarshalAs(UnmanagedType.BStr), In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_9();

    Borders Borders { [DispId(8)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(8)] [param: MarshalAs(UnmanagedType.Interface), In] set; }

    ParagraphFormat ParagraphFormat { [DispId(9)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(9)] [param: MarshalAs(UnmanagedType.Interface), In] set; }

    Font Font { [DispId(10)] [return: MarshalAs(UnmanagedType.Interface)] get; [DispId(10)] [param: MarshalAs(UnmanagedType.Interface), In] set; }
  }
}
