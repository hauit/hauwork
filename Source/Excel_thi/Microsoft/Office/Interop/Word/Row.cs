// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Row
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [DefaultMember("Range")]
  [TypeIdentifier]
  [CompilerGenerated]
  [Guid("00020950-0000-0000-C000-000000000046")]
  [ComImport]
  public interface Row
  {
    Range Range { [DispId(0)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap1_20();

    Cells Cells { [DispId(100)] [return: MarshalAs(UnmanagedType.Interface)] get; }
  }
}
