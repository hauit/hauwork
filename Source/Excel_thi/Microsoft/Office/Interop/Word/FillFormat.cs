// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.FillFormat
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using Microsoft.Office.Core;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("000209C8-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [CompilerGenerated]
  [ComImport]
  public interface FillFormat
  {
    [SpecialName]
    sealed extern void _VtblGap1_4();

    ColorFormat ForeColor { [DispId(101)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap2_38();

    PictureEffects PictureEffects { [DispId(122)] [return: MarshalAs(UnmanagedType.Interface)] get; }
  }
}
