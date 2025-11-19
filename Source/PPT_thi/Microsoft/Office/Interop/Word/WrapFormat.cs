// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.WrapFormat
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("000209C3-0000-0000-C000-000000000046")]
  [CompilerGenerated]
  [TypeIdentifier]
  [ComImport]
  public interface WrapFormat
  {
    [SpecialName]
    sealed extern void _VtblGap1_3();

    WdWrapType Type { [DispId(100)] get; [DispId(100)] [param: In] set; }

    WdWrapSideType Side { [DispId(101)] get; [DispId(101)] [param: In] set; }
  }
}
