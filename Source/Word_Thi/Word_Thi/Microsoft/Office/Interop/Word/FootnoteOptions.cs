// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.FootnoteOptions
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [CompilerGenerated]
  [TypeIdentifier]
  [Guid("BEA85A24-D7DA-4F3D-B58C-ED90FB01D615")]
  [ComImport]
  public interface FootnoteOptions
  {
    [SpecialName]
    sealed extern void _VtblGap1_5();

    WdNoteNumberStyle NumberStyle { [DispId(101)] get; [DispId(101)] [param: In] set; }

    int StartingNumber { [DispId(102)] get; [DispId(102)] [param: In] set; }
  }
}
