// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Options
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [TypeIdentifier]
  [Guid("000209B7-0000-0000-C000-000000000046")]
  [CompilerGenerated]
  [ComImport]
  public interface Options
  {
    [SpecialName]
    sealed extern void _VtblGap1_31();

    bool PrintProperties { [DispId(31)] get; [DispId(31)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_25();

    int SaveInterval { [DispId(45)] get; [DispId(45)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap3_108();

    bool CheckSpellingAsYouType { [DispId(276)] get; [DispId(276)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap4_227();

    bool PrintBackgrounds { [DispId(456)] get; [DispId(456)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap5_28();

    bool EnableLivePreview { [DispId(472)] get; [DispId(472)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap6_30();

    bool AllowOpenInDraftView { [DispId(490)] get; [DispId(490)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap7_18();

    bool UpdateFieldsWithTrackedChangesAtPrint { [DispId(503)] get; [DispId(503)] [param: In] set; }
  }
}
