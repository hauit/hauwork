// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word._ParagraphFormat
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [TypeIdentifier]
  [CompilerGenerated]
  [Guid("00020953-0000-0000-C000-000000000046")]
  [ComImport]
  public interface _ParagraphFormat
  {
    [SpecialName]
    sealed extern void _VtblGap1_6();

    WdParagraphAlignment Alignment { [DispId(101)] get; [DispId(101)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_14();

    float LineSpacing { [DispId(109)] get; [DispId(109)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap3_2();

    float SpaceBefore { [DispId(111)] get; [DispId(111)] [param: In] set; }

    float SpaceAfter { [DispId(112)] get; [DispId(112)] [param: In] set; }
  }
}
