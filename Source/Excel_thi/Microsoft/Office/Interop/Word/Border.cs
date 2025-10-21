// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Border
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("0002093B-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [CompilerGenerated]
  [DefaultMember("Visible")]
  [ComImport]
  public interface Border
  {
    [SpecialName]
    sealed extern void _VtblGap1_5();

    WdColorIndex ColorIndex { [DispId(1)] get; [DispId(1)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_3();

    WdLineWidth LineWidth { [DispId(4)] get; [DispId(4)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap3_4();

    WdColor Color { [DispId(7)] get; [DispId(7)] [param: In] set; }
  }
}
