// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Zoom
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("000209A6-0000-0000-C000-000000000046")]
  [CompilerGenerated]
  [DefaultMember("Percentage")]
  [TypeIdentifier]
  [ComImport]
  public interface Zoom
  {
    [SpecialName]
    sealed extern void _VtblGap1_3();

    int Percentage { [DispId(0)] get; [DispId(0)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_2();

    int PageRows { [DispId(2)] get; [DispId(2)] [param: In] set; }

    int PageColumns { [DispId(3)] get; [DispId(3)] [param: In] set; }
  }
}
