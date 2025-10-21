// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Pane
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("00020960-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [CompilerGenerated]
  [ComImport]
  public interface Pane
  {
    [SpecialName]
    sealed extern void _VtblGap1_5();

    bool DisplayRulers { [DispId(4)] get; [DispId(4)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_4();

    View View { [DispId(10)] [return: MarshalAs(UnmanagedType.Interface)] get; }
  }
}
