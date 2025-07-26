// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.View
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [Guid("000209A5-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [CompilerGenerated]
  [DefaultMember("Type")]
  [ComImport]
  public interface View
  {
    [SpecialName]
    sealed extern void _VtblGap1_9();

    bool ShowAll { [DispId(3)] get; [DispId(3)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap2_10();

    Zoom Zoom { [DispId(10)] [return: MarshalAs(UnmanagedType.Interface)] get; }

    [SpecialName]
    sealed extern void _VtblGap3_8();

    bool ShowTabs { [DispId(15)] get; [DispId(15)] [param: In] set; }

    bool ShowSpaces { [DispId(16)] get; [DispId(16)] [param: In] set; }

    [SpecialName]
    sealed extern void _VtblGap4_4();

    bool ShowHiddenText { [DispId(19)] get; [DispId(19)] [param: In] set; }
  }
}
