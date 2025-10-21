// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.ApplicationEvents4_Event
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [CompilerGenerated]
  [TypeIdentifier("00020905-0000-0000-c000-000000000046", "Microsoft.Office.Interop.Word.ApplicationEvents4_Event")]
  [ComEventInterface(typeof (ApplicationEvents4), typeof (ApplicationEvents4))]
  [ComImport]
  public interface ApplicationEvents4_Event
  {
    [SpecialName]
    sealed extern void _VtblGap1_8();

    event ApplicationEvents4_DocumentBeforeCloseEventHandler DocumentBeforeClose;
  }
}
