// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Source
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word
{
  [TypeIdentifier]
  [Guid("4A6AE865-199D-4EA3-9F6B-125BD9C40EDF")]
  [CompilerGenerated]
  [ComImport]
  public interface Source
  {
    [SpecialName]
    sealed extern void _VtblGap1_4();

    [DispId(104)]
    [return: MarshalAs(UnmanagedType.BStr)]
    string get_Field([MarshalAs(UnmanagedType.BStr), In] string Name);

    [DispId(104)]
    void set_Field([MarshalAs(UnmanagedType.BStr), In] string Name, [MarshalAs(UnmanagedType.BStr), In] string prop);
  }
}
