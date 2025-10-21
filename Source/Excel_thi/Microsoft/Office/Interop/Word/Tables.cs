// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.Tables
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.CustomMarshalers;

namespace Microsoft.Office.Interop.Word
{
  [Guid("0002094D-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [CompilerGenerated]
  [ComImport]
  public interface Tables : IEnumerable
  {
    [DispId(-4)]
    [return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalTypeRef = typeof (EnumeratorToEnumVariantMarshaler))]
    new IEnumerator GetEnumerator();

    int Count { [DispId(2)] get; }

    [SpecialName]
    sealed extern void _VtblGap1_3();

    Table this[[In] int Index] { [DispId(0)] [return: MarshalAs(UnmanagedType.Interface)] get; }
  }
}
