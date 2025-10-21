// Decompiled with JetBrains decompiler
// Type: Microsoft.Office.Interop.Word.TabStops
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.CustomMarshalers;

namespace Microsoft.Office.Interop.Word
{
  [Guid("00020955-0000-0000-C000-000000000046")]
  [TypeIdentifier]
  [CompilerGenerated]
  [DefaultMember("Item")]
  [ComImport]
  public interface TabStops : IEnumerable
  {
    [DispId(-4)]
    [return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalTypeRef = typeof (EnumeratorToEnumVariantMarshaler))]
    new IEnumerator GetEnumerator();
  }
}
