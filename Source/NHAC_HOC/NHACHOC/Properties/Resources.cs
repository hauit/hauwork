// Decompiled with JetBrains decompiler
// Type: NHACHOC.Properties.Resources
// Assembly: NHACHOC, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 203B0FF7-2323-40ED-93E4-470E5FE69D30
// Assembly location: C:\Users\Admin\Downloads\WE_19\GOC HO TRO\NHAC HOC\NHACHOC.exe

using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Resources;
using System.Runtime.CompilerServices;

namespace NHACHOC.Properties
{
  [CompilerGenerated]
  [GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
  [DebuggerNonUserCode]
  internal class Resources
  {
    private static ResourceManager resourceMan;
    private static CultureInfo resourceCulture;

    internal Resources()
    {
    }

    [EditorBrowsable(EditorBrowsableState.Advanced)]
    internal static ResourceManager ResourceManager
    {
      get
      {
        if (object.ReferenceEquals((object) NHACHOC.Properties.Resources.resourceMan, (object) null))
          NHACHOC.Properties.Resources.resourceMan = new ResourceManager("NHACHOC.Properties.Resources", typeof (NHACHOC.Properties.Resources).Assembly);
        return NHACHOC.Properties.Resources.resourceMan;
      }
    }

    [EditorBrowsable(EditorBrowsableState.Advanced)]
    internal static CultureInfo Culture
    {
      get => NHACHOC.Properties.Resources.resourceCulture;
      set => NHACHOC.Properties.Resources.resourceCulture = value;
    }

    internal static Bitmap _2022_08_25 => (Bitmap) NHACHOC.Properties.Resources.ResourceManager.GetObject("2022-08-25", NHACHOC.Properties.Resources.resourceCulture);

    internal static Bitmap hinh_anh_quyet_tam_on_thi_hoc_bai_17 => (Bitmap) NHACHOC.Properties.Resources.ResourceManager.GetObject("hinh-anh-quyet-tam-on-thi-hoc-bai-17", NHACHOC.Properties.Resources.resourceCulture);
  }
}
