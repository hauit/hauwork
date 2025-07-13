// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Sec8
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
{
  internal class Sec8
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec8.cau2(a, d);
        case 2:
          return Sec8.cau3(a, d);
        case 3:
          return Sec8.cau4(a, d);
        case 4:
          return Sec8.cau1(a, d);
        case 5:
          return Sec8.cau5(a, d);
        default:
          return "False";
      }
    }

    private static string cau1(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec8.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec8.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec8)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec8.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) Sec8.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, d.Worksheets[(object) "Materials"]);
      }
      catch (Exception ex)
      {
        return "Fales (Ten trang tinh)";
      }
      try
      {
        if (worksheet.PageSetup.RightHeader != "Confidential")
          return "False(Confidential)";
      }
      catch (Exception ex)
      {
        return "False (Left header)";
      }
      return "True";
    }

    private static string cau2(Application a, Workbook d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec8.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec8.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec8)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec8.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3.Target((CallSite) Sec8.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3, d.Worksheets[(object) "Materials"]);
      }
      catch (Exception ex)
      {
        return "Fales (Sale worksheet not found)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "A6", (object) "A6").Hyperlinks.Count != 1)
          return "False (Number of hyperlink)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "A6", (object) "A6").Hyperlinks[(object) 1].SubAddress != "Categories!A18")
        {
          // ISSUE: reference to a compiler-generated method
          return "False (" + worksheet.get_Range((object) "A6", (object) "A6").Hyperlinks[(object) 1].SubAddress + ")";
        }
      }
      catch (Exception ex)
      {
        return "False (Not apply hyperlink to A6)";
      }
      return str;
    }

    private static string cau3(Application a, Workbook d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec8.\u003Ccau3\u003Eo__SiteContainer4.\u003C\u003Ep__Site5 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec8.\u003Ccau3\u003Eo__SiteContainer4.\u003C\u003Ep__Site5 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec8)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec8.\u003Ccau3\u003Eo__SiteContainer4.\u003C\u003Ep__Site5.Target((CallSite) Sec8.\u003Ccau3\u003Eo__SiteContainer4.\u003C\u003Ep__Site5, d.Worksheets[(object) "Shareholders Info"]);
      }
      catch (Exception ex)
      {
        return "Fales (Olympic Medals worksheet not found)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "C5", (object) "C5").Hyperlinks.Count != 1)
          return "False (Number of hyperlink)";
        // ISSUE: reference to a compiler-generated method
        if (!worksheet.get_Range((object) "C5", (object) "C5").Hyperlinks[(object) 1].Address.Contains("tailspintoys.com/beyond.html"))
        {
          // ISSUE: reference to a compiler-generated method
          return "False (" + worksheet.get_Range((object) "C5", (object) "C5").Hyperlinks[(object) 1].Address + ")";
        }
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "C5", (object) "C5").Hyperlinks[(object) 1].TextToDisplay != "More Info")
          return "False(More Info)";
      }
      catch (Exception ex)
      {
        return "False (Not apply hyperlink to D11)";
      }
      return str;
    }

    private static string cau4(Application a, Workbook d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec8.\u003Ccau4\u003Eo__SiteContainer6.\u003C\u003Ep__Site7 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec8.\u003Ccau4\u003Eo__SiteContainer6.\u003C\u003Ep__Site7 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec8)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec8.\u003Ccau4\u003Eo__SiteContainer6.\u003C\u003Ep__Site7.Target((CallSite) Sec8.\u003Ccau4\u003Eo__SiteContainer6.\u003C\u003Ep__Site7, d.Worksheets[(object) "Summary"]);
      }
      catch (Exception ex)
      {
        return "Fales (Summary worksheet not found)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "A2", (object) "A2").Hyperlinks.Count != 1)
          return "False (chèn hyperlink tại A2)";
        // ISSUE: reference to a compiler-generated method
        if (!worksheet.get_Range((object) "A2", (object) "A2").Hyperlinks[(object) 1].Address.Contains("www.nodpublishers.com"))
        {
          // ISSUE: reference to a compiler-generated method
          return "False (" + worksheet.get_Range((object) "A2", (object) "A2").Hyperlinks[(object) 1].Address + ")";
        }
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "A2", (object) "A2").Hyperlinks[(object) 1].ScreenTip != "Company Website")
          return "False(Company Website)";
      }
      catch (Exception ex)
      {
        return "False (Not apply hyperlink to D11)";
      }
      return str;
    }

    private static string cau5(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec8.\u003Ccau5\u003Eo__SiteContainer8.\u003C\u003Ep__Site9 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec8.\u003Ccau5\u003Eo__SiteContainer8.\u003C\u003Ep__Site9 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec8)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec8.\u003Ccau5\u003Eo__SiteContainer8.\u003C\u003Ep__Site9.Target((CallSite) Sec8.\u003Ccau5\u003Eo__SiteContainer8.\u003C\u003Ep__Site9, d.Worksheets[(object) "Materials"]);
      }
      catch (Exception ex)
      {
        return "Fales (Ten trang tinh)";
      }
      try
      {
        if (worksheet.PageSetup.CenterFooter != "Page &P of &N")
          return "False(Page 1 of ?)";
      }
      catch (Exception ex)
      {
        return "False (Left header)";
      }
      return "True";
    }

    private static string cau6(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec8.\u003Ccau6\u003Eo__SiteContainera.\u003C\u003Ep__Siteb == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec8.\u003Ccau6\u003Eo__SiteContainera.\u003C\u003Ep__Siteb = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec8)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec8.\u003Ccau6\u003Eo__SiteContainera.\u003C\u003Ep__Siteb.Target((CallSite) Sec8.\u003Ccau6\u003Eo__SiteContainera.\u003C\u003Ep__Siteb, d.Worksheets[(object) "Materials"]);
        if (!a.ActiveWindow.FreezePanes)
          return "False(FreezePanes)";
        // ISSUE: reference to a compiler-generated method
        if (a.ActiveCell.get_Address((object) Missing.Value, (object) Missing.Value, External: (object) Missing.Value, RelativeTo: (object) Missing.Value) != "$A$6")
          return "False(để con trỏ ô A6 trước khi FreezePanes)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }
  }
}
