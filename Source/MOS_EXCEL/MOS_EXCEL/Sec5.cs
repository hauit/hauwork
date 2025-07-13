// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Sec5
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
{
  internal class Sec5
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec5.cau1(a, d);
        case 2:
          return Sec5.cau2(a, d);
        case 3:
          return Sec5.cau3(a, d);
        case 4:
          return Sec5.cau4(a, d);
        case 5:
          return Sec5.cau5(a, d);
        case 6:
          return Sec5.cau6(a, d);
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
        if (Sec5.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec5.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec5)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec5.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) Sec5.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, d.Worksheets[(object) "Fiction"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.Shapes.Count != 1)
          return "False (Number of picture)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.Shapes.Item((object) "picture 1").Rotation.ToString() != "0")
          return "False(quay 0 do)";
      }
      catch (Exception ex)
      {
        return "False (Number of picture)";
      }
      return "True";
    }

    private static string cau2(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec5.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec5.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec5)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec5.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3.Target((CallSite) Sec5.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3, d.Worksheets[(object) "October"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.Shapes.Count != 2)
          return "False (add picture)";
        // ISSUE: reference to a compiler-generated method
        if ((double) worksheet.Shapes.Item((object) 2).Left < 400.0)
          return "False(di chuyển đến bênh phải tiêu đề)";
        // ISSUE: reference to a compiler-generated method
        if ((double) worksheet.Shapes.Item((object) 2).Top > 100.0)
          return "False(di chuyển đến sau tiêu đề)";
      }
      catch (Exception ex)
      {
        return "False (chèn hình)";
      }
      return "True";
    }

    private static string cau3(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec5.\u003Ccau3\u003Eo__SiteContainer4.\u003C\u003Ep__Site5 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec5.\u003Ccau3\u003Eo__SiteContainer4.\u003C\u003Ep__Site5 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec5)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec5.\u003Ccau3\u003Eo__SiteContainer4.\u003C\u003Ep__Site5.Target((CallSite) Sec5.\u003Ccau3\u003Eo__SiteContainer4.\u003C\u003Ep__Site5, d.Worksheets[(object) "Non_Fiction"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.Shapes.Count != 1)
          return "False (khong them xoa hinh)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.Shapes.Item((object) 1).Fill.PictureEffects.Count != 1)
          return "False(chỉ remove backgroup)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.Shapes.Item((object) 1).Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectBackgroundRemoval)
          return "False(remove backgroup)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.Shapes.Item((object) 1).Fill.Pattern != MsoPatternType.msoPattern20Percent)
          return "False(Fill Pattern20)";
      }
      catch (Exception ex)
      {
        return "False (khong them xoa hinh)";
      }
      return "True";
    }

    private static string cau4(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec5.\u003Ccau4\u003Eo__SiteContainer6.\u003C\u003Ep__Site7 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec5.\u003Ccau4\u003Eo__SiteContainer6.\u003C\u003Ep__Site7 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec5)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec5.\u003Ccau4\u003Eo__SiteContainer6.\u003C\u003Ep__Site7.Target((CallSite) Sec5.\u003Ccau4\u003Eo__SiteContainer6.\u003C\u003Ep__Site7, d.Worksheets[(object) "Summary"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.Shapes.Count != 1)
          return "False (khong them xoa đối tượng)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.Shapes.Item((object) 1).AlternativeText != "Renewal data")
          return "False(sai chính tả)";
      }
      catch (Exception ex)
      {
        return "False (khong them xoa hinh)";
      }
      return "True";
    }

    private static string cau5(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec5.\u003Ccau5\u003Eo__SiteContainer8.\u003C\u003Ep__Site9 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec5.\u003Ccau5\u003Eo__SiteContainer8.\u003C\u003Ep__Site9 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec5)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec5.\u003Ccau5\u003Eo__SiteContainer8.\u003C\u003Ep__Site9.Target((CallSite) Sec5.\u003Ccau5\u003Eo__SiteContainer8.\u003C\u003Ep__Site9, d.Worksheets[(object) "Summary"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tinh Summary)";
      }
      try
      {
        if (worksheet.Shapes.Count != 1)
          return "False (không thêm xóa đối tượng)";
        // ISSUE: reference to a compiler-generated method
        string str = worksheet.Shapes.Item((object) 1).AutoShapeType.ToString();
        // ISSUE: reference to a compiler-generated method
        if (worksheet.Shapes.Item((object) 1).Type != MsoShapeType.msoAutoShape)
          return "False (not Shape)";
        if (str != "msoShapeVerticalScroll")
          return "False (wrong Shape)";
      }
      catch (Exception ex)
      {
        return "False (Number of shape)";
      }
      return "True";
    }

    private static string cau6(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec5.\u003Ccau6\u003Eo__SiteContainera.\u003C\u003Ep__Siteb == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec5.\u003Ccau6\u003Eo__SiteContainera.\u003C\u003Ep__Siteb = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec5)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec5.\u003Ccau6\u003Eo__SiteContainera.\u003C\u003Ep__Siteb.Target((CallSite) Sec5.\u003Ccau6\u003Eo__SiteContainera.\u003C\u003Ep__Siteb, d.Worksheets[(object) "Olympic Men Single Sculls"]);
      }
      catch (Exception ex)
      {
        return "Fales (Not found Olympic Men Single Sculls Worksheet)";
      }
      try
      {
        if (worksheet.Shapes.Count != 1)
          return "False (Number of Shape)";
        // ISSUE: reference to a compiler-generated method
        string str = worksheet.Shapes.Item((object) 1).AutoShapeType.ToString();
        // ISSUE: reference to a compiler-generated method
        if (worksheet.Shapes.Item((object) 1).Type != MsoShapeType.msoAutoShape)
          return "False (not Shape)";
        if (str != "msoShapeVerticalScroll")
          return "False (wrong Shape)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.Shapes.Item((object) 1).TextFrame2.TextRange.Text != "Top experts - from trail runners to CEOs to beloved authors - reveal the trails that fuel their dreams.")
          return "False (wrong text)";
      }
      catch (Exception ex)
      {
        return "False (Number of shape)";
      }
      return "True";
    }
  }
}
