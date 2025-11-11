// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S5_Text
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;

#nullable disable
namespace WindowsFormsApplication1;

internal class S5_Text
{
  public static string CheckCau(int causo, Application a, Presentation d)
  {
    switch (causo)
    {
      case 1:
        return S5_Text.cau1(a, d);
      case 2:
        return S5_Text.cau2(a, d);
      case 3:
        return S5_Text.cau3(a, d);
      case 4:
        return S5_Text.cau4(a, d);
      case 5:
        return S5_Text.cau5(a, d);
      case 6:
        return S5_Text.cau6(a, d);
      case 7:
        return S5_Text.cau7(a, d);
      case 8:
        return S5_Text.cau8(a, d);
      case 9:
        return S5_Text.cau9(a, d);
      case 10:
        return S5_Text.cau10(a, d);
      case 11:
        return S5_Text.cau11(a, d);
      case 12:
        return S5_Text.cau12(a, d);
      case 13:
        return S5_Text.cau13(a, d);
      case 14:
        return S5_Text.cau14(a, d);
      default:
        return "out Indext";
    }
  }

  private static string cau1(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 5].Shapes[(object) 3].TextFrame.TextRange.Text != "Company Confidential")
        return "False(Company Confidential)";
      if (!a.ActivePresentation.Slides[(object) 5].Shapes[(object) 3].Name.Contains("Footer Placeholder"))
        return "False(chen o footer)";
      if (a.ActivePresentation.Slides[(object) 6].Shapes.Count >= 3)
      {
        if (a.ActivePresentation.Slides[(object) 6].Shapes[(object) 3].TextFrame.TextRange.Text == "Company Confidential")
          return "False(chi cho trang 5)";
        if (a.ActivePresentation.Slides[(object) 6].Shapes[(object) 3].Name.Contains("Footer Placeholder"))
          return "False(chi cho trang 5)";
      }
    }
    catch (Exception ex)
    {
      return "False(them chu vao footer slide 5)";
    }
    return "True";
  }

  private static string cau2(Application a, Presentation d)
  {
    try
    {
      return a.ActivePresentation.Slides[(object) 3].Shapes[(object) "Content Placeholder 3"].Table.Style.Name != "Medium Style 1 - Accent 5" ? "False(Medium Style 1 - Accent 5)" : "True";
    }
    catch (Exception ex)
    {
      return "False (loi khong xac dinh)";
    }
  }

  private static string cau3(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 3].Hyperlinks.Count != 1)
        return "False(add hyperlink tren slide 3)";
      if (!a.ActivePresentation.Slides[(object) 3].Hyperlinks[1].Address.Contains("humongousinsurance.com"))
        return "False(http://www.humongousinsurance.com)";
      return a.ActivePresentation.Slides[(object) 3].Hyperlinks[1].TextToDisplay != "Click here to view on website" ? "False(text)" : "True";
    }
    catch (Exception ex)
    {
      return "False (add cho text khong phai cho shape)";
    }
  }

  private static string cau4(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 2].Comments.Count != 1)
        return "False(add 1 comment tren slide 2)";
      if (a.ActivePresentation.Slides[(object) 2].Comments[1].Text != "Update")
        return "False(Update)";
    }
    catch (Exception ex)
    {
      return "False(Gradient)";
    }
    return "True";
  }

  private static string cau5(Application a, Presentation d)
  {
    try
    {
      if (((object) a.ActivePresentation.Slides[(object) 3].Shapes[(object) "Content Placeholder 4"].Type).ToString() != "msoPlaceholder")
        return "False(table da bi xoa)";
      if (a.ActivePresentation.Slides[(object) 3].Shapes[(object) "Content Placeholder 4"].Table.Rows.Count.ToString() != "6")
        return "False(xoa dong 3)";
      if (a.ActivePresentation.Slides[(object) 3].Shapes[(object) "Content Placeholder 4"].Table.Rows[3].Cells[1].Shape.TextFrame.TextRange.Text.Contains("Sinusitis"))
        return "False(Sinusitis)";
      if (a.ActivePresentation.Slides[(object) 3].Shapes[(object) "Content Placeholder 4"].Table.Columns.Count.ToString() != "5")
        return "False(chen cot)";
      return !a.ActivePresentation.Slides[(object) 3].Shapes[(object) "Content Placeholder 4"].Table.Columns[5].Cells[1].Shape.TextFrame.TextRange.Text.Contains("Percentage Uninsured") ? "False(Percentage Uninsured)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something Wrong)";
    }
  }

  private static string cau6(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 7].Shapes[(object) "Table 1"].Table.Style.Name != "Medium Style 2 - Accent 1")
        return "False(Medium Style 2 - Accent 1)";
      if (a.ActivePresentation.Slides[(object) 7].Shapes[(object) "Table 1"].Table.HorizBanding.ToString() != "False")
        return "False(col Banded)";
      return a.ActivePresentation.Slides[(object) 7].Shapes[(object) "Table 1"].Table.VertBanding.ToString() != "True" ? "False(row Banded)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau7(Application a, Presentation d)
  {
    try
    {
      int count = a.ActivePresentation.Slides[(object) 3].Shapes[(object) "Content Placeholder 6"].Table.Columns.Count;
      if (count.ToString() != "3")
        return "False (xoa cot)";
      count = a.ActivePresentation.Slides[(object) 3].Shapes[(object) "Content Placeholder 6"].Table.Rows.Count;
      return count.ToString() != "7" ? "Fales (them dong)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau8(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 4].Shapes.Count != 2)
        return "False (chuyen doi tu noi dung thanh smartArt)";
      try
      {
        if (a.ActivePresentation.Slides[(object) 4].Shapes[(object) 2].SmartArt.Layout.Name != "Vertical Picture Accent List")
          return "False(Vertical Picture Accent List)";
      }
      catch (Exception ex)
      {
        return "False(chuyen doi tu noi dung thanh smartArt)";
      }
      return "True";
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau9(Application a, Presentation d)
  {
    try
    {
      if (((object) a.ActivePresentation.Slides[(object) 2].Shapes[(object) "TextBox 3"].TextFrame.VerticalAnchor).ToString() != "msoAnchorTop")
        return "False (Top)";
      return ((object) a.ActivePresentation.Slides[(object) 2].Shapes[(object) "TextBox 3"].TextFrame2.TextRange.get_Characters().Font.Caps).ToString() != "msoSmallCaps" ? "False(SmallCaps)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau10(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 5].Shapes.Count != 6)
        return "False(thêm highlighter)";
      return !a.ActivePresentation.Slides[(object) 5].Shapes[(object) 6].Name.Contains("nk") ? "False (thêm ink)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau11(Application a, Presentation d)
  {
    try
    {
      return a.ActivePresentation.Slides[(object) 7].Shapes[(object) 2].TextFrame.TextRange.Font.Color.RGB.ToString() != "6968388" ? "False (chọn màu)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau12(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 4].Shapes[(object) "Content Placeholder 3"].Table.Rows.Count.ToString() != "11")
        return "False (number of rows)";
      if (a.ActivePresentation.Slides[(object) 4].Shapes[(object) "Content Placeholder 3"].Table.Rows[10].Cells[1].Shape.TextFrame.TextRange.Text != "Z1")
        return "False (delete wrong row)";
      string name = a.ActivePresentation.Slides[(object) 4].Shapes[(object) "Content Placeholder 3"].Table.Style.Name;
      return name != "Light Style 1 - Accent 1" ? $"False ({name})" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau13(Application a, Presentation d)
  {
    try
    {
      string name = a.ActivePresentation.Slides[(object) 5].Shapes[(object) "Content Placeholder 3"].Table.Style.Name;
      return name != "Light Style 1 - Accent 1" ? $"False ({name})" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau14(Application a, Presentation d)
  {
    try
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Slide slide = a.ActivePresentation.Slides[(object) 3];
      // ISSUE: variable of a compiler-generated type
      Chart chart = slide.Shapes[(object) "Content Placeholder 3"].Chart;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Workbook>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Workbook), typeof (S5_Text)));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: variable of a compiler-generated type
      Workbook workbook = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, chart.ChartData.Workbook);
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (S5_Text)));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site2.Target((CallSite) S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site2, workbook.Worksheets[(object) 1]);
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target1 = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site3.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite3 = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site3;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target2 = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site4.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite4 = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site4;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site5 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site5 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj1 = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site5.Target((CallSite) S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site5, worksheet.get_Range((object) "F1", (object) "F1").Text);
      object obj2 = target2((CallSite) pSite4, obj1, "Site 5");
      if (target1((CallSite) pSite3, obj2))
        return "False (Cell F1)";
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site6 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site6 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target3 = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site6.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite6 = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site6;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site7 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site7 = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target4 = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site7.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite7 = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site7;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site8 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site8 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj3 = S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site8.Target((CallSite) S5_Text.\u003Ccau14\u003Eo__SiteContainer0.\u003C\u003Ep__Site8, worksheet.get_Range((object) "F2", (object) "F2").Text);
      object obj4 = target4((CallSite) pSite7, obj3, "46%");
      if (target3((CallSite) pSite6, obj4))
        return "False (Cell F2)";
      return chart.Legend.Width < 200.0 ? "False (Not include Site 5)" : str;
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau15(Application a, Presentation d)
  {
    try
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Slide slide = a.ActivePresentation.Slides[(object) 3];
      // ISSUE: variable of a compiler-generated type
      Chart chart = slide.Shapes[(object) "Content Placeholder 3"].Chart;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitea == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitea = CallSite<Func<CallSite, object, Workbook>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Workbook), typeof (S5_Text)));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: variable of a compiler-generated type
      Workbook workbook = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitea.Target((CallSite) S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitea, chart.ChartData.Workbook);
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Siteb == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Siteb = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (S5_Text)));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Siteb.Target((CallSite) S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Siteb, workbook.Worksheets[(object) 1]);
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitec == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitec = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target1 = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitec.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSitec = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitec;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sited == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sited = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target2 = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sited.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSited = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sited;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitee == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitee = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj1 = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitee.Target((CallSite) S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitee, worksheet.get_Range((object) "F1", (object) "F1").Text);
      object obj2 = target2((CallSite) pSited, obj1, "Site 5");
      if (target1((CallSite) pSitec, obj2))
        return "False (Cell F1)";
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitef == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitef = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target3 = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitef.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSitef = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Sitef;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Site10 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Site10 = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target4 = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Site10.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite10 = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Site10;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Site11 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Site11 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj3 = S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Site11.Target((CallSite) S5_Text.\u003Ccau15\u003Eo__SiteContainer9.\u003C\u003Ep__Site11, worksheet.get_Range((object) "F2", (object) "F2").Text);
      object obj4 = target4((CallSite) pSite10, obj3, "46%");
      if (target3((CallSite) pSitef, obj4))
        return "False (Cell F2)";
      return chart.Legend.Width < 200.0 ? "False (Not include Site 5)" : str;
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau16(Application a, Presentation d)
  {
    try
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Slide slide = a.ActivePresentation.Slides[(object) 3];
      // ISSUE: variable of a compiler-generated type
      Chart chart = slide.Shapes[(object) "Content Placeholder 3"].Chart;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site13 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site13 = CallSite<Func<CallSite, object, Workbook>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Workbook), typeof (S5_Text)));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: variable of a compiler-generated type
      Workbook workbook = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site13.Target((CallSite) S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site13, chart.ChartData.Workbook);
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site14 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site14 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (S5_Text)));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site14.Target((CallSite) S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site14, workbook.Worksheets[(object) 1]);
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site15 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site15 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target1 = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site15.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite15 = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site15;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site16 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site16 = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target2 = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site16.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite16 = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site16;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site17 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site17 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj1 = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site17.Target((CallSite) S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site17, worksheet.get_Range((object) "F1", (object) "F1").Text);
      object obj2 = target2((CallSite) pSite16, obj1, "Site 5");
      if (target1((CallSite) pSite15, obj2))
        return "False (Cell F1)";
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site18 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site18 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target3 = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site18.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite18 = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site18;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site19 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site19 = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target4 = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site19.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite19 = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site19;
      // ISSUE: reference to a compiler-generated field
      if (S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site1a == null)
      {
        // ISSUE: reference to a compiler-generated field
        S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site1a = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S5_Text), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj3 = S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site1a.Target((CallSite) S5_Text.\u003Ccau16\u003Eo__SiteContainer12.\u003C\u003Ep__Site1a, worksheet.get_Range((object) "F2", (object) "F2").Text);
      object obj4 = target4((CallSite) pSite19, obj3, "46%");
      if (target3((CallSite) pSite18, obj4))
        return "False (Cell F2)";
      return chart.Legend.Width < 200.0 ? "False (Not include Site 5)" : str;
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }
}
