// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S2_Chart
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;

#nullable disable
namespace WindowsFormsApplication1;

internal class S2_Chart
{
  public static string CheckCau(int causo, Application a, Presentation d)
  {
    switch (causo)
    {
      case 1:
        return S2_Chart.cau1(a, d);
      case 2:
        return S2_Chart.cau2(a, d);
      case 3:
        return S2_Chart.cau3(a, d);
      case 4:
        return S2_Chart.cau4(a, d);
      case 5:
        return S2_Chart.cau5(a, d);
      case 6:
        return S2_Chart.cau6(a, d);
      case 7:
        return S2_Chart.cau7(a, d);
      case 8:
        return S2_Chart.cau8(a, d);
      case 9:
        return S2_Chart.cau9(a, d);
      case 10:
        return S2_Chart.cau10(a, d);
      default:
        return "case 11";
    }
  }

  private static string cau1(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 7].Shapes.Count != 3)
        return "False(add chart)";
      if (((object) a.ActivePresentation.Slides[(object) 7].Shapes[(object) 3].Type).ToString() != "msoChart")
        return "False(chen chart)";
      if (!((object) a.ActivePresentation.Slides[(object) 7].Shapes[(object) 3].Chart.ChartType).ToString().Contains("Line"))
        return "False(line chart)";
      // ISSUE: variable of a compiler-generated type
      Chart chart = a.ActivePresentation.Slides[(object) 7].Shapes[(object) 3].Chart;
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Workbook>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Workbook), typeof (S2_Chart)));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: variable of a compiler-generated type
      Workbook workbook = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, chart.ChartData.Workbook);
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (S2_Chart)));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2.Target((CallSite) S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2, workbook.Worksheets[(object) 1]);
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target1 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite3 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3;
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target2 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite4 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4;
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj1 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5.Target((CallSite) S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5, worksheet.get_Range((object) "A2", (object) "A2").Text);
      object obj2 = target2((CallSite) pSite4, obj1, "2012");
      if (target1((CallSite) pSite3, obj2))
        return "False (Cell A2)";
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site6 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site6 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target3 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site6.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite6 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site6;
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site7 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site7 = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target4 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site7.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite7 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site7;
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site8 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site8 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj3 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site8.Target((CallSite) S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site8, worksheet.get_Range((object) "B1", (object) "B1").Text);
      object obj4 = target4((CallSite) pSite7, obj3, "New Customers");
      if (target3((CallSite) pSite6, obj4))
        return "False (Cell B1)";
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site9 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site9 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target5 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site9.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite9 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site9;
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitea == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitea = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target6 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitea.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSitea = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitea;
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Siteb == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Siteb = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj5 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Siteb.Target((CallSite) S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Siteb, worksheet.get_Range((object) "B2", (object) "B2").Text);
      object obj6 = target6((CallSite) pSitea, obj5, "1700000");
      if (target5((CallSite) pSite9, obj6))
        return "False (Cell B2)";
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitec == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitec = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target7 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitec.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSitec = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitec;
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sited == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sited = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target8 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sited.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSited = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sited;
      // ISSUE: reference to a compiler-generated field
      if (S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitee == null)
      {
        // ISSUE: reference to a compiler-generated field
        S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitee = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S2_Chart), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj7 = S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitee.Target((CallSite) S2_Chart.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Sitee, worksheet.get_Range((object) "B4", (object) "B4").Text);
      object obj8 = target8((CallSite) pSited, obj7, "3200000");
      return target7((CallSite) pSitec, obj8) ? "False (Cell B4)" : "True";
    }
    catch (Exception ex)
    {
      return "False (loi khong xac dinh)";
    }
  }

  private static string cau2(Application a, Presentation d)
  {
    try
    {
      if (((object) a.ActivePresentation.Slides[(object) 4].Shapes[(object) "Content Placeholder 6"].Type).ToString() != "msoPlaceholder")
        return "False(chart da bi xoa)";
      return a.ActivePresentation.Slides[(object) 4].Shapes[(object) "Content Placeholder 6"].Chart.Legend.Position.ToString() != "xlLegendPositionTop" ? "False(Top)" : "True";
    }
    catch (Exception ex)
    {
      return "False (them lagend)";
    }
  }

  private static string cau3(Application a, Presentation d)
  {
    try
    {
      // ISSUE: variable of a compiler-generated type
      Slide slide = a.ActivePresentation.Slides[(object) 6];
      // ISSUE: variable of a compiler-generated type
      Chart chart = slide.Shapes[(object) "Chart 4"].Chart;
      return ((object) chart.ChartType).ToString() != "xl3DColumnClustered" ? "False(3DColumnClustered)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau4(Application a, Presentation d)
  {
    try
    {
      // ISSUE: variable of a compiler-generated type
      Slide slide = a.ActivePresentation.Slides[(object) 4];
      // ISSUE: variable of a compiler-generated type
      Chart chart = slide.Shapes[(object) 4].Chart;
      return chart.ChartType != XlChartType.xl3DColumnClustered ? "False (ChartType)" : "True";
    }
    catch (Exception ex)
    {
      return "False (add chart chưa đúng)";
    }
  }

  private static string cau5(Application a, Presentation d)
  {
    try
    {
      // ISSUE: variable of a compiler-generated type
      Slide slide = a.ActivePresentation.Slides[(object) 4];
      // ISSUE: variable of a compiler-generated type
      Chart chart = slide.Shapes[(object) 3].Chart;
      return chart.ChartType != XlChartType.xlBarClustered ? "False (ChartType)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau6(Application a, Presentation d)
  {
    try
    {
      // ISSUE: variable of a compiler-generated type
      Slide slide = a.ActivePresentation.Slides[(object) 5];
      // ISSUE: variable of a compiler-generated type
      Chart chart = slide.Shapes[(object) 2].Chart;
      return !chart.DataTable.ShowLegendKey ? "False (LegendKey)" : "True";
    }
    catch (Exception ex)
    {
      return "False (DataTable)";
    }
  }

  private static string cau7(Application a, Presentation d)
  {
    try
    {
      // ISSUE: variable of a compiler-generated type
      Slide slide = a.ActivePresentation.Slides[(object) 6];
      // ISSUE: variable of a compiler-generated type
      Chart chart = slide.Shapes[(object) 3].Chart;
      return chart.ChartType != XlChartType.xlLineMarkers ? "False (ChartType)" : "True";
    }
    catch (Exception ex)
    {
      return "False (add chart chưa đúng)";
    }
  }

  private static string cau8(Application a, Presentation d) => "True";

  private static string cau9(Application a, Presentation d) => "True";

  private static string cau10(Application a, Presentation d) => "True";
}
