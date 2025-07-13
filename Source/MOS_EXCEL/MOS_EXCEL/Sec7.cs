// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Sec7
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
{
  internal class Sec7
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec7.cau16(a, d);
        case 2:
          return Sec7.cau22(a, d);
        case 3:
          return Sec7.cau15(a, d);
        case 4:
          return Sec7.cau9(a, d);
        case 5:
          return Sec7.cau14(a, d);
        case 6:
          return Sec7.cau19(a, d);
        case 7:
          return Sec7.cau20(a, d);
        case 8:
          return Sec7.cau17(a, d);
        case 9:
          return Sec7.cau21(a, d);
        case 10:
          return Sec7.cau10(a, d);
        case 11:
          return Sec7.cau13(a, d);
        case 12:
          return Sec7.cau11(a, d);
        case 13:
          return Sec7.cau18(a, d);
        case 14:
          return Sec7.cau12(a, d);
        case 15:
          return Sec7.cau15(a, d);
        case 16:
          return Sec7.cau16(a, d);
        case 17:
          return Sec7.cau17(a, d);
        case 18:
          return Sec7.cau18(a, d);
        case 19:
          return Sec7.cau19(a, d);
        case 20:
          return Sec7.cau20(a, d);
        case 21:
          return Sec7.cau21(a, d);
        case 22:
          return Sec7.cau22(a, d);
        default:
          return "False Out Range";
      }
    }

    private static string cau1(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) Sec7.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, d.Worksheets[(object) "Key Applications"]);
      }
      catch (Exception ex)
      {
        return "Fales (Ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "J2", (object) "J2").SparklineGroups.Count != 1)
          return "False(chen Sparkline)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "J2", (object) "J2").SparklineGroups[(object) 1].SourceData != "B2:G2")
          return "False(B2:G2)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "J2", (object) "J2").SparklineGroups[(object) 1].Type != XlSparkType.xlSparkColumn)
          return "False(Column)";
      }
      catch (Exception ex)
      {
        return "False (something wrong)";
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
        if (Sec7.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3.Target((CallSite) Sec7.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3, d.Worksheets[(object) "Key Applications"]);
      }
      catch (Exception ex)
      {
        return "Fales (Tên Trang Tính)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (Number of shape)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 1);
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        if (chart.Legend.Width < 100.0)
          return "False (add series)";
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site4 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site4 = CallSite<Func<CallSite, object, SeriesCollection>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (SeriesCollection), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        SeriesCollection seriesCollection = Sec7.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site4.Target((CallSite) Sec7.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site4, chart.SeriesCollection(Type.Missing));
        if (seriesCollection.Count != 3)
          return "False(co 3 series)";
        // ISSUE: reference to a compiler-generated method
        if (!seriesCollection.Item((object) 3).FormulaR1C1.Contains("R32C2:R32C7"))
          return "False(add series sai)";
      }
      catch (Exception ex)
      {
        return "False(Something wrong)";
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
        if (Sec7.\u003Ccau3\u003Eo__SiteContainer5.\u003C\u003Ep__Site6 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau3\u003Eo__SiteContainer5.\u003C\u003Ep__Site6 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau3\u003Eo__SiteContainer5.\u003C\u003Ep__Site6.Target((CallSite) Sec7.\u003Ccau3\u003Eo__SiteContainer5.\u003C\u003Ep__Site6, d.Worksheets[(object) "Q1 Sales"]);
      }
      catch (Exception ex)
      {
        return "Fales (Ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "H6", (object) "H36").SparklineGroups.Count != 1)
          return "False(chen Sparkline)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "H6", (object) "H36").SparklineGroups[(object) 1].SourceData != "C6:E36")
          return "False(C6:E6)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "H6", (object) "H36").SparklineGroups[(object) 1].Type != XlSparkType.xlSparkLine)
          return "False(Column)";
      }
      catch (Exception ex)
      {
        return "False (something wrong)";
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
        if (Sec7.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site8 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site8 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site8.Target((CallSite) Sec7.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site8, d.Worksheets[(object) "Q2 Sales"]);
      }
      catch (Exception ex)
      {
        return "Fales (Tên Trang Tính)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (Number of shape)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 1);
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        if (chart.Legend.Width < 90.0)
          return "False (add series)";
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site9 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site9 = CallSite<Func<CallSite, object, SeriesCollection>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (SeriesCollection), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        SeriesCollection seriesCollection = Sec7.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site9.Target((CallSite) Sec7.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site9, chart.SeriesCollection(Type.Missing));
        if (seriesCollection.Count != 3)
          return "False(co 3 series)";
        // ISSUE: reference to a compiler-generated method
        if (!seriesCollection.Item((object) 3).FormulaR1C1.Contains("'Q2 Sales'!R6C5:R36C5,3)"))
          return "False(add series sai)";
      }
      catch (Exception ex)
      {
        return "False(Something wrong)";
      }
      return "True";
    }

    private static string cau5(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        if (d.Sheets.Count != 5)
          return "False(Duy chuyen chart qua trang tin moi)";
        if (d.Worksheets.Count != 4)
          return "False(dung Move chart, khong tao trang tinh moi roi cut qua)";
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau5\u003Eo__SiteContainera.\u003C\u003Ep__Siteb == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau5\u003Eo__SiteContainera.\u003C\u003Ep__Siteb = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau5\u003Eo__SiteContainera.\u003C\u003Ep__Siteb.Target((CallSite) Sec7.\u003Ccau5\u003Eo__SiteContainera.\u003C\u003Ep__Siteb, d.Worksheets[(object) "Outbound Calls"]);
      }
      catch (Exception ex)
      {
        return "Fales (Outbound Calls)";
      }
      try
      {
        if (worksheet.Shapes.Count != 1)
          return "False (Move not copy)";
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 1);
        // ISSUE: variable of a compiler-generated type
        Chart chart = shape.Chart;
        if (chart.ChartType == XlChartType.xlLine)
          return "False (Move line chart)";
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        object sheet = d.Sheets[(object) "Whale Tour Sales"];
      }
      catch (Exception ex)
      {
        return "False(Whale Tour Sales)";
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
        if (Sec7.\u003Ccau6\u003Eo__SiteContainerc.\u003C\u003Ep__Sited == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau6\u003Eo__SiteContainerc.\u003C\u003Ep__Sited = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau6\u003Eo__SiteContainerc.\u003C\u003Ep__Sited.Target((CallSite) Sec7.\u003Ccau6\u003Eo__SiteContainerc.\u003C\u003Ep__Sited, d.Worksheets[(object) "Inbound call"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) "Chart 1");
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        if (chart.Legend.Width < 200.0)
          return "False (Swap colunms)";
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau6\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau6\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee = CallSite<Func<CallSite, object, SeriesCollection>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (SeriesCollection), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        SeriesCollection seriesCollection = Sec7.\u003Ccau6\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee.Target((CallSite) Sec7.\u003Ccau6\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee, chart.SeriesCollection(Type.Missing));
        if (seriesCollection.Count != 6)
          return "False(co 6 series)";
        // ISSUE: reference to a compiler-generated method
        if (seriesCollection.Item((object) 3).FormulaR1C1 != "=SERIES('Inbound call'!R17C2,'Inbound call'!R2C3:R2C6,'Inbound call'!R17C3:R17C6,3)")
          return "False(Swap colunms)";
      }
      catch (Exception ex)
      {
        return "False(Something wrong)";
      }
      return "True";
    }

    private static string cau7(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau7\u003Eo__SiteContainerf.\u003C\u003Ep__Site10 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau7\u003Eo__SiteContainerf.\u003C\u003Ep__Site10 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau7\u003Eo__SiteContainerf.\u003C\u003Ep__Site10.Target((CallSite) Sec7.\u003Ccau7\u003Eo__SiteContainerf.\u003C\u003Ep__Site10, d.Worksheets[(object) "Demographics"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 1);
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (insert a chart)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau7\u003Eo__SiteContainerf.\u003C\u003Ep__Site11 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau7\u003Eo__SiteContainerf.\u003C\u003Ep__Site11 = CallSite<Func<CallSite, object, SeriesCollection>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (SeriesCollection), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        SeriesCollection seriesCollection = Sec7.\u003Ccau7\u003Eo__SiteContainerf.\u003C\u003Ep__Site11.Target((CallSite) Sec7.\u003Ccau7\u003Eo__SiteContainerf.\u003C\u003Ep__Site11, chart.SeriesCollection(Type.Missing));
        if (seriesCollection.Count != 1)
          return "False(co 1 series)";
        // ISSUE: reference to a compiler-generated method
        if (seriesCollection.Item((object) 1).FormulaR1C1 != "=SERIES(Demographics!R3C8,Demographics!R4C7:R9C7,Demographics!R4C8:R9C8,1)")
          return "False(series)";
        if (chart.ChartTitle.Text != "Donations by Age Group")
          return "False(Donations by Age Group)";
        if (chart.ChartType != XlChartType.xl3DColumnClustered)
          return "False(chartType)";
      }
      catch (Exception ex)
      {
        return "False(Something wrong)";
      }
      return "True";
    }

    private static string cau8(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau8\u003Eo__SiteContainer12.\u003C\u003Ep__Site13 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau8\u003Eo__SiteContainer12.\u003C\u003Ep__Site13 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau8\u003Eo__SiteContainer12.\u003C\u003Ep__Site13.Target((CallSite) Sec7.\u003Ccau8\u003Eo__SiteContainer12.\u003C\u003Ep__Site13, d.Worksheets[(object) "New York City"]);
      }
      catch (Exception ex)
      {
        return "Fales (New York City worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (chèn chart)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 1);
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        if (chart.ChartType != XlChartType.xlColumnClustered)
          return "False (not 3DPie)";
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau8\u003Eo__SiteContainer12.\u003C\u003Ep__Site14 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau8\u003Eo__SiteContainer12.\u003C\u003Ep__Site14 = CallSite<Func<CallSite, object, SeriesCollection>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (SeriesCollection), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        SeriesCollection seriesCollection = Sec7.\u003Ccau8\u003Eo__SiteContainer12.\u003C\u003Ep__Site14.Target((CallSite) Sec7.\u003Ccau8\u003Eo__SiteContainer12.\u003C\u003Ep__Site14, chart.SeriesCollection(Type.Missing));
        if (seriesCollection.Count != 1)
          return "False(co 1 series)";
        // ISSUE: reference to a compiler-generated method
        if (seriesCollection.Item((object) 1).FormulaR1C1 != "=SERIES('New York City'!R4C4,'New York City'!R5C2:R21C2,'New York City'!R5C4:R21C4,1)")
          return "False(series)";
      }
      catch (Exception ex)
      {
        return "False(khong xác định)";
      }
      return "True";
    }

    private static string cau9(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau9\u003Eo__SiteContainer15.\u003C\u003Ep__Site16 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau9\u003Eo__SiteContainer15.\u003C\u003Ep__Site16 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau9\u003Eo__SiteContainer15.\u003C\u003Ep__Site16.Target((CallSite) Sec7.\u003Ccau9\u003Eo__SiteContainer15.\u003C\u003Ep__Site16, d.Worksheets[(object) "London"]);
      }
      catch (Exception ex)
      {
        return "Fales (London worksheet not found)";
      }
      if (worksheet.Shapes.Count != 2)
        return "False (Number of shape)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 2);
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (not Chart)";
      }
      try
      {
        if (chart.DataTable.ShowLegendKey)
          return "False(không show LegendKey)";
      }
      catch (Exception ex)
      {
        return "Fasle(show dataTable)";
      }
      return "True";
    }

    private static string cau10(Application a, Workbook d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau10\u003Eo__SiteContainer17.\u003C\u003Ep__Site18 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau10\u003Eo__SiteContainer17.\u003C\u003Ep__Site18 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau10\u003Eo__SiteContainer17.\u003C\u003Ep__Site18.Target((CallSite) Sec7.\u003Ccau10\u003Eo__SiteContainer17.\u003C\u003Ep__Site18, d.Worksheets[(object) "New Accounts"]);
      }
      catch (Exception ex)
      {
        return "Fales (New Accounts worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (Number of shape)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 1);
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        if (chart.PlotBy == XlRowCol.xlRows)
          return "False (biểu đồ dòng)";
      }
      catch (Exception ex)
      {
        return "False (not add title)";
      }
      return str;
    }

    private static string cau11(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau11\u003Eo__SiteContainer19.\u003C\u003Ep__Site1a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau11\u003Eo__SiteContainer19.\u003C\u003Ep__Site1a = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau11\u003Eo__SiteContainer19.\u003C\u003Ep__Site1a.Target((CallSite) Sec7.\u003Ccau11\u003Eo__SiteContainer19.\u003C\u003Ep__Site1a, d.Worksheets[(object) "Summary"]);
      }
      catch (Exception ex)
      {
        return "Fales (Summary worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (Number of shape)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 1);
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        if (chart.ChartType != XlChartType.xlColumnClustered)
          return "False (not 3DPie)";
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau11\u003Eo__SiteContainer19.\u003C\u003Ep__Site1b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau11\u003Eo__SiteContainer19.\u003C\u003Ep__Site1b = CallSite<Func<CallSite, object, SeriesCollection>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (SeriesCollection), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        SeriesCollection seriesCollection = Sec7.\u003Ccau11\u003Eo__SiteContainer19.\u003C\u003Ep__Site1b.Target((CallSite) Sec7.\u003Ccau11\u003Eo__SiteContainer19.\u003C\u003Ep__Site1b, chart.SeriesCollection(Type.Missing));
        if (seriesCollection.Count != 2)
          return "False(co 2 series)";
        // ISSUE: reference to a compiler-generated method
        if (seriesCollection.Item((object) 2).FormulaR1C1 != "=SERIES(Summary!R5C3,Summary!R6C1:R12C1,Summary!R6C3:R12C3,2)")
          return "False (=SERIES(Summary!R5C3,Summary!R6C1:R12C1,Summary!R6C3:R12C3,2))";
      }
      catch (Exception ex)
      {
        return "False (không xát định)";
      }
      return "True";
    }

    private static string cau12(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau12\u003Eo__SiteContainer1c.\u003C\u003Ep__Site1d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau12\u003Eo__SiteContainer1c.\u003C\u003Ep__Site1d = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau12\u003Eo__SiteContainer1c.\u003C\u003Ep__Site1d.Target((CallSite) Sec7.\u003Ccau12\u003Eo__SiteContainer1c.\u003C\u003Ep__Site1d, d.Worksheets[(object) "Enrollment"]);
      }
      catch (Exception ex)
      {
        return "Fales (Enrollment worksheet not found)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "G5", (object) "G25").SparklineGroups.Count != 1)
          return "False(chen Sparkline)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "G5", (object) "G25").SparklineGroups[(object) 1].SourceData != "D5:F25")
          return "False(D5:F25)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "G5", (object) "G25").SparklineGroups[(object) 1].Type != XlSparkType.xlSparkColumn)
          return "False(Column)";
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
      }
      catch (Exception ex)
      {
        return "False (không xát định)";
      }
      return "True";
    }

    private static string cau13(Application a, Workbook d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site1f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site1f = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site1f.Target((CallSite) Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site1f, d.Worksheets[(object) "Graduation"]);
      }
      catch (Exception ex)
      {
        return "Fales (Graduation worksheet not found)";
      }
      if (worksheet.Shapes.Count != 0)
        return "False (Move Chart)";
      try
      {
        if (d.Worksheets.Count != 5)
          return "False(dùng chức năng move chart)";
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site20 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site20 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site20.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite20 = Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site20;
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site21 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site21 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec7), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        object obj = Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site21.Target((CallSite) Sec7.\u003Ccau13\u003Eo__SiteContainer1e.\u003C\u003Ep__Site21, d.Sheets[(object) "Graduation Chart"]);
        str = target((CallSite) pSite20, obj);
      }
      catch (Exception ex)
      {
        return "False (Graduation Chart)";
      }
      return "True";
    }

    private static string cau14(Application a, Workbook d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau14\u003Eo__SiteContainer22.\u003C\u003Ep__Site23 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau14\u003Eo__SiteContainer22.\u003C\u003Ep__Site23 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau14\u003Eo__SiteContainer22.\u003C\u003Ep__Site23.Target((CallSite) Sec7.\u003Ccau14\u003Eo__SiteContainer22.\u003C\u003Ep__Site23, d.Worksheets[(object) "Instructional Hours"]);
      }
      catch (Exception ex)
      {
        return "Fales (Instructional Hours worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (Number of shape)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 1);
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau14\u003Eo__SiteContainer22.\u003C\u003Ep__Site24 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau14\u003Eo__SiteContainer22.\u003C\u003Ep__Site24 = CallSite<Func<CallSite, object, Axis>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Axis), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Axis axis = Sec7.\u003Ccau14\u003Eo__SiteContainer22.\u003C\u003Ep__Site24.Target((CallSite) Sec7.\u003Ccau14\u003Eo__SiteContainer22.\u003C\u003Ep__Site24, chart.Axes((object) XlAxisType.xlValue));
        if (axis.AxisTitle.Text != "Hours")
          return "False(Hours)";
      }
      catch (Exception ex)
      {
        return "False (not add title)";
      }
      return str;
    }

    private static string cau15(Application a, Workbook d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau15\u003Eo__SiteContainer25.\u003C\u003Ep__Site26 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau15\u003Eo__SiteContainer25.\u003C\u003Ep__Site26 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau15\u003Eo__SiteContainer25.\u003C\u003Ep__Site26.Target((CallSite) Sec7.\u003Ccau15\u003Eo__SiteContainer25.\u003C\u003Ep__Site26, d.Worksheets[(object) "Inventory"]);
      }
      catch (Exception ex)
      {
        return "False (Inventory)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (Number of shape)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 1);
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        if (chart.ChartTitle.Top != 2.0)
          return "False()";
      }
      catch (Exception ex)
      {
        return "False(show Title)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau15\u003Eo__SiteContainer25.\u003C\u003Ep__Site27 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau15\u003Eo__SiteContainer25.\u003C\u003Ep__Site27 = CallSite<Func<CallSite, object, Series>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Series), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Series series = Sec7.\u003Ccau15\u003Eo__SiteContainer25.\u003C\u003Ep__Site27.Target((CallSite) Sec7.\u003Ccau15\u003Eo__SiteContainer25.\u003C\u003Ep__Site27, chart.SeriesCollection((object) 1));
        if (!series.HasDataLabels)
          return "False(show Data)";
      }
      catch (Exception ex)
      {
        return "False()";
      }
      return str;
    }

    private static string cau16(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau16\u003Eo__SiteContainer28.\u003C\u003Ep__Site29 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau16\u003Eo__SiteContainer28.\u003C\u003Ep__Site29 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau16\u003Eo__SiteContainer28.\u003C\u003Ep__Site29.Target((CallSite) Sec7.\u003Ccau16\u003Eo__SiteContainer28.\u003C\u003Ep__Site29, d.Worksheets[(object) "Next Semester"]);
      }
      catch (Exception ex)
      {
        return "Fales (Next Semester worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (Number of shape)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 1);
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        if (chart.ChartType != XlChartType.xlColumnClustered)
          return "False (not 3DPie)";
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau16\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau16\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a = CallSite<Func<CallSite, object, SeriesCollection>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (SeriesCollection), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        SeriesCollection seriesCollection = Sec7.\u003Ccau16\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a.Target((CallSite) Sec7.\u003Ccau16\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a, chart.SeriesCollection(Type.Missing));
        if (seriesCollection.Count != 1)
          return "False(co 2 series)";
        // ISSUE: reference to a compiler-generated method
        if (seriesCollection.Item((object) 1).FormulaR1C1 != "=SERIES('Next Semester'!R3C5,'Next Semester'!R4C1:R21C1,'Next Semester'!R4C5:R21C5,1)")
          return "False (=SERIES('Next Semester'!R3C5,'Next Semester'!R4C1:R21C1,'Next Semester'!R4C5:R21C5,1))";
      }
      catch (Exception ex)
      {
        return "False (không xác định)";
      }
      return "True";
    }

    private static string cau17(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2c = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2c.Target((CallSite) Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2c, d.Worksheets[(object) "Enrollment Summary"]);
      }
      catch (Exception ex)
      {
        return "Fales (Enrollment Summary worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (Number of shape)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Shape shape = worksheet.Shapes.Item((object) 1);
        chart = shape.Chart;
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2d = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec7), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2d.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite2d = Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2d;
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2e = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec7), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite2e = Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2e;
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2f = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec7), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        object obj1 = Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2f.Target((CallSite) Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2f, chart.ChartStyle);
        object obj2 = target2((CallSite) pSite2e, obj1, "268");
        if (target1((CallSite) pSite2d, obj2))
          return "False (ChartStyle)";
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site30 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site30 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec7), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site30.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite30 = Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site30;
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site31 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site31 = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec7), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site31.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite31 = Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site31;
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site32 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site32 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec7), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        object obj3 = Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site32.Target((CallSite) Sec7.\u003Ccau17\u003Eo__SiteContainer2b.\u003C\u003Ep__Site32, chart.ChartColor);
        object obj4 = target4((CallSite) pSite31, obj3, "19");
        if (target3((CallSite) pSite30, obj4))
          return "False(Color)";
      }
      catch (Exception ex)
      {
        return "False (không xác định)";
      }
      return "True";
    }

    private static string cau18(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau18\u003Eo__SiteContainer33.\u003C\u003Ep__Site34 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau18\u003Eo__SiteContainer33.\u003C\u003Ep__Site34 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau18\u003Eo__SiteContainer33.\u003C\u003Ep__Site34.Target((CallSite) Sec7.\u003Ccau18\u003Eo__SiteContainer33.\u003C\u003Ep__Site34, d.Worksheets[(object) "New Policies"]);
      }
      catch (Exception ex)
      {
        return "Fales (New Policies worksheet not found)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "J5", (object) "J13").SparklineGroups.Count != 1)
          return "False(chen Sparkline)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "J5", (object) "J13").SparklineGroups[(object) 1].SourceData != "B5:G13")
          return "False(B5:G13)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet.get_Range((object) "J5", (object) "J13").SparklineGroups[(object) 1].Type != XlSparkType.xlSparkColumnStacked100)
          return "False(Win/Los)";
      }
      catch (Exception ex)
      {
        return "False (Chen Sparkline)";
      }
      return "True";
    }

    private static string cau19(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau19\u003Eo__SiteContainer35.\u003C\u003Ep__Site36 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau19\u003Eo__SiteContainer35.\u003C\u003Ep__Site36 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau19\u003Eo__SiteContainer35.\u003C\u003Ep__Site36.Target((CallSite) Sec7.\u003Ccau19\u003Eo__SiteContainer35.\u003C\u003Ep__Site36, d.Worksheets[(object) "New Policies"]);
      }
      catch (Exception ex)
      {
        return "Fales (New Policies worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (không thêm xóa chart)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        chart = worksheet.Shapes.Item((object) 1).Chart;
      }
      catch (Exception ex)
      {
        return "False (not Chart)";
      }
      try
      {
        if (chart.DataTable.ShowLegendKey)
          return "False(Layout)";
      }
      catch (Exception ex1)
      {
        try
        {
          if (chart.Legend.Position != XlLegendPosition.xlLegendPositionBottom)
            return "False(sai layout)";
        }
        catch (Exception ex2)
        {
          return "False(Chon Layout)";
        }
        return "True";
      }
      return "False(Layout)";
    }

    private static string cau20(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site38 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site38 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site38.Target((CallSite) Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site38, d.Worksheets[(object) "Summary"]);
      }
      catch (Exception ex)
      {
        return "Fales (Summary worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (không thêm xóa chart)";
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Chart chart = worksheet.Shapes.Item((object) 1).Chart;
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site39 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site39 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec7), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site39.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite39 = Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site39;
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site3a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site3a = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec7), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site3a.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite3a = Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site3a;
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site3b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site3b = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec7), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        object obj1 = Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site3b.Target((CallSite) Sec7.\u003Ccau20\u003Eo__SiteContainer37.\u003C\u003Ep__Site3b, chart.ChartColor);
        object obj2 = target2((CallSite) pSite3a, obj1, "11");
        if (target1((CallSite) pSite39, obj2))
          return "False(sai màu)";
      }
      catch (Exception ex)
      {
        return "False (Graduation Chart)";
      }
      return "True";
    }

    private static string cau21(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau21\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau21\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3d = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau21\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3d.Target((CallSite) Sec7.\u003Ccau21\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3d, d.Worksheets[(object) "Comparison"]);
      }
      catch (Exception ex)
      {
        return "Fales (Comparison worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (không thêm xóa chart)";
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Chart chart = worksheet.Shapes.Item((object) 1).Chart;
        if (chart.PlotBy != XlRowCol.xlRows)
          return "False(Biểu đồ Cột)";
      }
      catch (Exception ex)
      {
        return "False (Chart)";
      }
      return "True";
    }

    private static string cau22(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau22\u003Eo__SiteContainer3e.\u003C\u003Ep__Site3f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau22\u003Eo__SiteContainer3e.\u003C\u003Ep__Site3f = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec7.\u003Ccau22\u003Eo__SiteContainer3e.\u003C\u003Ep__Site3f.Target((CallSite) Sec7.\u003Ccau22\u003Eo__SiteContainer3e.\u003C\u003Ep__Site3f, d.Worksheets[(object) "Score Distribution"]);
      }
      catch (Exception ex)
      {
        return "Fales (Score Distribution worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (không thêm xóa chart)";
      // ISSUE: variable of a compiler-generated type
      Chart chart;
      try
      {
        // ISSUE: reference to a compiler-generated method
        chart = worksheet.Shapes.Item((object) 1).Chart;
      }
      catch (Exception ex)
      {
        return "False (Chart)";
      }
      try
      {
        if (chart.Legend != null)
          return "False(Legend)";
      }
      catch (Exception ex)
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec7.\u003Ccau22\u003Eo__SiteContainer3e.\u003C\u003Ep__Site40 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec7.\u003Ccau22\u003Eo__SiteContainer3e.\u003C\u003Ep__Site40 = CallSite<Func<CallSite, object, Series>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Series), typeof (Sec7)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Series series = Sec7.\u003Ccau22\u003Eo__SiteContainer3e.\u003C\u003Ep__Site40.Target((CallSite) Sec7.\u003Ccau22\u003Eo__SiteContainer3e.\u003C\u003Ep__Site40, chart.SeriesCollection((object) 1));
        if (!series.HasDataLabels)
          return "False(show Data)";
      }
      return "True";
    }
  }
}
