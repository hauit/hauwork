// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.CheckWork3
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using Microsoft.Office.Interop.Excel;
using System;
using System.Reflection;

namespace MOS_EXCEL_TEST
{
  public static class CheckWork3
  {
    public static string Check(int cauLon, Application a, Workbook d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork3.Cau0(a, d);
        case 1:
          return CheckWork3.Cau1(a, d);
        case 2:
          return CheckWork3.Cau2(a, d);
        case 3:
          return CheckWork3.Cau3(a, d);
        case 4:
          return CheckWork3.Cau4(a, d);
        case 5:
          return CheckWork3.Cau5(a, d);
        case 6:
          return CheckWork3.Cau6(a, d);
        case 7:
          return CheckWork3.Cau7(a, d);
        case 8:
          return CheckWork3.Cau8(a, d);
        case 9:
          return CheckWork3.Cau9(a, d);
        default:
          return "Default 1";
      }
    }

    private static string Cau0(Application a, Workbook d)
    {
      string str = "True";
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Summary"];
      }
      catch (Exception ex)
      {
        return "Fales (Summary worksheet not found)";
      }
      try
      {
        if (worksheet.get_Range((object) "A2", (object) "A2").Hyperlinks.Count != 1)
          return "False (chèn hyperlink tại A2)";
        if (!worksheet.get_Range((object) "A2", (object) "A2").Hyperlinks[(object) 1].Address.Contains("www.nodpublishers.com"))
          return "False (" + worksheet.get_Range((object) "A2", (object) "A2").Hyperlinks[(object) 1].Address + ")";
        if (worksheet.get_Range((object) "A2", (object) "A2").Hyperlinks[(object) 1].ScreenTip != "Company Website")
          return "False(Company Website)";
      }
      catch (Exception ex)
      {
        return "False (Not apply hyperlink to D11)";
      }
      return str;
    }

    private static string Cau1(Application a, Workbook d)
    {
      try
      {
        if (!((_Worksheet) d.Worksheets[(object) "Historical Sales"]).Application.ActiveWindow.DisplayFormulas)
          return "False(bật show Formulas)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau2(Application a, Workbook d)
    {
      try
      {
        object documentProperties = d.BuiltinDocumentProperties;
        object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (Binder) null, documentProperties, new object[1]
        {
          (object) "Subject"
        });
        if (target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (Binder) null, target, new object[0]).ToString() != "")
          return "False";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
      return "True";
    }

    private static string Cau3(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Authors"];
        if (!worksheet.get_Range((object) "D2", (object) "D2").Formula.ToString().Contains("=IF([@[Books Sold]]>10000,500,100)"))
          return "False(=IF([@[Books Sold]]>10000,500,100))";
        if (!worksheet.get_Range((object) "D37", (object) "D37").Formula.ToString().Contains("=IF([@[Books Sold]]>10000,500,100)"))
          return "False(=IF([@[Books Sold]]>10000,500,100))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau4(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Sales"];
        if (!worksheet.get_Range((object) "E2", (object) "E2").Formula.ToString().Contains("=UPPER(LEFT([@City],3))"))
          return "False(=UPPER(LEFT([@City],3))";
        if (!worksheet.get_Range((object) "E20", (object) "E20").Formula.ToString().Contains("=UPPER(LEFT([@City],3))"))
          return "False(=UPPER(LEFT([@City],3))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau5(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Summary"];
      }
      catch (Exception ex)
      {
        return "Fales (Summary worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (Number of shape)";
      Chart chart;
      try
      {
        chart = worksheet.Shapes.Item((object) 1).Chart;
      }
      catch (Exception ex)
      {
        return "False (Not Chart)";
      }
      try
      {
        if (chart.ChartType != XlChartType.xlColumnClustered)
          return "False (not 3DPie)";
        SeriesCollection seriesCollection = (SeriesCollection) chart.SeriesCollection(Type.Missing);
        if (seriesCollection.Count != 2)
          return "False(co 2 series)";
        if (seriesCollection.Item((object) 2).FormulaR1C1 != "=SERIES(Summary!R5C3,Summary!R6C1:R12C1,Summary!R6C3:R12C3,2)")
          return "False (=SERIES(Summary!R5C3,Summary!R6C1:R12C1,Summary!R6C3:R12C3,2))";
      }
      catch (Exception ex)
      {
        return "False (không xát định)";
      }
      return "True";
    }

    private static string Cau6(Application a, Workbook d)
    {
      string str = "True";
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Olympic 100 meters"];
      }
      catch (Exception ex)
      {
        return "Fales(Olympic 100 meters worksheet not found)";
      }
      object indentLevel;
      object horizontalAlignment;
      try
      {
        indentLevel = worksheet.get_Range((object) "A2", (object) "F2").IndentLevel;
        horizontalAlignment = worksheet.get_Range((object) "A2", (object) "F2").HorizontalAlignment;
      }
      catch (Exception ex)
      {
        return "False (Not aligment)";
      }
      if (indentLevel.ToString() != "1")
        return "False(" + indentLevel.ToString() + ")";
      return horizontalAlignment.ToString() != "-4152" ? "False (not HorizontalAlignment)" : str;
    }

    private static string Cau7(Application a, Workbook d)
    {
      string str = "True";
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Olympic 100 meters"];
      }
      catch (Exception ex)
      {
        return "Fales (Olympic 100 meters worksheet not found)";
      }
      Name name;
      try
      {
        name = (Name) worksheet.get_Range((object) "E3", (object) "E34").Name;
      }
      catch (Exception ex)
      {
        return "False (Wrong range or not name)";
      }
      return name.Name != "Code" ? "False (wrong name)" : str;
    }

    private static string Cau8(Application a, Workbook d)
    {
      string str = "True";
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Olympic 100 meters"];
      }
      catch (Exception ex)
      {
        return "Fales(Olympic 100 meters worksheet not found)";
      }
      FormatConditions formatConditions;
      try
      {
        formatConditions = worksheet.get_Range((object) "F3", (object) "F34").FormatConditions;
      }
      catch (Exception ex)
      {
        return "False (not Conditions Formating)";
      }
      if (formatConditions.Count != 1)
        return "False (have " + (object) formatConditions.Count + " rules)";
      IconSetCondition conSetCondition;
      IconSet iconSet;
      try
      {
        conSetCondition = (IconSetCondition) formatConditions.Item((object) 1);
        iconSet = (IconSet) conSetCondition.IconSet;
      }
      catch (Exception ex)
      {
        return "False (not iconset)";
      }
      if (iconSet.ID != XlIconSet.xl3Flags)
        return "False (3Flags)";
      if (iconSet.Count != 3)
        return "False (have " + (object) iconSet.Count + "IconSet)";
      if (conSetCondition.IconCriteria.Count != 3)
        return "False (have " + (object) conSetCondition.IconCriteria.Count + "IconCriteria)";
      if (conSetCondition.IconCriteria[(object) 1].Icon != XlIcon.xlIconGreenFlag)
        return "False (IconGreenFlag)";
      if (conSetCondition.IconCriteria[(object) 2].Icon != XlIcon.xlIconYellowFlag)
        return "False (IconYellowFlag)";
      if (conSetCondition.IconCriteria[(object) 3].Icon != XlIcon.xlIconRedFlag)
        return "False (IconRedFlag)";
      if (conSetCondition.IconCriteria[(object) 2].Type != XlConditionValueTypes.xlConditionValueNumber)
        return "False (ConditionValueNumber2)";
      if (conSetCondition.IconCriteria[(object) 3].Type != XlConditionValueTypes.xlConditionValueNumber)
        return "False (ConditionValueNumber3)";
      if (conSetCondition.IconCriteria[(object) 2].Value.ToString() != "10")
        return "False (Value 2)";
      if (conSetCondition.IconCriteria[(object) 3].Value.ToString() != "11")
        return "False (Value 3)";
      if (conSetCondition.IconCriteria[(object) 2].Operator != 7)
        return "False (Operator 2)";
      return conSetCondition.IconCriteria[(object) 3].Operator != 7 ? "False (Operator 3)" : str;
    }

    private static string Cau9(Application a, Workbook d)
    {
      string str = "True";
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Olympic 100 meters"];
      }
      catch (Exception ex)
      {
        return "Fales(Olympic 100 meters worksheet not found)";
      }
      if (worksheet.PageSetup.PrintArea != "$A$2:$C$34")
        return "False (PrintArea=" + worksheet.PageSetup.PrintArea + ")";
      return !worksheet.PageSetup.PrintGridlines ? "False(not active PrintGridlines)" : str;
    }
  }
}
