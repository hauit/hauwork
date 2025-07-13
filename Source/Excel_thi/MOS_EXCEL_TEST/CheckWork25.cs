// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.CheckWork25
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using Microsoft.Office.Interop.Excel;
using System;

namespace MOS_EXCEL_TEST
{
  public static class CheckWork25
  {
    public static string Check(int cauLon, Application a, Workbook d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork25.Cau0(a, d);
        case 1:
          return CheckWork25.Cau1(a, d);
        case 2:
          return CheckWork25.Cau2(a, d);
        case 3:
          return CheckWork25.Cau3(a, d);
        case 4:
          return CheckWork25.Cau4(a, d);
        case 5:
          return CheckWork25.Cau5(a, d);
        case 6:
          return CheckWork25.Cau6(a, d);
        case 7:
          return CheckWork25.Cau7(a, d);
        case 8:
          return CheckWork25.Cau8(a, d);
        case 9:
          return CheckWork25.Cau9(a, d);
        default:
          return "Default 1";
      }
    }

    private static string Cau0(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Summer Sales"];
        if (worksheet.Shapes.Count != 6)
          return "False(chen bieu do)";
        Chart chart = worksheet.Shapes.Item((object) 6).Chart;
        if (chart.ChartType != XlChartType.xl3DPie)
          return "False(3DPie)";
        SeriesCollection seriesCollection = (SeriesCollection) chart.SeriesCollection(Type.Missing);
        if (seriesCollection.Count != 1)
          return "False(co 1 series)";
        if (seriesCollection.Item((object) 1).FormulaR1C1 != "=SERIES('Summer Sales'!R8C6,'Summer Sales'!R9C2:R14C2,'Summer Sales'!R9C6:R14C6,1)")
          return "False(add series sai)";
        if ((double) worksheet.Shapes.Item((object) 6).Left < 500.0)
          return "False(dat lai vi tri bieu do)";
        if ((double) worksheet.Shapes.Item((object) 6).Top < 200.0)
          return "False(dat lai vi tri bieu do)";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
      return "True";
    }

    private static string Cau1(Application a, Workbook d)
    {
      try
      {
        SeriesCollection seriesCollection = (SeriesCollection) ((_Worksheet) d.Worksheets[(object) "Summer Sales"]).Shapes.Item((object) "Chart 4").Chart.SeriesCollection(Type.Missing);
        if (seriesCollection.Count != 4)
          return "False(them 1 seri)";
        if (seriesCollection.Item((object) 4).FormulaR1C1 != "=SERIES('Summer Sales'!R8C6,'Summer Sales'!R12C2:R14C2,'Summer Sales'!R12C6:R14C6,4)")
          return "False(add series sai)";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
      return "True";
    }

    private static string Cau2(Application a, Workbook d)
    {
      try
      {
        Chart chart = ((_Worksheet) d.Worksheets[(object) "Spring Sales"]).Shapes.Item((object) "Chart 6").Chart;
        if (chart.ChartTitle.Text != "New Products")
          return "False(ChartTitle)";
        if (((Axis) chart.Axes((object) XlAxisType.xlValue)).AxisTitle.Text != "Gross Sales")
          return "False(Gross Sales)";
        if (((Axis) chart.Axes((object) XlAxisType.xlCategory)).AxisTitle.Text != "Months")
          return "False(Months)";
      }
      catch (Exception ex)
      {
        return "False(Add axit label)";
      }
      return "True";
    }

    private static string Cau3(Application a, Workbook d)
    {
      try
      {
        if (((_Worksheet) d.Worksheets[(object) "Spring Sales"]).Shapes.Item((object) "Chart 8").Chart.ChartColor.ToString() != "15")
          return "False";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
      return "True";
    }

    private static string Cau4(Application a, Workbook d)
    {
      try
      {
        if (((_Worksheet) d.Worksheets[(object) "Spring Sales"]).Shapes.Item((object) "Chart 6").Chart.PlotBy != XlRowCol.xlRows)
          return "False(switch Row/Column)";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
      return "True";
    }

    private static string Cau5(Application a, Workbook d)
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
      object numberFormat;
      try
      {
        numberFormat = worksheet.get_Range((object) "F3", (object) "F34").NumberFormat;
      }
      catch (Exception ex)
      {
        return "False (not apply number format)";
      }
      return numberFormat.ToString() != "0.00" ? "False (" + numberFormat.ToString() + ")" : str;
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
