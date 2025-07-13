// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.CheckWork5
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using Microsoft.Office.Interop.Excel;
using System;

namespace MOS_EXCEL_TEST
{
  public static class CheckWork5
  {
    public static string Check(int cauLon, Application a, Workbook d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork5.Cau0(a, d);
        case 1:
          return CheckWork5.Cau1(a, d);
        case 2:
          return CheckWork5.Cau2(a, d);
        case 3:
          return CheckWork5.Cau3(a, d);
        case 4:
          return CheckWork5.Cau4(a, d);
        case 5:
          return CheckWork5.Cau5(a, d);
        case 6:
          return CheckWork5.Cau6(a, d);
        case 7:
          return CheckWork5.Cau7(a, d);
        case 8:
          return CheckWork5.Cau8(a, d);
        case 9:
          return CheckWork5.Cau9(a, d);
        default:
          return "Default 1";
      }
    }

    private static string Cau0(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Quote"];
        if (worksheet.get_Range((object) "E12", (object) "E12").Text.ToString() != "Quantity")
          return "False (E12)";
        if (worksheet.get_Range((object) "F12", (object) "F12").Text.ToString() != "Unit Price ")
          return "False (F12)";
        if (worksheet.get_Range((object) "F30", (object) "F30").Text.ToString() != "Total")
          return "False (F30)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau1(Application a, Workbook d)
    {
      try
      {
        string name = ((Style) ((_Worksheet) d.Worksheets[(object) "Prices"]).get_Range((object) "A1", (object) "A1").Style).Name;
        if (name != "Title")
          return "False(" + name + ")";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau2(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Orders"];
      }
      catch (Exception ex)
      {
        return "Fales (trang tính Orders)";
      }
      ListObject listObject;
      try
      {
        listObject = worksheet.get_Range((object) "A1", (object) "A1").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      try
      {
        Microsoft.Office.Interop.Excel.Filter filter = listObject.AutoFilter.Filters[1];
        if (!filter.On)
          return "False(filter tren cột 1)";
        if (filter.Criteria1.ToString() != "=Alpine Ski House")
          return "False(filter cột 1 chọn Alpine Ski House)";
      }
      catch (Exception ex)
      {
        return "False()";
      }
      return "True";
    }

    private static string Cau3(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Prices"];
        if (!worksheet.get_Range((object) "J5", (object) "J5").Formula.ToString().Contains("=[@[Unit Price]]*$L$2"))
          return "False(=[@[Unit Price]]*$L$2)";
        if (!worksheet.get_Range((object) "J25", (object) "J25").Formula.ToString().Contains("=[@[Unit Price]]*$L$2"))
          return "False(=[@[Unit Price]]*$L$2)";
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
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Prices"];
        if (!worksheet.get_Range((object) "G5", (object) "G5").Formula.ToString().Contains("=IF([@[Inventory Level]]<15%,\"Low\",\"\")"))
          return "False(=IF([@[Inventory Level]]<15%,\"Low\",\"\"))";
        if (!worksheet.get_Range((object) "G25", (object) "G25").Formula.ToString().Contains("=IF([@[Inventory Level]]<15%,\"Low\",\"\")"))
          return "False(=IF([@[Inventory Level]]<15%,\"Low\",\"\"))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau5(Application a, Workbook d)
    {
      string str = "True";
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Inventory"];
      }
      catch (Exception ex)
      {
        return "False (Inventory)";
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
        if (chart.ChartTitle.Top != 2.0)
          return "False()";
      }
      catch (Exception ex)
      {
        return "False(show Title)";
      }
      try
      {
        if (!((Series) chart.SeriesCollection((object) 1)).HasDataLabels)
          return "False(show Data)";
      }
      catch (Exception ex)
      {
        return "False()";
      }
      return str;
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
