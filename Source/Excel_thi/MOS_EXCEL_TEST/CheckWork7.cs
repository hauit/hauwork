// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.CheckWork7
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using Microsoft.Office.Interop.Excel;
using System;

namespace MOS_EXCEL_TEST
{
  public static class CheckWork7
  {
    public static string Check(int cauLon, Application a, Workbook d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork7.Cau0(a, d);
        case 1:
          return CheckWork7.Cau1(a, d);
        case 2:
          return CheckWork7.Cau2(a, d);
        case 3:
          return CheckWork7.Cau3(a, d);
        case 4:
          return CheckWork7.Cau4(a, d);
        case 5:
          return CheckWork7.Cau5(a, d);
        case 6:
          return CheckWork7.Cau6(a, d);
        case 7:
          return CheckWork7.Cau7(a, d);
        case 8:
          return CheckWork7.Cau8(a, d);
        case 9:
          return CheckWork7.Cau9(a, d);
        default:
          return "Default 1";
      }
    }

    private static string Cau0(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "New Policies"];
        if (worksheet.get_Range((object) "A5", (object) "A13").HorizontalAlignment.ToString() != "-4131")
          return "False(Left(Indent))";
        if (worksheet.get_Range((object) "A5", (object) "A13").IndentLevel.ToString() != "1")
          return "False(Indent 1)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau1(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "New Policies"];
      }
      catch (Exception ex)
      {
        return "Fales (New Policies worksheet not found)";
      }
      try
      {
        if (worksheet.get_Range((object) "J5", (object) "J13").SparklineGroups.Count != 1)
          return "False(chen Sparkline)";
        if (worksheet.get_Range((object) "J5", (object) "J13").SparklineGroups[(object) 1].SourceData != "B5:G13")
          return "False(B5:G13)";
        if (worksheet.get_Range((object) "J5", (object) "J13").SparklineGroups[(object) 1].Type != XlSparkType.xlSparkColumnStacked100)
          return "False(Win/Los)";
      }
      catch (Exception ex)
      {
        return "False (Chen Sparkline)";
      }
      return "True";
    }

    private static string Cau2(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "New Policies"];
      }
      catch (Exception ex)
      {
        return "Fales (trang tính New Policies)";
      }
      ListObject listObject;
      try
      {
        listObject = worksheet.get_Range((object) "A4", (object) "A4").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      if (!listObject.ShowTotals)
        return "False(show dồng tổng)";
      if (worksheet.get_Range((object) "b14", (object) "b14").Formula.ToString() != "=SUBTOTAL(109,[January])")
        return "False(B14)";
      if (worksheet.get_Range((object) "H14", (object) "H14").Formula.ToString() != "=SUBTOTAL(109,[Total])")
        return "False(H14)";
      if (worksheet.get_Range((object) "I14", (object) "I14").Formula.ToString() != "")
        return "False(I14)";
      return worksheet.get_Range((object) "J14", (object) "J14").Formula.ToString() != "" ? "False(J14)" : "True";
    }

    private static string Cau3(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "New Policies"];
        if (!worksheet.get_Range((object) "I5", (object) "I5").Formula.ToString().Contains("=COUNTBLANK(Table1[@[January]:[June]])"))
          return "False(=COUNTBLANK(Table1[@[January]:[June]]))";
        if (!worksheet.get_Range((object) "I13", (object) "I13").Formula.ToString().Contains("=COUNTBLANK(Table1[@[January]:[June]])"))
          return "False(=COUNTBLANK(Table1[@[January]:[June]]))";
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
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Contact"];
        if (!worksheet.get_Range((object) "C5", (object) "C5").Formula.ToString().Contains("=CONCATENATE([@[First Name]],\"@humongousinsurance.com\")"))
          return "False(=CONCATENATE([@[First Name]],\"@humongousinsurance.com\"))";
        if (!worksheet.get_Range((object) "C13", (object) "C13").Formula.ToString().Contains("=CONCATENATE([@[First Name]],\"@humongousinsurance.com\")"))
          return "False(=CONCATENATE([@[First Name]],\"@humongousinsurance.com\"))";
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
        worksheet = (Worksheet) d.Worksheets[(object) "New Policies"];
      }
      catch (Exception ex)
      {
        return "Fales (New Policies worksheet not found)";
      }
      if (worksheet.Shapes.Count < 1)
        return "False (không thêm xóa chart)";
      Chart chart;
      try
      {
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

    private static string Cau6(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Customers by Order"];
        if (worksheet.get_Range((object) "C502", (object) "C502").Text.ToString() != "")
          return "False";
        if (worksheet.get_Range((object) "B502", (object) "B502").Text.ToString() != "")
          return "False";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
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
