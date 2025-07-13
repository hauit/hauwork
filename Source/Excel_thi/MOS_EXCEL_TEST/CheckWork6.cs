// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.CheckWork6
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using Microsoft.Office.Interop.Excel;
using System;
using System.Reflection;

namespace MOS_EXCEL_TEST
{
  public static class CheckWork6
  {
    public static string Check(int cauLon, Application a, Workbook d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork6.Cau0(a, d);
        case 1:
          return CheckWork6.Cau1(a, d);
        case 2:
          return CheckWork6.Cau2(a, d);
        case 3:
          return CheckWork6.Cau3(a, d);
        case 4:
          return CheckWork6.Cau4(a, d);
        case 5:
          return CheckWork6.Cau5(a, d);
        case 6:
          return CheckWork6.Cau6(a, d);
        case 7:
          return CheckWork6.Cau7(a, d);
        case 8:
          return CheckWork6.Cau8(a, d);
        case 9:
          return CheckWork6.Cau9(a, d);
        default:
          return "Default 1";
      }
    }

    private static string Cau0(Application a, Workbook d)
    {
      try
      {
        if (((_Worksheet) d.Worksheets[(object) "Last Semester"]).get_Range((object) "A3", (object) "F3").WrapText.ToString() != "True")
          return "False (A3:F3)";
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
        if (d.Names.Count != 1)
          return "False(Tạo 1 range name)";
        if (d.Names.Item((object) 1, (object) Missing.Value, (object) Missing.Value).Name != "Enrollment")
          return "False(Enrollment)";
        string str = d.Names.Item((object) 1, (object) Missing.Value, (object) Missing.Value).RefersToLocal.ToString();
        if (str != "='Enrollment Summary'!$A$3:$B$7")
          return "False(" + str + ")";
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
        worksheet = (Worksheet) d.Worksheets[(object) "Revenue"];
      }
      catch (Exception ex)
      {
        return "Fales (trang tính New Accounts)";
      }
      try
      {
        ListObject listObject = worksheet.get_Range((object) "A3", (object) "A3").ListObject;
      }
      catch (Exception ex)
      {
        return "False (A3)";
      }
      ListObject listObject1;
      try
      {
        listObject1 = worksheet.get_Range((object) "B7", (object) "B7").ListObject;
      }
      catch (Exception ex)
      {
        return "False (B7)";
      }
      try
      {
        if (listObject1.Range.get_Address((object) Missing.Value, (object) Missing.Value, External: (object) Missing.Value, RelativeTo: (object) Missing.Value) != "$A$3:$B$7")
          return "False(" + listObject1.Range.get_Address((object) Missing.Value, (object) Missing.Value, External: (object) Missing.Value, RelativeTo: (object) Missing.Value) + ")";
      }
      catch (Exception ex)
      {
        return "False(chưa chuyển sang table)";
      }
      try
      {
        if (((TableStyle) listObject1.TableStyle).Name != "TableStyleLight14")
          return "False(sai kiểu)";
      }
      catch (Exception ex)
      {
        return "False(không xác định lien quan đến Kiểu)";
      }
      return "True";
    }

    private static string Cau3(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Last Semester"];
      }
      catch (Exception ex)
      {
        return "Fales (trang tính Last Semester)";
      }
      try
      {
        ListObject listObject = worksheet.get_Range((object) "B6", (object) "B6").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      return worksheet.get_Range((object) "B6", (object) "B6").Text.ToString() != "Health & Beauty" ? "False" : "True";
    }

    private static string Cau4(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Next Semester"];
      }
      catch (Exception ex)
      {
        return "Fales (Next Semester worksheet not found)";
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
        if (seriesCollection.Count != 1)
          return "False(co 2 series)";
        if (seriesCollection.Item((object) 1).FormulaR1C1 != "=SERIES('Next Semester'!R3C5,'Next Semester'!R4C1:R21C1,'Next Semester'!R4C5:R21C5,1)")
          return "False (=SERIES('Next Semester'!R3C5,'Next Semester'!R4C1:R21C1,'Next Semester'!R4C5:R21C5,1))";
      }
      catch (Exception ex)
      {
        return "False (không xác định)";
      }
      return "True";
    }

    private static string Cau5(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Enrollment Summary"];
      }
      catch (Exception ex)
      {
        return "Fales (Enrollment Summary worksheet not found)";
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
        if (chart.ChartStyle.ToString() != "268")
          return "False (ChartStyle)";
        if (chart.ChartColor.ToString() != "19")
          return "False(Color)";
      }
      catch (Exception ex)
      {
        return "False (không xác định)";
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
