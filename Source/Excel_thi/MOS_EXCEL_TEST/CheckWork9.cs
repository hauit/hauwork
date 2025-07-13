// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.CheckWork9
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using Microsoft.Office.Interop.Excel;
using System;
using System.Reflection;

namespace MOS_EXCEL_TEST
{
  public static class CheckWork9
  {
    public static string Check(int cauLon, Application a, Workbook d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork9.Cau0(a, d);
        case 1:
          return CheckWork9.Cau1(a, d);
        case 2:
          return CheckWork9.Cau2(a, d);
        case 3:
          return CheckWork9.Cau3(a, d);
        case 4:
          return CheckWork9.Cau4(a, d);
        case 5:
          return CheckWork9.Cau5(a, d);
        case 6:
          return CheckWork9.Cau6(a, d);
        case 7:
          return CheckWork9.Cau7(a, d);
        case 8:
          return CheckWork9.Cau8(a, d);
        case 9:
          return CheckWork9.Cau9(a, d);
        default:
          return "Default 1";
      }
    }

    private static string Cau0(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Products"];
        if (!a.ActiveWindow.FreezePanes)
          return "False(FreezePanes)";
        if (a.ActiveCell.get_Address((object) Missing.Value, (object) Missing.Value, External: (object) Missing.Value, RelativeTo: (object) Missing.Value) != "$A$3")
          return "False(để con trỏ ô A3 trước khi FreezePanes)";
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
        if (((_Worksheet) d.Worksheets[(object) "Products"]).get_Range((object) "A1", (object) "A1").HorizontalAlignment.ToString() != "-4131")
          return "False(Left)";
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
        worksheet = (Worksheet) d.Worksheets[(object) "Products"];
      }
      catch (Exception ex)
      {
        return "Fales (Not found Worksheet)";
      }
      FormatConditions formatConditions;
      try
      {
        formatConditions = worksheet.get_Range((object) "E3", (object) "E54").FormatConditions;
      }
      catch (Exception ex)
      {
        return "False (Not FormatCondition)";
      }
      if (formatConditions.Count != 1)
        return "False (Number of FormatCondition)";
      IconSetCondition conSetCondition;
      IconSet iconSet;
      try
      {
        conSetCondition = (IconSetCondition) formatConditions.Item((object) 1);
        iconSet = (IconSet) conSetCondition.IconSet;
      }
      catch (Exception ex)
      {
        return "False  (Not Iconset)";
      }
      if (iconSet.ID != XlIconSet.xl3TrafficLights1)
        return "False (loại)";
      if (iconSet.Count != 3)
        return "False (not 3 Icon)";
      if (conSetCondition.IconCriteria.Count != 3)
        return "False (not 3 Criteria)";
      if (conSetCondition.IconCriteria[(object) 1].Icon != XlIcon.xlIconRedCircleWithBorder)
        return "False (0FilledBoxes)";
      if (conSetCondition.IconCriteria[(object) 2].Icon != XlIcon.xlIconYellowCircle)
        return "False (1FilledBox)";
      if (conSetCondition.IconCriteria[(object) 3].Icon != XlIcon.xlIconGreenCircle)
        return "False (2FilledBoxes)";
      if (conSetCondition.IconCriteria[(object) 1].Type != XlConditionValueTypes.xlConditionValuePercent || conSetCondition.IconCriteria[(object) 2].Type != XlConditionValueTypes.xlConditionValuePercent || conSetCondition.IconCriteria[(object) 3].Type != XlConditionValueTypes.xlConditionValuePercent)
        return "False (type Number)";
      return conSetCondition.IconCriteria[(object) 1].Operator != 7 || conSetCondition.IconCriteria[(object) 2].Operator != 7 || conSetCondition.IconCriteria[(object) 3].Operator != 7 ? "False (Operator)" : "True";
    }

    private static string Cau3(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Products"];
      }
      catch (Exception ex)
      {
        return "Fales (trang tính Products)";
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
      try
      {
        if (((TableStyle) listObject.TableStyle).Name != "TableStyleMedium1")
          return "False(sai style)";
      }
      catch (Exception ex)
      {
        return "False(style không xac đinh)";
      }
      return "True";
    }

    private static string Cau4(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Products"];
        if (!worksheet.get_Range((object) "G3", (object) "G3").Formula.ToString().Contains("=[@[Current Value]]*Increase"))
          return "False(=[@[Current Value]]*Increase)";
        if (!worksheet.get_Range((object) "G54", (object) "G54").Formula.ToString().Contains("=[@[Current Value]]*Increase"))
          return "False(=[@[Current Value]]*Increase)";
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
        return "False (không thêm xóa chart)";
      try
      {
        if (worksheet.Shapes.Item((object) 1).Chart.ChartColor.ToString() != "11")
          return "False(sai màu)";
      }
      catch (Exception ex)
      {
        return "False (Graduation Chart)";
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
