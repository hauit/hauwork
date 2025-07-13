// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.CheckWork10
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using Microsoft.Office.Interop.Excel;
using System;
using System.Reflection;

namespace MOS_EXCEL_TEST
{
  public static class CheckWork10
  {
    public static string Check(int cauLon, Application a, Workbook d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork10.Cau0(a, d);
        case 1:
          return CheckWork10.Cau1(a, d);
        case 2:
          return CheckWork10.Cau2(a, d);
        case 3:
          return CheckWork10.Cau3(a, d);
        case 4:
          return CheckWork10.Cau4(a, d);
        case 5:
          return CheckWork10.Cau5(a, d);
        case 6:
          return CheckWork10.Cau6(a, d);
        case 7:
          return CheckWork10.Cau7(a, d);
        case 8:
          return CheckWork10.Cau8(a, d);
        case 9:
          return CheckWork10.Cau9(a, d);
        default:
          return "Default 1";
      }
    }

    private static string Cau0(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Summary"];
      }
      catch (Exception ex)
      {
        return "Fales (Not found Worksheet)";
      }
      FormatConditions formatConditions;
      try
      {
        formatConditions = worksheet.get_Range((object) "F4", (object) "F11").FormatConditions;
      }
      catch (Exception ex)
      {
        return "False (Not FormatCondition)";
      }
      if (formatConditions.Count != 1)
        return "False (Number of FormatCondition)";
      FormatCondition formatCondition;
      try
      {
        formatCondition = (FormatCondition) formatConditions.Item((object) 1);
      }
      catch (Exception ex)
      {
        return "False  (Not CellValue)";
      }
      if (formatCondition.Type != 1)
        return "False (Sai kiểu)";
      if (formatCondition.Operator != 5)
        return "False (chọn sai toán tử)";
      return formatCondition.Formula1 != "=5000000" ? "False (=5000000)" : "True";
    }

    private static string Cau1(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Region 1"];
      }
      catch (Exception ex)
      {
        return "Fales (trang tính Region 1)";
      }
      ListObject listObject;
      try
      {
        listObject = worksheet.get_Range((object) "A3", (object) "A3").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      try
      {
        if (listObject.Sort.SortFields[(object) 1].Key.get_Address((object) Missing.Value, (object) Missing.Value, External: (object) Missing.Value, RelativeTo: (object) Missing.Value) != "$A$4:$A$11")
          return "False(Product)";
        if (listObject.Sort.SortFields[(object) 1].Order != XlSortOrder.xlAscending)
          return "False(Product->A to Z)";
        if (listObject.Sort.SortFields[(object) 2].Key.get_Address((object) Missing.Value, (object) Missing.Value, External: (object) Missing.Value, RelativeTo: (object) Missing.Value) != "$F$4:$F$11")
          return "False(Total Sales)";
        if (listObject.Sort.SortFields[(object) 2].Order != XlSortOrder.xlDescending)
          return "False(Total Sales->lớn đến nhỏ)";
      }
      catch (Exception ex)
      {
        return "False(chưa sort đủ 2 trường cùng lúc)";
      }
      return "True";
    }

    private static string Cau2(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Projections"];
        if (!worksheet.get_Range((object) "C4", (object) "C4").Formula.ToString().Contains("=[@[Quarter 1]]*Q2_Increase"))
          return "False(=[@[Quarter 1]]*Q2_Increase)";
        if (!worksheet.get_Range((object) "C11", (object) "C11").Formula.ToString().Contains("=[@[Quarter 1]]*Q2_Increase"))
          return "False(=[@[Quarter 1]]*Q2_Increase)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau3(Application a, Workbook d)
    {
      try
      {
        if (!((_Worksheet) d.Worksheets[(object) "Summary"]).get_Range((object) "B15", (object) "B15").Formula.ToString().Contains("=MAX(F4:F11)"))
          return "False(=MAX(F4:F11))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau4(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Comparison"];
      }
      catch (Exception ex)
      {
        return "Fales (Comparison worksheet not found)";
      }
      if (worksheet.Shapes.Count != 1)
        return "False (không thêm xóa chart)";
      try
      {
        if (worksheet.Shapes.Item((object) 1).Chart.PlotBy != XlRowCol.xlRows)
          return "False(Biểu đồ Cột)";
      }
      catch (Exception ex)
      {
        return "False (Chart)";
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
