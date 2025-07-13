// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.CheckWork27
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using Microsoft.Office.Interop.Excel;
using System;

namespace MOS_EXCEL_TEST
{
  public static class CheckWork27
  {
    public static string Check(int cauLon, Application a, Workbook d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork27.Cau0(a, d);
        case 1:
          return CheckWork27.Cau1(a, d);
        case 2:
          return CheckWork27.Cau2(a, d);
        case 3:
          return CheckWork27.Cau3(a, d);
        case 4:
          return CheckWork27.Cau4(a, d);
        case 5:
          return CheckWork27.Cau5(a, d);
        case 6:
          return CheckWork27.Cau6(a, d);
        case 7:
          return CheckWork27.Cau7(a, d);
        case 8:
          return CheckWork27.Cau8(a, d);
        case 9:
          return CheckWork27.Cau9(a, d);
        default:
          return "Default 1";
      }
    }

    private static string Cau0(Application a, Workbook d)
    {
      try
      {
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "roster"];
        if (worksheet.get_Range((object) "A1", (object) "A1").EntireRow.Hidden.ToString() == "False")
          return "False(ẩn dòng 1)";
        if (worksheet.get_Range((object) "A2", (object) "A2").EntireRow.Hidden.ToString() == "False")
          return "False(ẩn dòng 2)";
        if (worksheet.get_Range((object) "A3", (object) "A3").EntireRow.Hidden.ToString() == "False")
          return "False(ẩn dòng 3)";
        if (worksheet.get_Range((object) "A4", (object) "A4").EntireRow.Hidden.ToString() == "False")
          return "False(ẩn dòng 4)";
        if (worksheet.get_Range((object) "A5", (object) "A5").EntireRow.Hidden.ToString() == "False")
          return "False(ẩn dòng 5)";
        if (worksheet.get_Range((object) "A6", (object) "A6").EntireRow.Hidden.ToString() == "False")
          return "False(ẩn dòng 6)";
        if (worksheet.get_Range((object) "A7", (object) "A7").EntireRow.Hidden.ToString() != "False")
          return "False(không ẩn dòng 7)";
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
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Student"];
        if (worksheet.get_Range((object) "D1", (object) "D1").Text.ToString() != "")
          return "False";
        if (worksheet.get_Range((object) "E1", (object) "E1").Text.ToString() == "GPA")
          return "False";
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
        if (((_Worksheet) d.Worksheets[(object) "roster"]).get_Range((object) "C8", (object) "C8").Formula.ToString() != "=PROPER(A8)")
          return "False(=PROPER(A8))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau3(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "roster"];
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.PageSetup.PrintTitleRows != "$7:$7")
          return "False(row 7)";
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
        worksheet = (Worksheet) d.Worksheets[(object) "Q2 Sales"];
      }
      catch (Exception ex)
      {
        return "Fales(Q2 Sales)";
      }
      try
      {
        if (worksheet.get_Range((object) "F6", (object) "F6").Text.ToString() != "=AVERAGE(Table2[@[April]:[June]])")
          return "False(cho hien cong thuc tren Q2 Sales)";
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
