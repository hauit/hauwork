// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.CheckWork19
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using Microsoft.Office.Interop.Excel;
using System;

namespace MOS_EXCEL_TEST
{
  public static class CheckWork19
  {
    public static string Check(int cauLon, Application a, Workbook d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork19.Cau0(a, d);
        case 1:
          return CheckWork19.Cau1(a, d);
        case 2:
          return CheckWork19.Cau2(a, d);
        case 3:
          return CheckWork19.Cau3(a, d);
        case 4:
          return CheckWork19.Cau4(a, d);
        case 5:
          return CheckWork19.Cau5(a, d);
        case 6:
          return CheckWork19.Cau6(a, d);
        case 7:
          return CheckWork19.Cau7(a, d);
        case 8:
          return CheckWork19.Cau8(a, d);
        case 9:
          return CheckWork19.Cau9(a, d);
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
        worksheet = (Worksheet) d.Worksheets[(object) "Customers"];
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      ListObject listObject;
      try
      {
        listObject = worksheet.ListObjects[(object) 1];
      }
      catch (Exception ex)
      {
        return "False (Table)";
      }
      if (!listObject.ShowTableStyleRowStripes)
        return "False (banded rows)";
      try
      {
      }
      catch (Exception ex)
      {
        return "False (Wrong range)";
      }
      return str;
    }

    private static string Cau1(Application a, Workbook d)
    {
      Worksheet worksheet;
      try
      {
        worksheet = (Worksheet) d.Worksheets[(object) "Customers"];
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.get_Range((object) "I499", (object) "I499").Text.ToString() != "Canada" || worksheet.get_Range((object) "I500", (object) "I500").Text.ToString() != "Canada" || worksheet.get_Range((object) "I501", (object) "I501").Text.ToString() != "Canada")
          return "False(sort sai Level 1)";
        if (worksheet.get_Range((object) "G499", (object) "G499").Text.ToString() != "QC" || worksheet.get_Range((object) "G2", (object) "G2").Text.ToString() != "AK")
          return "False(sort sai Level 2)";
        if (worksheet.get_Range((object) "h2", (object) "h2").Text.ToString() != "16202")
          return "False(sort sai Level 3)";
        if (worksheet.get_Range((object) "h4", (object) "h4").Text.ToString() != "64577")
          return "False(sort sai Level 3)";
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      return "True";
    }

    private static string Cau2(Application a, Workbook d)
    {
      try
      {
        if (((_Worksheet) d.Worksheets[(object) "Customers"]).get_Range((object) "N2", (object) "N2").Formula.ToString() != "=AVERAGE(Table1[CurrenAge])")
          return "False(=AVERAGE(Table1[CurrenAge]))";
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
        if (((_Worksheet) d.Worksheets[(object) "Products"]).get_Range((object) "D2", (object) "D32").NumberFormat.ToString() != "0.000")
          return "False(chuỗi định dạng phải la: 0.000)";
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
        worksheet = (Worksheet) d.Worksheets[(object) "Orders"];
      }
      catch (Exception ex)
      {
        return "Fales (worrksheet)";
      }
      FormatConditions formatConditions;
      try
      {
        formatConditions = worksheet.get_Range((object) "G2", (object) "G526").FormatConditions;
      }
      catch (Exception ex)
      {
        return "False (Not FormatCondition)";
      }
      if (formatConditions.Count != 1)
        return "False (Number of FormatCondition)";
      AboveAverage aboveAverage;
      try
      {
        aboveAverage = (AboveAverage) formatConditions.Item((object) 1);
        if (aboveAverage.AboveBelow != XlAboveBelow.xlAboveAverage)
          return "False(AboveAverage)";
      }
      catch (Exception ex)
      {
        return "False(AboveAverage)";
      }
      try
      {
        if (aboveAverage.Font.Color.ToString() != "24832")
          return "False(sai format)";
      }
      catch (Exception ex)
      {
        return "False(somthing wrong)";
      }
      return "True";
    }

    private static string Cau5(Application a, Workbook d)
    {
      try
      {
        if (((_Worksheet) d.Worksheets[(object) "Orders"]).get_Range((object) "J2", (object) "J2").Formula.ToString() != "=MAX(G2:G526)")
          return "False(==MAX(G2:G526)))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
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
