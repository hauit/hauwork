// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_TEST.CheckWork22
// Assembly: MOS_EXCEL_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4906C5E6-19FD-43AA-92D5-354E7B06C83B
// Assembly location: C:\Users\Admin\Downloads\WE_19\TEST_EXCEL_19\Chay File Nay.exe

using Microsoft.Office.Interop.Excel;
using System;

namespace MOS_EXCEL_TEST
{
  public static class CheckWork22
  {
    public static string Check(int cauLon, Application a, Workbook d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork22.Cau0(a, d);
        case 1:
          return CheckWork22.Cau1(a, d);
        case 2:
          return CheckWork22.Cau2(a, d);
        case 3:
          return CheckWork22.Cau3(a, d);
        case 4:
          return CheckWork22.Cau4(a, d);
        case 5:
          return CheckWork22.Cau5(a, d);
        case 6:
          return CheckWork22.Cau6(a, d);
        case 7:
          return CheckWork22.Cau7(a, d);
        case 8:
          return CheckWork22.Cau8(a, d);
        case 9:
          return CheckWork22.Cau9(a, d);
        default:
          return "Default 1";
      }
    }

    private static string Cau0(Application a, Workbook d)
    {
      try
      {
        if (((_Worksheet) d.Worksheets[(object) "Car Inventory"]).PageSetup.PrintTitleRows != "$9:$9")
          return "False($9:$9)";
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
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Car Inventory"];
        if (worksheet.get_Range((object) "C12", (object) "C12").Text.ToString() != "Truck" || worksheet.get_Range((object) "C15", (object) "C15").Text.ToString() != "Truck" || worksheet.get_Range((object) "C19", (object) "C19").Text.ToString() != "Truck")
          return "False( tim va thay the)";
        if (worksheet.get_Range((object) "C38", (object) "C38").Text.ToString() != "Truck")
          return "False( tim va thay the)";
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
        Worksheet worksheet = (Worksheet) d.Worksheets[(object) "Car Inventory"];
        if (worksheet.get_Range((object) "J10", (object) "J10").Formula.ToString() != "=95%*Inventory[@Price]")
          return "False(=95%*Inventory[@Price])";
        if (!worksheet.get_Range((object) "J10", (object) "J10").NumberFormat.ToString().Contains("$"))
          return "False(kieu tien te)";
        if (worksheet.get_Range((object) "J40", (object) "J40").Formula.ToString() != "=95%*Inventory[@Price]")
          return "False(=95%*Inventory[@Price])";
        if (!worksheet.get_Range((object) "J40", (object) "J40").NumberFormat.ToString().Contains("$"))
          return "False(kieu tien te)";
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
        if (((_Worksheet) d.Worksheets[(object) "Car Inventory"]).get_Range((object) "J9", (object) "J9").WrapText.ToString() != "True")
          return "False(WrapText)";
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
        if (((TableStyle) ((_Worksheet) d.Worksheets[(object) "Car Inventory"]).ListObjects[(object) "Inventory"].TableStyle).Name != "TableStyleLight17")
          return "False(TableStyleLight17)";
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
