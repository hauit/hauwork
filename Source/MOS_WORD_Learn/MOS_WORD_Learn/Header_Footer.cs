// Decompiled with JetBrains decompiler
// Type: MOS_WORD_LEARN.Header_Footer
// Assembly: MOS_WORD_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;

namespace MOS_WORD_LEARN
{
  internal class Header_Footer
  {
    public static string CheckCau(int causo, Application a, _Document d)
    {
      switch (causo)
      {
        case 1:
          return Header_Footer.cau9(a, d);
        case 2:
          return Header_Footer.cau3(a, d);
        case 3:
          return Header_Footer.cau7(a, d);
        case 4:
          return Header_Footer.cau4(a, d);
        case 5:
          return Header_Footer.cau2(a, d);
        case 6:
          return Header_Footer.cau8(a, d);
        case 7:
          return Header_Footer.cau7(a, d);
        case 8:
          return Header_Footer.cau8(a, d);
        case 9:
          return Header_Footer.cau9(a, d);
        case 10:
          return Header_Footer.cau10(a, d);
        default:
          return "False(default)";
      }
    }

    private static string cau1(Application a, _Document d)
    {
      try
      {
        if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count != 1)
          return "False(insert header)";
        object Index = (object) "Rectangle 197";
        if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[ref Index].TextFrame.TextRange.Text != "[DOCUMENT TITLE]\r")
          return "False(Banded header)";
      }
      catch (Exception ex)
      {
        return "False (Banded header)";
      }
      return "True";
    }

    private static string cau2(Application a, _Document d)
    {
      try
      {
        if (d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle != WdPageNumberStyle.wdPageNumberStyleArabic)
          return "False";
        if (!d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(" | Page\r\r"))
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau3(Application a, _Document d)
    {
      try
      {
        if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count != 1)
          return "False(insert header)";
        object Index = (object) "Rectangle 197";
        if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[ref Index].TextFrame.TextRange.Text != "[DOCUMENT TITLE]\r")
          return "False(Banded header)";
      }
      catch (Exception ex)
      {
        return "False (Banded header)";
      }
      return "True";
    }

    private static string cau4(Application a, _Document d)
    {
      try
      {
        if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text != "MOS\r\r")
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau5(Application a, _Document d)
    {
      try
      {
        if (!d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(" | Page\r\r"))
          return "False(sai kieu)";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau6(Application a, _Document d)
    {
      try
      {
        if (d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text != "\r")
          return "False(sai kieu)";
        if (d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count != 2)
          return "False(sai kieu)";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau7(Application a, _Document d)
    {
      try
      {
        if (!d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains("pg. "))
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau8(Application a, _Document d)
    {
      try
      {
        if (!d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains("\r\r"))
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau9(Application a, _Document d)
    {
      try
      {
        if (!d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains("\r\a [DOCUMENT TITLE]\a\a\r"))
          return "False";
        if (d.Sections[1].PageSetup.DifferentFirstPageHeaderFooter != -1)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau10(Application a, _Document d)
    {
      try
      {
        if (!d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(" | Page\r\r"))
          return "False(sai kieu)";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static bool cau20_2010(Application a, _Document d)
    {
      try
      {
        bool flag = false;
        string text = d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text;
        if (d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Text.Contains("| [Type the company address]") || !text.Contains("Water | [Type the company address]"))
          return false;
        foreach (Microsoft.Office.Interop.Word.Shape shape in d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
        {
          if (shape.Type == MsoShapeType.msoGroup)
            flag = shape.RelativeHorizontalPosition == WdRelativeHorizontalPosition.wdRelativeHorizontalPositionRightMarginArea;
        }
        return flag;
      }
      catch (Exception ex)
      {
        return false;
      }
    }

    private static bool cau25_2010(Application a, _Document d)
    {
      try
      {
        return (int) d.PageSetup.FooterDistance == 57;
      }
      catch (Exception ex)
      {
        return false;
      }
    }
  }
}
