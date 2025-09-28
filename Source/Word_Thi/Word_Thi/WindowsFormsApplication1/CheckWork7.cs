// Decompiled with JetBrains decompiler
// Type: MOS_WORD_TEST.CheckWork7
// Assembly: MOS_WORD_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using Microsoft.Office.Interop.Word;
using System;
using System.Reflection;

namespace MOS_WORD_TEST
{
  public static class CheckWork7
  {
    public static string Check(int cauLon, Application a, Document d)
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
        case 10:
          return CheckWork7.Cau10(a, d);
        case 11:
          return CheckWork7.Cau11(a, d);
        case 12:
          return CheckWork7.Cau12(a, d);
        case 13:
          return CheckWork7.Cau13(a, d);
        case 14:
          return CheckWork7.Cau14(a, d);
        case 15:
          return CheckWork7.Cau15(a, d);
        case 16:
          return CheckWork7.Cau16(a, d);
        case 17:
          return CheckWork7.Cau17(a, d);
        case 18:
          return CheckWork7.Cau18(a, d);
        case 19:
          return CheckWork7.Cau19(a, d);
        case 20:
          return CheckWork7.Cau20(a, d);
        case 21:
          return CheckWork7.Cau21(a, d);
        case 22:
          return CheckWork7.Cau22(a, d);
        case 23:
          return CheckWork7.Cau23(a, d);
        case 24:
          return CheckWork7.Cau24(a, d);
        case 25:
          return CheckWork7.Cau25(a, d);
        case 26:
          return CheckWork7.Cau26(a, d);
        case 27:
          return CheckWork7.Cau27(a, d);
        case 28:
          return CheckWork7.Cau28(a, d);
        case 29:
          return CheckWork7.Cau29(a, d);
        case 30:
          return CheckWork7.Cau30(a, d);
        case 31:
          return CheckWork7.Cau31(a, d);
        case 32:
          return CheckWork7.Cau32(a, d);
        case 33:
          return CheckWork7.Cau33(a, d);
        case 34:
          return CheckWork7.Cau34(a, d);
        case 35:
          return CheckWork7.Cau35(a, d);
        case 36:
          return CheckWork7.Cau36(a, d);
        case 37:
          return CheckWork7.Cau37(a, d);
        case 38:
          return CheckWork7.Cau38(a, d);
        case 39:
          return CheckWork7.Cau39(a, d);
        default:
          return "";
      }
    }

    private static string Cau0(Application a, Document d)
    {
      try
      {
        if (!d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains("\r\a [DOCUMENT TITLE]\a\a\r"))
          return "False(sai kieu)";
        if (d.Sections[1].PageSetup.DifferentFirstPageHeaderFooter != -1)
          return "False(DifferentFirstPageHeaderFooter)";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string Cau1(Application a, Document d)
    {
      try
      {
        for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.ParagraphFormat.LineSpacing.ToString() != "16.8")
            return "False(1.4)";
        }
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau2(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("The picture fits in your document") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(không chỉnh sửa văn bản)";
        if (d.Paragraphs[Index].Range.ListFormat.ListString != "5.")
          return "Falas(chuot phai vào số 1 cột 2 chọn continue)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau3(Application a, Document d)
    {
      try
      {
        if (d.InlineShapes.Count != 2)
          return "False(không thêm xóa đổi layout đối tượng)";
        if ((double) d.InlineShapes[1].SmartArt.Nodes[(object) 1].Shapes.ThreeD.BevelTopDepth != 4.0)
          return "False(bevel)";
      }
      catch (Exception ex)
      {
        return "False (Something wrong)";
      }
      return "True";
    }

    private static string Cau4(Application a, Document d)
    {
      try
      {
        object Index = (object) "Picture 11";
        if (d.Shapes[ref Index].WrapFormat.Type != WdWrapType.wdWrapSquare)
          return "False(Square)";
      }
      catch (Exception ex)
      {
        return "False (đổi Wraptext cho ảnh)";
      }
      return "True";
    }

    private static string Cau5(Application a, Document d)
    {
      try
      {
        int Index1 = 0;
        for (int Index2 = 1; Index2 < d.Paragraphs.Count; ++Index2)
        {
          if (d.Paragraphs[Index2].Range.Text.Contains("The muffin tray will still be hot"))
          {
            Index1 = Index2;
            break;
          }
        }
        if (Index1 == 0)
          return "False(khồng chỉnh sửa văn bảng)";
        if (d.Paragraphs[Index1].Range.Characters[1].Text != "(")
          return "False(chèn ký tự code 255)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau6(Application a, Document d)
    {
      string str = "True";
      object Index = (object) "Heading 1";
      return d.Styles[ref Index].Font.Bold != -1 ? "False" : str;
    }

    private static string Cau7(Application a, Document d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Paragraphs paragraphs = d.Paragraphs;
      foreach (Paragraph paragraph in paragraphs)
      {
        if (paragraph.Range.Text == "Practices\r" || paragraph.Range.Text == "Divisions\r" || paragraph.Range.Text == "Basics\r" || paragraph.Range.Text == "Web Sites\r" || paragraph.Range.Text == "Long-Term 2017 Problem Synopses\r" || paragraph.Range.Text == "Scoring Examples\r" || paragraph.Range.Text == "World Competition Locations and Venues 2016\r")
        {
          if (paragraph.OutlineLevel != WdOutlineLevel.wdOutlineLevel1)
            return "False";
        }
        else if (paragraph.Range.Text != "\f\r" && paragraph.Range.Text != "\f" && paragraph.OutlineLevel == WdOutlineLevel.wdOutlineLevel1)
          return "False";
      }
      return str;
    }

    private static string Cau8(Application a, Document d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Paragraphs paragraphs = d.Paragraphs;
      foreach (Paragraph paragraph in paragraphs)
      {
        if ((paragraph.Range.Text == "Practices\r" || paragraph.Range.Text == "Divisions\r" || paragraph.Range.Text == "Basics\r" || paragraph.Range.Text == "Web Sites\r" || paragraph.Range.Text == "Long-Term 2017 Problem Synopses\r" || paragraph.Range.Text == "Scoring Examples\r" || paragraph.Range.Text == "World Competition Locations and Venues 2016\r") && ((double) paragraph.SpaceAfter != 6.0 || (double) paragraph.SpaceBefore != 6.0 || (int) paragraph.LineSpacing != 12))
          return "False";
      }
      return str;
    }

    private static string Cau9(Application a, Document d)
    {
      string str1 = "True";
      // ISSUE: variable of a compiler-generated type
      Paragraphs paragraphs = d.Paragraphs;
      string str2 = "";
      foreach (Paragraph paragraph in paragraphs)
        str2 += paragraph.Range.Text;
      return !str2.Contains("National Odyssey of the Mind\rMichigan Odyssey of the Mind\r\r\f\rLong-Term 2017 Problem Synopses") ? "False" : str1;
    }

    private static string Cau10(Application a, Document d)
    {
      string str = "True";
      object Index = (object) "Picture 4";
      // ISSUE: variable of a compiler-generated type
      Shape shape = d.Shapes[ref Index];
      return shape.WrapFormat.Type != WdWrapType.wdWrapThrough || (double) shape.Left != 432.0 || (double) shape.Top != 108.0 ? "False" : str;
    }

    private static string Cau11(Application a, Document d)
    {
      string str = "True";
      object Index = (object) "Text Box 1";
      // ISSUE: variable of a compiler-generated type
      Shape shape = d.Shapes[ref Index];
      return shape.WrapFormat.Type != WdWrapType.wdWrapThrough || shape.TextFrame.TextRange.Text != "Figure 1-Charging Up Your Mind\r" ? "False" : str;
    }

    private static string Cau12(Application a, Document d)
    {
      string str = "True";
      object Index1 = (object) "Text Box 2";
      // ISSUE: variable of a compiler-generated type
      Shape shape = d.Shapes[ref Index1];
      if (shape.WrapFormat.Type != WdWrapType.wdWrapTopBottom || shape.RelativeHorizontalSize != WdRelativeHorizontalSize.wdRelativeHorizontalSizeMargin || shape.RelativeVerticalSize != WdRelativeVerticalSize.wdRelativeVerticalSizeMargin || (int) shape.Top != 14 || shape.RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin || shape.RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph || shape.TextFrame.TextRange.Font.Name != "Calibri" || shape.TextFrame.TextRange.Font.TextColor.RGB != -738131969 || shape.TextFrame.TextRange.Text != "\"Young people are our greatest natural resource\" - Dr. Sam, Founder of Odyssey of the Mind\r")
        return "False";
      for (int Index2 = 5; Index2 < 15; ++Index2)
      {
        if (d.Paragraphs[Index2].Range.Text == "\"Young people are our greatest natural resource\" - Dr. Sam, Founder of Odyssey of the Mind\r")
          return "False";
      }
      return str;
    }

    private static string Cau13(Application a, Document d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Hyperlink hyperlink1 = (Hyperlink) null;
      foreach (Hyperlink hyperlink2 in d.Hyperlinks)
      {
        if (hyperlink2.Name == "_Long-Term_2017_Problem")
          hyperlink1 = hyperlink2;
      }
      return hyperlink1 == null || hyperlink1.Address != null || !hyperlink1.TextToDisplay.Contains("problem solving") ? "False" : str;
    }

    private static string Cau14(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Divisions\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count)
        return "False";
      // ISSUE: variable of a compiler-generated type
      ListFormat listFormat1 = d.Paragraphs[Index + 4].Range.ListFormat;
      // ISSUE: variable of a compiler-generated type
      ListFormat listFormat2 = d.Paragraphs[Index + 8].Range.ListFormat;
      return listFormat1.ListPictureBullet.AlternativeText != "*" || listFormat2.ListPictureBullet.AlternativeText != "*" || !listFormat1.ListPictureBullet.IsPictureBullet || !listFormat2.ListPictureBullet.IsPictureBullet || (double) d.Paragraphs[Index + 4].LineSpacing != 15.0 || (double) d.Paragraphs[Index + 8].LineSpacing != 15.0 ? "False" : str;
    }

    private static string Cau15(Application a, Document d)
    {
      string str = "True";
      return d.Sections[2].PageSetup.TextColumns.Count != 2 ? "False" : str;
    }

    private static string Cau16(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Basics\r" && Index < d.Paragraphs.Count)
        ++Index;
      return Index >= d.Paragraphs.Count || !d.Paragraphs[Index + 2].Range.Text.Contains("What?\tKickoff for Odyssey of the Mind!\r") || !d.Paragraphs[Index + 3].Range.Text.Contains("Where?\tHigh School Fine Arts Center Auditorium\r") || !d.Paragraphs[Index + 4].Range.Text.Contains("When?\tWednesday, October 26th at 7:00 p.m.\r") || !d.Paragraphs[Index + 5].Range.Text.Contains("Who?\tParents, students, everyone!\r") ? "False" : str;
    }

    private static string Cau17(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Basics\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count)
        return "False";
      // ISSUE: variable of a compiler-generated type
      ListFormat listFormat1 = d.Paragraphs[Index + 2].Range.ListFormat;
      // ISSUE: variable of a compiler-generated type
      ListFormat listFormat2 = d.Paragraphs[Index + 5].Range.ListFormat;
      return listFormat1.ListPictureBullet.AlternativeText != "*" || listFormat2.ListPictureBullet.AlternativeText != "*" || !listFormat1.ListPictureBullet.IsPictureBullet || !listFormat2.ListPictureBullet.IsPictureBullet || (double) d.Paragraphs[Index + 2].LineSpacing != 12.0 || (double) d.Paragraphs[Index + 5].LineSpacing != 12.0 ? "False" : str;
    }

    private static string Cau18(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Basics\r" && Index < d.Paragraphs.Count)
        ++Index;
      return Index >= d.Paragraphs.Count || d.Paragraphs[Index + 2].Range.Font.Bold != -1 || d.Paragraphs[Index + 5].Range.Font.Bold != -1 ? "False" : str;
    }

    private static string Cau19(Application a, Document d)
    {
      string str = "True";
      bool flag = false;
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Basics\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count)
        return "False";
      foreach (TabStop tabStop in d.Paragraphs[Index + 2].TabStops)
      {
        if (tabStop.CustomTab && (double) tabStop.Position == 90.0)
          flag = true;
      }
      return !flag ? "False" : str;
    }

    private static string Cau20(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Web Sites\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count)
        return "False";
      // ISSUE: variable of a compiler-generated type
      ListFormat listFormat1 = d.Paragraphs[Index + 1].Range.ListFormat;
      // ISSUE: variable of a compiler-generated type
      ListFormat listFormat2 = d.Paragraphs[Index + 2].Range.ListFormat;
      return listFormat1.ListPictureBullet.AlternativeText != "*" || listFormat2.ListPictureBullet.AlternativeText != "*" || !listFormat1.ListPictureBullet.IsPictureBullet || !listFormat2.ListPictureBullet.IsPictureBullet || (double) d.Paragraphs[Index + 1].LineSpacing != 12.0 || (double) d.Paragraphs[Index + 2].LineSpacing != 12.0 ? "False" : str;
    }

    private static string Cau21(Application a, Document d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Hyperlink hyperlink1 = (Hyperlink) null;
      // ISSUE: variable of a compiler-generated type
      Hyperlink hyperlink2 = (Hyperlink) null;
      foreach (Hyperlink hyperlink3 in d.Hyperlinks)
      {
        if (hyperlink3.Name == "http://odysseyofthemind.com/")
          hyperlink1 = hyperlink3;
        else if (hyperlink3.Name == "http://miodyssey.com/")
          hyperlink2 = hyperlink3;
      }
      return hyperlink1 == null || !hyperlink1.TextToDisplay.Contains("National Odyssey of the Mind") || hyperlink2 == null || !hyperlink2.TextToDisplay.Contains("Michigan Odyssey of the Mind") ? "False" : str;
    }

    private static string Cau22(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Long-Term 2017 Problem Synopses\r" && Index < d.Paragraphs.Count)
        ++Index;
      return Index >= d.Paragraphs.Count || !d.Paragraphs[Index + 1].Range.Text.Contains("All problems have an 8-minute time limit") ? "False" : str;
    }

    private static string Cau23(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Problem 1: Catch Us If You Can\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count || d.Paragraphs[Index].Range.Bold != -1 || d.Paragraphs[Index].Range.Font.TextColor.RGB != -671023105)
        return "False";
      while (d.Paragraphs[Index].Range.Text != "Problem 2: Odd-a-Bot\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count || d.Paragraphs[Index].Range.Bold != -1 || d.Paragraphs[Index].Range.Font.TextColor.RGB != -671023105)
        return "False";
      while (d.Paragraphs[Index].Range.Text != "Problem 3: Classics... It's Time, OMER\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count || d.Paragraphs[Index].Range.Bold != -1 || d.Paragraphs[Index].Range.Font.TextColor.RGB != -671023105)
        return "False";
      while (!d.Paragraphs[Index].Range.Text.Contains("Problem 4: Ready, Set, Balsa, Build") && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count || d.Paragraphs[Index].Range.Bold != -1 || d.Paragraphs[Index].Range.Font.TextColor.RGB != -671023105)
        return "False";
      while (!d.Paragraphs[Index].Range.Text.Contains("Problem 5: To Be Continued: A Superhero Cliffhanger") && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count || d.Paragraphs[Index].Range.Bold != -1 || d.Paragraphs[Index].Range.Font.TextColor.RGB != -671023105)
        return "False";
      while (d.Paragraphs[Index].Range.Text != "Primary: Movin' Out! (for Grades K-2)\r" && Index < d.Paragraphs.Count)
        ++Index;
      return Index >= d.Paragraphs.Count || d.Paragraphs[Index].Range.Bold != -1 || d.Paragraphs[Index].Range.Font.TextColor.RGB != -671023105 ? "False" : str;
    }

    private static string Cau24(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Problem 1: Catch Us If You Can\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count || (double) d.Paragraphs[Index].Range.Font.Size != 11.0 || (double) d.Paragraphs[Index + 1].Range.Font.Size != 10.0)
        return "False";
      while (d.Paragraphs[Index].Range.Text != "Problem 2: Odd-a-Bot\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count || (double) d.Paragraphs[Index].Range.Font.Size != 11.0 || (double) d.Paragraphs[Index + 1].Range.Font.Size != 10.0)
        return "False";
      while (d.Paragraphs[Index].Range.Text != "Problem 3: Classics... It's Time, OMER\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count || (double) d.Paragraphs[Index].Range.Font.Size != 11.0 || (double) d.Paragraphs[Index + 1].Range.Font.Size != 10.0)
        return "False";
      while (!d.Paragraphs[Index].Range.Text.Contains("Problem 4: Ready, Set, Balsa, Build") && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count || (double) d.Paragraphs[Index].Range.Font.Size != 11.0 || (double) d.Paragraphs[Index + 1].Range.Font.Size != 10.0)
        return "False";
      while (!d.Paragraphs[Index].Range.Text.Contains("Problem 5: To Be Continued: A Superhero Cliffhanger") && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count || (double) d.Paragraphs[Index].Range.Font.Size != 11.0 || (double) d.Paragraphs[Index + 1].Range.Font.Size != 10.0)
        return "False";
      while (d.Paragraphs[Index].Range.Text != "Primary: Movin' Out! (for Grades K-2)\r" && Index < d.Paragraphs.Count)
        ++Index;
      return Index >= d.Paragraphs.Count || (double) d.Paragraphs[Index].Range.Font.Size != 11.0 || (double) d.Paragraphs[Index + 1].Range.Font.Size != 10.0 ? "False" : str;
    }

    private static string Cau25(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Scoring Examples\r" && Index < d.Paragraphs.Count)
        ++Index;
      return Index >= d.Paragraphs.Count || d.Paragraphs[Index - 1].Range.Text != "\f" || d.Sections.Count != 4 ? "False" : str;
    }

    private static string Cau26(Application a, Document d)
    {
      string str = "True";
      return d.Footnotes.Count != 1 || d.Footnotes[1].Range.Text != "Points are deducted for additional time." ? "False" : str;
    }

    private static string Cau27(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Scoring Examples\r" && Index < d.Paragraphs.Count)
        ++Index;
      return Index >= d.Paragraphs.Count || !d.Paragraphs[Index + 1].Range.Text.Contains("Michigan 2016 Region 4 Aesop Gone Viral Div") ? "False" : str;
    }

    private static string Cau28(Application a, Document d)
    {
      string str = "False";
      foreach (Microsoft.Office.Interop.Word.Table table in d.Tables)
      {
        if (table.Columns.Count == 8)
        {
          // ISSUE: reference to a compiler-generated method
          if (table.PreferredWidthType != WdPreferredWidthType.wdPreferredWidthAuto || table.Cell(1, 1).Range.Text != "Michigan 2016 Region 4 Aesop Gone Viral Div 1\r\a")
            return "False";
          str = "True";
        }
      }
      return str;
    }

    private static string Cau29(Application a, Document d)
    {
      string str = "False";
      foreach (Microsoft.Office.Interop.Word.Table table in d.Tables)
      {
        if (table.Columns.Count == 8)
        {
          // ISSUE: reference to a compiler-generated method
          if (table.PreferredWidthType != WdPreferredWidthType.wdPreferredWidthAuto || table.Cell(1, 1).Range.Text != "Michigan 2016 Region 4 Aesop Gone Viral Div 1\r\a" || table.AutoFormatType != 1 || table.Rows.Count != 10)
            return "False";
          // ISSUE: reference to a compiler-generated method
          float width = table.Cell(1, 1).Width;
          // ISSUE: reference to a compiler-generated method
          if ((int) table.Cell(1, 1).Width != 506 || table.Borders.Count != 8 || table.Borders[WdBorderType.wdBorderBottom].ColorIndex != WdColorIndex.wdGray25)
            return "False";
          str = "True";
        }
      }
      return str;
    }

    private static string Cau30(Application a, Document d)
    {
      string str = "False";
      int Index1 = 1;
      while (d.Paragraphs[Index1].Range.Text != "Scoring Examples\r" && Index1 < d.Paragraphs.Count)
        ++Index1;
      if (Index1 >= d.Paragraphs.Count || !d.Paragraphs[Index1 + 1].Range.Text.Contains("Michigan 2016 Region 4 Aesop Gone Viral Div"))
        return "False";
      for (int Index2 = Index1 + 80; Index2 < Index1 + 88; ++Index2)
      {
        if (d.Paragraphs[Index2].Range.Text.Contains("Table 1-Michigan Odyssey Web Site, Oct. 22, 2016"))
          return "True";
      }
      return str;
    }

    private static string Cau31(Application a, Document d)
    {
      string str = "True";
      if (d.Bookmarks.Count != 1)
        return "False";
      object Index = (object) "Scoring";
      return d.Bookmarks[ref Index].Range.Text != "Scoring Examples" ? "False" : str;
    }

    private static string Cau32(Application a, Document d)
    {
      string str = "True";
      return d.InlineShapes.Count != 1 || d.InlineShapes[1].SmartArt.Nodes.Count != 5 || d.InlineShapes[1].SmartArt.Nodes[(object) 1].TextFrame2.TextRange.Text != "Problem 1: No- Cycle Recycle" ? "False" : str;
    }

    private static string Cau33(Application a, Document d)
    {
      string str = "False";
      int count = d.Paragraphs.Count;
      for (int Index = count - 5; Index < count; ++Index)
      {
        if (d.Paragraphs[Index].Range.Text.Contains("Figure 2-OdysseyOfTheMind.com, Oct. 22, 2016"))
          return "True";
      }
      return str;
    }

    private static string Cau34(Application a, Document d)
    {
      string str = "True";
      return d.InlineShapes.Count != 1 || d.InlineShapes[1].SmartArt.Nodes.Count != 5 || d.InlineShapes[1].SmartArt.Nodes[(object) 1].TextFrame2.TextRange.Text != "Problem 1: No- Cycle Recycle" || d.InlineShapes[1].SmartArt.QuickStyle.Name != "Simple Fill" || d.InlineShapes[1].SmartArt.QuickStyle.Id != "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1" || (int) d.InlineShapes[1].Width != 525 || (int) d.InlineShapes[1].Height != 338 ? "False" : str;
    }

    private static string Cau35(Application a, Document d)
    {
      string str = "True";
      object documentProperties = d.BuiltInDocumentProperties;
      object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (Binder) null, documentProperties, new object[1]
      {
        (object) "Keywords"
      });
      return target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (Binder) null, target, new object[0]).ToString() != "MOS Word OM Test" ? "False" : str;
    }

    private static string Cau36(Application a, Document d)
    {
      string str = "True";
      return a.Options.SaveInterval != 4 || !d.EmbedTrueTypeFonts ? "False" : str;
    }

    private static string Cau37(Application a, Document d)
    {
      string str = "True";
      return d.ShowGrammaticalErrors ? "False" : str;
    }

    private static string Cau38(Application a, Document d)
    {
      string str = "True";
      return !a.ActiveWindow.DocumentMap ? "False" : str;
    }

    private static string Cau39(Application a, Document d) => "True";
  }
}
