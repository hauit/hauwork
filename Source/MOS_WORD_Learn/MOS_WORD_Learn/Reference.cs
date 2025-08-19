// Decompiled with JetBrains decompiler
// Type: MOS_WORD_LEARN.Reference
// Assembly: MOS_WORD_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.Office.Interop.Word;
using System;

namespace MOS_WORD_LEARN
{
  internal class Reference
  {
    public static string CheckCau(int causo, Application a, _Document d)
    {
      switch (causo)
      {
        case 1:
          return Reference.cau12(a, d);
        case 2:
          return Reference.cau15(a, d);
        case 3:
          return Reference.cau9(a, d);
        case 4:
          return Reference.cau10(a, d);
        case 5:
          return Reference.cau13(a, d);
        case 6:
          return Reference.cau11(a, d);
        case 7:
          return Reference.cau14(a, d);
        case 8:
          return Reference.cau1(a, d);
        case 9:
          return Reference.cau9(a, d);
        case 10:
          return Reference.cau10(a, d);
        case 11:
          return Reference.cau11(a, d);
        case 12:
          return Reference.cau12(a, d);
        case 13:
          return Reference.cau13(a, d);
        case 14:
          return Reference.cau14(a, d);
        case 15:
          return Reference.cau15(a, d);
        case 16:
          return Reference.cau16(a, d);
        case 17:
          return Reference.cau17(a, d);
        case 18:
          return Reference.cau18(a, d);
        case 19:
          return Reference.cau19(a, d);
        case 20:
          return Reference.cau20(a, d);
        default:
          return "default reference";
      }
    }

    private static string cau1(Application a, _Document d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (d.Bibliography.Sources[1].get_Field("Year") != "2001")
          return "False(Year)";
      }
      catch (Exception ex)
      {
        return "False(Something wrong)";
      }
      return "True";
    }

    private static string cau2(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("This Month"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False (Text was edited)";
        if (d.Paragraphs[Index + 1].Range.Text != "Contents\r")
          return "False (Auto Table 1)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau3(Application a, _Document d)
    {
      try
      {
        int Index1 = 1;
        while (Index1 < d.Paragraphs.Count && !d.Paragraphs[Index1].Range.Text.Contains("Registration Dates"))
          ++Index1;
        if (Index1 >= d.Paragraphs.Count)
          return "False (Heading was edited)";
        if (d.Paragraphs[Index1].Range.Text != "Registration Dates\u0002\r")
          return "False(wrong place)";
        for (int Index2 = 1; Index2 < d.Paragraphs.Count; ++Index2)
        {
          if (d.Paragraphs[Index2].Range.Text.Contains("IMPORTANT footnotes has not been received"))
            return "False(Cut not copy)";
        }
        if (d.Footnotes.Count != 1)
          return "False(number of footnote)";
        if (!d.Footnotes[1].Range.Text.Contains("IMPORTANT footnotes has not been received"))
          return "False(contain)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau4(Application a, _Document d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (d.Bibliography.Sources[1].get_Field("Year") != "2001")
          return "False(Year)";
      }
      catch (Exception ex)
      {
        return "False(Something wrong)";
      }
      return "True";
    }

    private static string cau5(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("This Month"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False (Text was edited)";
        if (d.Paragraphs[Index + 1].Range.Text != "Contents\r")
          return "False (Auto Table 1)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau6(Application a, _Document d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (d.Bibliography.Sources[1].get_Field("Title") != "PMMOS")
          return "False(Title)";
      }
      catch (Exception ex)
      {
        return "False(Something wrong)";
      }
      return "True";
    }

    private static string cau7(Application a, _Document d)
    {
      try
      {
        if (d.Footnotes.Count == 0)
          return "False(chen footnote)";
        if (d.Footnotes.Count > 1)
          return "False(chen du footnote)";
        if (!d.Footnotes[1].Range.Text.Contains("The table below lists the summary counts."))
          return "False(The table below lists the summary counts.)";
        if (!d.Paragraphs[2].Range.Text.Contains("Introduction\u0002\r"))
          return "False(note tieu de Introduction)";
      }
      catch (Exception ex)
      {
        return "False(Something wrong)";
      }
      return "True";
    }

    private static string cau8(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("do this without you"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False (Text was edited)";
        if (d.Paragraphs[Index + 1].Range.Text != "Contents\r")
          return "False (Auto Table 1)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau9(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Contents"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False (Text was edited)";
        if (d.Paragraphs[Index + 3].Range.Text.Contains("Summary"))
          return "False (Update entire table)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau10(Application a, _Document d)
    {
      try
      {
        if (d.Footnotes.Count != 1)
          return "False(add footnote)";
        if (!d.Footnotes[1].Range.Text.Contains("Free to join"))
          return "False(Free to join)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau11(Application a, _Document d)
    {
      try
      {
        if (d.Footnotes.Count != 2)
          return "False";
        if (d.Endnotes.Count != 0)
          return "False";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau12(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[2].Range.Text != "Contents\r")
          return "False";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau13(Application a, _Document d)
    {
      try
      {
        if (d.Footnotes.Count != 1)
          return "False(chen foodnode)";
        if (!d.Footnotes[1].Range.Text.Contains("Includes digital files."))
          return "False(Includes digital files.)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau14(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Description"))
          ++Index;
        if (Index == 0)
          return "False(Description)";
        if (!d.Paragraphs[Index + 2].Range.Text.Contains("(Manufacturing1)"))
          return "False";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau15(Application a, _Document d)
    {
      try
      {
        int num = 0;
        for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("TABLE OF CONTENTS"))
          {
            num = Index;
            break;
          }
        }
        if (num == 0)
          return "False(Table of Contents)";
        if (!d.Paragraphs[num + 2].Range.Text.ToLower().Contains("programs at other universities"))
          return "False";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau16(Application a, _Document d) => "True";

    private static string cau17(Application a, _Document d) => "True";

    private static string cau18(Application a, _Document d) => "True";

    private static string cau19(Application a, _Document d) => "True";

    private static string cau20(Application a, _Document d) => "True";

    private static string cau14_2010(Application a, _Document d)
    {
      try
      {
        if (d.TablesOfContents.Count != 1 || d.TablesOfContents[1].LowerHeadingLevel != 2 || d.TablesOfContents[1].UpperHeadingLevel != 1 || !d.TablesOfContents[1].UseHeadingStyles || !d.TablesOfContents[1].UseHyperlinks || !d.TablesOfContents[1].RightAlignPageNumbers || !d.TablesOfContents[1].IncludePageNumbers || !d.TablesOfContents[1].UseHyperlinks || !d.TablesOfContents[1].HidePageNumbersInWeb || d.TablesOfContents[1].HeadingStyles.Count != 0 || d.TablesOfContents[1].TabLeader != WdTabLeader.wdTabLeaderSpaces || d.TablesOfContents.Format != WdTocFormat.wdTOCDistinctive)
          return "False";
        for (int Index = 1; Index < 20; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Insert table of contents here"))
            return "False";
        }
        return "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }
  }
}
