// Decompiled with JetBrains decompiler
// Type: MOS_WORD_LEARN.CheckWord
// Assembly: MOS_WORD_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Reflection;

namespace MOS_WORD_LEARN
{
  internal class CheckWord
  {
    public static string CheckCau(int causo, Application a, _Document d)
    {
      switch (causo)
      {
        case 1:
          return CheckWord.cau1(a, d);
        case 2:
          return CheckWord.cau2(a, d);
        case 3:
          return CheckWord.cau3(a, d);
        case 4:
          return CheckWord.cau4(a, d);
        case 5:
          return CheckWord.cau5(a, d);
        case 6:
          return CheckWord.cau6(a, d);
        case 7:
          return CheckWord.cau7(a, d);
        case 8:
          return CheckWord.cau8(a, d);
        case 9:
          return CheckWord.cau9(a, d);
        case 10:
          return CheckWord.cau10(a, d);
        case 11:
          return CheckWord.cau11(a, d);
        case 12:
          return CheckWord.cau12(a, d);
        case 13:
          return CheckWord.cau13(a, d);
        case 14:
          return CheckWord.cau14(a, d);
        case 15:
          return CheckWord.cau15(a, d);
        case 16:
          return CheckWord.cau16(a, d);
        case 17:
          return CheckWord.cau17(a, d);
        case 18:
          return CheckWord.cau18(a, d);
        case 19:
          return CheckWord.cau19(a, d);
        case 20:
          return CheckWord.cau20(a, d);
        case 21:
          return CheckWord.cau21(a, d);
        case 22:
          return CheckWord.cau22(a, d);
        case 23:
          return CheckWord.cau23(a, d);
        case 24:
          return CheckWord.cau24(a, d);
        case 25:
          return CheckWord.cau25(a, d);
        case 26:
          return CheckWord.cau26(a, d);
        case 27:
          return CheckWord.cau27(a, d);
        case 28:
          return CheckWord.cau28(a, d);
        case 29:
          return CheckWord.cau29(a, d);
        case 30:
          return CheckWord.cau30(a, d);
        case 31:
          return CheckWord.cau31(a, d);
        case 32:
          return CheckWord.cau32(a, d);
        case 33:
          return CheckWord.cau33(a, d);
        case 34:
          return CheckWord.cau34(a, d);
        case 35:
          return CheckWord.cau35(a, d);
        case 36:
          return CheckWord.cau36(a, d);
        case 37:
          return CheckWord.cau37(a, d);
        case 38:
          return CheckWord.cau38(a, d);
        case 39:
          return CheckWord.cau39(a, d);
        case 40:
          return CheckWord.cau40(a, d);
        default:
          return "default check";
      }
    }

    private static string cau1(Application a, _Document d)
    {
      try
      {
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph = (Paragraph) null;
        int Index;
        for (Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("3200000"))
          {
            paragraph = d.Paragraphs[Index];
            break;
          }
        }
        return paragraph == null || !paragraph.Range.ListFormat.ListPictureBullet.IsPictureBullet || !d.Paragraphs[Index].Range.ListFormat.ListPictureBullet.IsPictureBullet || !d.Paragraphs[Index + 1].Range.ListFormat.ListPictureBullet.IsPictureBullet || !d.Paragraphs[Index + 2].Range.ListFormat.ListPictureBullet.IsPictureBullet || !d.Paragraphs[Index + 3].Range.ListFormat.ListPictureBullet.IsPictureBullet || !d.Paragraphs[Index + 4].Range.ListFormat.ListPictureBullet.IsPictureBullet || !d.Paragraphs[Index + 5].Range.ListFormat.ListPictureBullet.IsPictureBullet || !d.Paragraphs[Index + 6].Range.ListFormat.ListPictureBullet.IsPictureBullet || !d.Paragraphs[Index + 7].Range.ListFormat.ListPictureBullet.IsPictureBullet ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau2(Application a, _Document d)
    {
      try
      {
        object Index1 = (object) "Left Arrow 3";
        if (!d.Shapes[ref Index1].TextFrame.TextRange.Text.Contains("Salt Deficiencies"))
          return "False";
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph = (Paragraph) null;
        for (int Index2 = 1; Index2 < d.Paragraphs.Count - 3; ++Index2)
        {
          if (d.Paragraphs[Index2].Range.Text.Contains("Salt Deficiencies"))
          {
            paragraph = d.Paragraphs[Index2];
            break;
          }
        }
        return paragraph != null ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau3(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Picture 5";
        return (double) d.Shapes[ref Index].Rotation != 30.0 ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau4(Application a, _Document d)
    {
      try
      {
        if (d.Hyperlinks.Count == 0)
          return "False";
        foreach (Hyperlink hyperlink in d.Hyperlinks)
        {
          if (hyperlink.Address.Contains("mos.edu.vn") && hyperlink.TextToDisplay.Contains("click here"))
            return "True";
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau5(Application a, _Document d)
    {
      try
      {
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph = (Paragraph) null;
        int Index;
        for (Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Ha Noi"))
          {
            paragraph = d.Paragraphs[Index];
            break;
          }
        }
        return paragraph == null || paragraph.Range.ListFormat.ListLevelNumber != 2 || d.Paragraphs[Index + 1].Range.ListFormat.ListLevelNumber != 2 ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau6(Application a, _Document d)
    {
      try
      {
        foreach (OtherCorrectionsException correctionsException in a.AutoCorrect.OtherCorrectionsExceptions)
        {
          if (correctionsException.Name == "Contoso" || correctionsException.Name == "contoso")
            return "True";
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau7(Application a, _Document d)
    {
      try
      {
        int num = 1;
        foreach (ThemeFont themeFont in d.DocumentTheme.ThemeFontScheme.MajorFont)
        {
          if (num == 1 && themeFont.Name != "Century Gothic" || num == 2 && themeFont.Name != "" || num == 3 && themeFont.Name != "")
            return "False";
          ++num;
        }
        return "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau8(Application a, _Document d)
    {
      try
      {
        return Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "*.dotx").Length == 0 ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau9(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Picture 5";
        return d.Shapes[ref Index].Hyperlink.Name != "_Water" ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau10(Application a, _Document d)
    {
      try
      {
        string str = "";
        for (int index = d.Paragraphs.Count - 10; index < d.Paragraphs.Count; ++index)
          str += d.Paragraphs[index + 1].Range.Text;
        return str.Contains("Insert Bibliography Here") || !str.Contains("IIG. (2010). MOS. HCM: LTP.") ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau11(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Picture 5";
        return d.Shapes[ref Index].WrapFormat.Type != WdWrapType.wdWrapBehind ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau12(Application a, _Document d)
    {
      try
      {
        return a.AutoCorrect.ReplaceText ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau13(Application a, _Document d)
    {
      try
      {
        int Index1 = 1;
        bool flag = true;
        while (Index1 < d.Paragraphs.Count - 2)
        {
          if (d.Paragraphs[Index1].Range.Text.Contains("History"))
            flag = false;
          ++Index1;
          if (!flag)
          {
            for (int Index2 = Index1; Index2 < Index1 + 4; ++Index2)
            {
              if ((double) d.Paragraphs[Index2].Range.Font.Spacing != 1.0)
                return "False";
            }
            for (int Index3 = 1; Index3 < Index1; ++Index3)
            {
              if ((double) d.Paragraphs[Index3].Range.Font.Spacing == 1.0)
                return "False";
            }
            for (int Index4 = Index1 + 4; Index4 < Index1 + 7; ++Index4)
            {
              if ((double) d.Paragraphs[Index4].Range.Font.Spacing == 1.0)
                return "False";
            }
            return "True";
          }
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau14(Application a, _Document d)
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

    private static string cau15(Application a, _Document d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated method
        return !a.Options.get_DefaultFilePath(WdDefaultFilePath.wdAutoRecoverPath).Contains("recover") ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau16(Application a, _Document d)
    {
      try
      {
        d.ActiveWindow.View.ShowHiddenText = true;
        int Index = 1;
        bool flag = true;
        while (Index < d.Paragraphs.Count - 2)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("WATER") && d.Paragraphs[Index + 1].Range.Text.Contains("Industry"))
            flag = false;
          ++Index;
          if (!flag)
          {
            int hidden = d.Paragraphs[Index - 1].Range.Font.Hidden;
            return d.Paragraphs[Index - 1].Range.Font.Fill.Type != MsoFillType.msoFillSolid || d.Paragraphs[Index - 1].Range.Font.Fill.GradientColorType != MsoGradientColorType.msoGradientColorMixed || d.Paragraphs[Index - 1].Range.Font.Fill.ForeColor.RGB != 16777215 || d.Paragraphs[Index - 1].Range.Font.Outline != 0 || d.Paragraphs[Index - 1].Range.Font.Hidden != -1 ? "False" : "True";
          }
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau17(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        bool flag = true;
        while (Index < d.Paragraphs.Count - 2)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Water"))
            flag = false;
          ++Index;
          if (!flag)
            return d.Paragraphs[Index + 1].LineSpacingRule == WdLineSpacing.wdLineSpaceAtLeast || d.Paragraphs[Index].LineSpacingRule != WdLineSpacing.wdLineSpaceAtLeast || (double) d.Paragraphs[Index].LineSpacing != 18.0 ? "False" : "True";
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau18(Application a, _Document d)
    {
      try
      {
        return !File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "Newsletter.docx")) ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau19(Application a, _Document d)
    {
      try
      {
        int num = 0;
        foreach (Microsoft.Office.Interop.Word.Shape shape in a.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
        {
          if (shape.Type == MsoShapeType.msoTextEffect)
          {
            if (shape.TextEffect.Text != "Evaluate" || shape.TextEffect.FontName != "Verdana" || (double) shape.TextEffect.FontSize != 90.0)
              return "False";
            ++num;
          }
        }
        return num == 0 ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau20(Application a, _Document d)
    {
      try
      {
        string text = d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text;
        if (d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Text.Contains("Warter"))
          return "False";
        return !text.Contains("Water") ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau21(Application a, _Document d)
    {
      try
      {
        foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
        {
          if (shape.Type == MsoShapeType.msoTextBox && shape.Name == "Text Box 4" && shape.Fill.ForeColor.RGB != 15849926)
            return "False";
        }
        return "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau22(Application a, _Document d)
    {
      try
      {
        return a.Options.CheckSpellingAsYouType ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau23(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        bool flag = true;
        while (Index < d.Paragraphs.Count - 2)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Water"))
            flag = false;
          ++Index;
          if (!flag)
            return (double) d.Paragraphs[Index].FirstLineIndent >= 44.0 || (double) d.Paragraphs[Index].FirstLineIndent <= 43.0 ? "False" : "True";
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau24(Application a, _Document d)
    {
      try
      {
        int num1 = -1;
        int num2 = 0;
        string str = "";
        for (int index = 0; index < d.Paragraphs.Count; ++index)
          str += d.Paragraphs[index + 1].Range.Text;
        do
        {
          num2 = str.IndexOf("HoChiMinh", num2 + 1);
          ++num1;
        }
        while (num2 > 0);
        if (num1 <= 3)
          return "False";
        return str.IndexOf("Hanoi", 1) > 0 ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau25(Application a, _Document d)
    {
      try
      {
        return (int) d.PageSetup.FooterDistance != 57 ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau26(Application a, _Document d)
    {
      try
      {
        return !d.Windows.SyncScrollingSideBySide ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau27(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        bool flag = true;
        while (Index < d.Paragraphs.Count)
        {
          if (d.Paragraphs[Index].Range.Text.Length < 9 && d.Paragraphs[Index].Range.Text.Contains("WATER"))
            flag = false;
          ++Index;
          if (!flag)
            return d.Paragraphs[Index - 1].Range.Comments.Count <= 0 || d.Paragraphs[Index - 1].Range.Comments[1].Range.Text != "Verify" ? "False" : "True";
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau28(Application a, _Document d)
    {
      try
      {
        int verticalPercentScrolled = d.ActiveWindow.VerticalPercentScrolled;
        int splitVertical = d.ActiveWindow.SplitVertical;
        return verticalPercentScrolled > 0 && splitVertical > 0 ? d.ActiveWindow.Split.ToString() : "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau29(Application a, _Document d)
    {
      try
      {
        return d.ProtectionType == WdProtectionType.wdAllowOnlyComments ? "True" : "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau30(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        bool flag = true;
        while (Index < d.Paragraphs.Count)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("WATER") && d.Paragraphs[Index].Range.Text.Length < 9)
            flag = false;
          ++Index;
          if (!flag)
            return d.Paragraphs[Index].KeepWithNext == -1 || d.Paragraphs[Index].PageBreakBefore == -1 || d.Paragraphs[Index].KeepTogether != -1 ? "False" : "True";
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau31(Application a, _Document d)
    {
      try
      {
        return !File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "Water.docx")) || a.ActiveDocument.Name != "Water.docx" || a.ActiveDocument.Kind != WdDocumentKind.wdDocumentNotSpecified ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau32(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        bool flag = true;
        while (Index < d.Paragraphs.Count)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("WATER") && d.Paragraphs[Index].Range.Text.Length < 9)
            flag = false;
          ++Index;
          if (!flag)
            return d.Paragraphs[Index - 1].PageBreakBefore == -1 || d.Paragraphs[Index - 1].KeepTogether == -1 || d.Paragraphs[Index - 1].KeepWithNext != -1 ? "False" : "True";
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau33(Application a, _Document d)
    {
      try
      {
        object Index1 = (object) "IIG";
        // ISSUE: variable of a compiler-generated type
        Reviewer reviewer1 = a.ActiveWindow.View.Reviewers[ref Index1];
        if (reviewer1.Visible)
          return "False";
        object Index2 = (object) "Admin";
        // ISSUE: variable of a compiler-generated type
        Reviewer reviewer2 = a.ActiveWindow.View.Reviewers[ref Index2];
        return !reviewer2.Visible ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau34(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        bool flag = true;
        while (Index < d.Paragraphs.Count)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Chemical:Use:Type of water"))
            flag = false;
          ++Index;
          if (!flag)
            return Index + 4 > d.Paragraphs.Count || !d.Paragraphs[Index].Range.Text.Contains("Agriculture :Irrigation:Freshwater") || !d.Paragraphs[Index + 1].Range.Text.Contains("Domestic :Plumping:Fresh/Salt") || !d.Paragraphs[Index + 2].Range.Text.Contains("Manufacturing:Washing:Fresh/Salt") || !d.Paragraphs[Index + 3].Range.Text.Contains("Chemicals:Solvents:Freshwater") ? "False" : "True";
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau35(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
          return "False";
        object obj = (object) Missing.Value;
        for (int Index = 1; Index < d.Paragraphs.Count - 10; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Insert table here"))
                        return "False";
                }
        if (d.Tables[1].Columns.PreferredWidth.ToString() != "72")
                    return "False";
                if (d.Tables[1].Columns.Count != 4)
                    return "False";
                if (d.Tables[1].Rows.Count != 3)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau36(Application a, _Document d)
    {
      try
      {
        bool flag1 = false;
        bool flag2 = false;
        foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
        {
          if (shape.Type == MsoShapeType.msoPicture)
          {
            flag2 = (double) shape.LeftRelative == 50.0;
            flag1 = shape.RelativeHorizontalPosition == WdRelativeHorizontalPosition.wdRelativeHorizontalPositionInnerMarginArea;
          }
        }
        return flag1 && flag2 ? "True" : "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau37(Application a, _Document d)
    {
      try
      {
        return a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderTop].LineStyle != WdLineStyle.wdLineStyleSingle || a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderTop].LineWidth != WdLineWidth.wdLineWidth100pt || a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderTop].Color != WdColor.wdColorAutomatic || a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderLeft].LineStyle != WdLineStyle.wdLineStyleSingle || a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderLeft].LineWidth != WdLineWidth.wdLineWidth100pt || a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderLeft].Color != WdColor.wdColorAutomatic || a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderRight].LineStyle != WdLineStyle.wdLineStyleSingle || a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderRight].LineWidth != WdLineWidth.wdLineWidth100pt || a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderRight].Color != WdColor.wdColorAutomatic || a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderBottom].LineStyle != WdLineStyle.wdLineStyleSingle || a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderBottom].LineWidth != WdLineWidth.wdLineWidth100pt || a.ActiveDocument.Sections[1].Borders[WdBorderType.wdBorderBottom].Color != WdColor.wdColorAutomatic || a.ActiveDocument.Sections[1].Borders.DistanceFrom != WdBorderDistanceFrom.wdBorderDistanceFromText || a.ActiveDocument.Sections[1].Borders.Shadow || !a.ActiveDocument.Sections[1].Borders.AlwaysInFront || !a.ActiveDocument.Sections[1].Borders.SurroundHeader || !a.ActiveDocument.Sections[1].Borders.SurroundFooter || a.ActiveDocument.Sections[1].Borders.JoinBorders || a.ActiveDocument.Sections[1].Borders.DistanceFromTop != 1 || a.ActiveDocument.Sections[1].Borders.DistanceFromLeft != 4 || a.ActiveDocument.Sections[1].Borders.DistanceFromBottom != 1 || a.ActiveDocument.Sections[1].Borders.DistanceFromRight != 4 || !a.ActiveDocument.Sections[1].Borders.EnableFirstPageInSection || !a.ActiveDocument.Sections[1].Borders.EnableOtherPagesInSection ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau38(Application a, _Document d)
    {
      try
      {
        return !d.MailMerge.DataSource.Name.Contains("partnersList") || d.MailMerge.MainDocumentType != WdMailMergeMainDocType.wdFormLetters ? "False" : "True  " + CheckWord.cau39(a, d);
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau39(Application a, _Document d)
    {
      try
      {
        if (d.MailMerge.Fields.Count != 1 || d.MailMerge.Fields[1].Type != WdFieldType.wdFieldMergeField || !d.MailMerge.Fields[1].Code.Text.Contains("Email_Address"))
          return "False";
        foreach (Paragraph paragraph in d.Paragraphs)
        {
          if (paragraph.Range.Text.Contains("Insert Email item here"))
            return "False";
        }
        return "True " + CheckWord.cau40(a, d);
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau40(Application a, _Document d)
    {
      try
      {
        if (d.MailMerge.Fields.Count != 1 || d.MailMerge.Fields[1].Type != WdFieldType.wdFieldMergeField)
          return "False";
        int Index = 1;
        bool flag = true;
        while (Index < d.Paragraphs.Count - 1)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("IIG"))
            flag = false;
          ++Index;
          if (!flag)
            return !d.Paragraphs[Index].Range.Text.Contains("sale@litware.com") ? "False" : "True";
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }
  }
}
