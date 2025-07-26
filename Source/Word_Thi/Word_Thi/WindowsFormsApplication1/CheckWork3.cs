// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.CheckWork3
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Word;
using System;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
{
  public static class CheckWork3
  {
    public static string Check(int cauLon, Application a, Document d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork3.Cau0(a, d);
        case 1:
          return CheckWork3.Cau1(a, d);
        case 2:
          return CheckWork3.Cau2(a, d);
        case 3:
          return CheckWork3.Cau3(a, d);
        case 4:
          return CheckWork3.Cau4(a, d);
        case 5:
          return CheckWork3.Cau5(a, d);
        case 6:
          return CheckWork3.Cau6(a, d);
        case 7:
          return CheckWork3.Cau7(a, d);
        case 8:
          return CheckWork3.Cau8(a, d);
        case 9:
          return CheckWork3.Cau9(a, d);
        case 10:
          return CheckWork3.Cau10(a, d);
        case 11:
          return CheckWork3.Cau11(a, d);
        case 12:
          return CheckWork3.Cau12(a, d);
        case 13:
          return CheckWork3.Cau13(a, d);
        case 14:
          return CheckWork3.Cau14(a, d);
        case 15:
          return CheckWork3.Cau15(a, d);
        case 16:
          return CheckWork3.Cau16(a, d);
        case 17:
          return CheckWork3.Cau17(a, d);
        case 18:
          return CheckWork3.Cau18(a, d);
        case 19:
          return CheckWork3.Cau19(a, d);
        case 20:
          return CheckWork3.Cau20(a, d);
        case 21:
          return CheckWork3.Cau21(a, d);
        case 22:
          return CheckWork3.Cau22(a, d);
        case 23:
          return CheckWork3.Cau23(a, d);
        case 24:
          return CheckWork3.Cau24(a, d);
        case 25:
          return CheckWork3.Cau25(a, d);
        case 26:
          return CheckWork3.Cau26(a, d);
        case 27:
          return CheckWork3.Cau27(a, d);
        case 28:
          return CheckWork3.Cau28(a, d);
        case 29:
          return CheckWork3.Cau29(a, d);
        default:
          return "";
      }
    }

    private static string Cau0(Application a, Document d)
    {
      try
      {
        object Index = (object) "Heading 1";
        if (d.Styles[ref Index].Borders[WdBorderType.wdBorderBottom].Color.ToString() != "-721354753")
          return "False(sai kiểu)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau1(Application a, Document d)
    {
      try
      {
        int num = 0;
        for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Where to find us"))
          {
            num = Index;
            break;
          }
        }
        if (num == 0)
          return "False(không tìm thấy Where to find us)";
        if (d.Paragraphs[num - 1].Range.Text != "\f")
          return "False(chen ngắt section)";
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
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Genre"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(không thay đổi nọi dung)";
        // ISSUE: reference to a compiler-generated field
        if (CheckWork3.\u003CCau2\u003Eo__SiteContainer1.\u003C\u003Ep__Site2 == null)
        {
          // ISSUE: reference to a compiler-generated field
          CheckWork3.\u003CCau2\u003Eo__SiteContainer1.\u003C\u003Ep__Site2 = CallSite<Func<CallSite, object, Style>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Style), typeof (CheckWork3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Style style1 = CheckWork3.\u003CCau2\u003Eo__SiteContainer1.\u003C\u003Ep__Site2.Target((CallSite) CheckWork3.\u003CCau2\u003Eo__SiteContainer1.\u003C\u003Ep__Site2, d.Paragraphs[Index].Range.CharacterStyle);
        if (style1.NameLocal != "Subtle Emphasis")
          return "False (Subtle Emphasis)";
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("External Parther"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(không thay đổi nọi dung)";
        // ISSUE: reference to a compiler-generated field
        if (CheckWork3.\u003CCau2\u003Eo__SiteContainer1.\u003C\u003Ep__Site3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          CheckWork3.\u003CCau2\u003Eo__SiteContainer1.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, Style>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Style), typeof (CheckWork3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Style style2 = CheckWork3.\u003CCau2\u003Eo__SiteContainer1.\u003C\u003Ep__Site3.Target((CallSite) CheckWork3.\u003CCau2\u003Eo__SiteContainer1.\u003C\u003Ep__Site3, d.Paragraphs[Index].Range.CharacterStyle);
        if (style2.NameLocal != "Subtle Emphasis")
          return "False (Subtle Emphasis)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau3(Application a, Document d)
    {
      int num = 0;
      for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
      {
        if (d.Paragraphs[Index].Range.Text.Contains("Where to find us"))
        {
          num = Index;
          break;
        }
      }
      if (num == 0)
        return "False(không tìm thấy Where to find us)";
      return d.Paragraphs[num + 1].Range.InlineShapes.Count != 1 ? "False(chèn anh vào đoạn sau tiêu đề Where to find us)" : "True";
    }

    private static string Cau4(Application a, Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
          return "False(không thêm xóa table)";
        if (d.Tables[1].Rows.HeadingFormat != 9999999)
          return "False(cho tiêu đề lặp lại)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau5(Application a, Document d)
    {
      try
      {
        if (d.Sections.Count < 3)
          return "False";
        if (d.Sections[2].PageSetup.TextColumns.Count != 2)
          return "False(chia 2 cột)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau6(Application a, Document d)
    {
      try
      {
        object Index1 = (object) "Picture 1";
        if (d.Shapes[ref Index1].WrapFormat.Type != WdWrapType.wdWrapThrough)
          return "False (Worng picture  layout)";
        object Index2 = (object) "Picture 1";
        if ((double) d.Shapes[ref Index2].Left != -999996.0)
          return "False (Alignment Right)";
        object Index3 = (object) "Picture 1";
        if ((double) d.Shapes[ref Index3].Top != -999999.0)
          return "False (Alignment top)";
        object Index4 = (object) "Picture 1";
        if (d.Shapes[ref Index4].RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin)
          return "False(HorizontalPositionMargin)";
        object Index5 = (object) "Picture 1";
        return d.Shapes[ref Index5].RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin ? "False(VerticalPositionMargin)" : "True";
      }
      catch (Exception ex)
      {
        return "False(change wrap text to picture)";
      }
    }

    private static string Cau7(Application a, Document d)
    {
      try
      {
        int num = 0;
        foreach (Shape shape in d.Shapes)
        {
          if (shape.Name.Contains("Text Box"))
          {
            if (shape.TextFrame.TextRange.Text.Contains("Figure 1-Man playing golf"))
              return "True";
            num = -1;
          }
        }
        if (num == -1)
          return "False (Figure 1-Man playing golf)";
        return num == 0 ? "Falss(inser caption to picture)" : "True";
      }
      catch (Exception ex)
      {
        return "False(wrong from quesion 7)";
      }
    }

    private static string Cau8(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (d.Paragraphs[Index].Range.Text != "Eagle – Two strokes under par\r\a" && Index < d.Paragraphs.Count)
          ++Index;
        return Index < d.Paragraphs.Count ? "False(Convert to text)" : "True";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
    }

    private static string Cau9(Application a, Document d)
    {
      try
      {
        if (d.Sections[2].PageSetup.TextColumns.Count != 2)
          return "False(Column<>2)";
        string text = d.Sections[2].Range.Text;
        if (!text.Substring(0, 40).Contains("Albatross"))
          return "False(begin Albatross – Three... )";
        return !text.Substring(text.Length - 40).Contains("Triple bogey – three strokes over par") ? "False(end with ...three strokes over par)" : "True";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
    }

    private static string Cau10(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Eagle – Two strokes under par") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        // ISSUE: variable of a compiler-generated type
        ListFormat listFormat1 = d.Paragraphs[Index - 1].Range.ListFormat;
        // ISSUE: variable of a compiler-generated type
        ListFormat listFormat2 = d.Paragraphs[Index + 5].Range.ListFormat;
        if (!listFormat1.ListPictureBullet.IsPictureBullet)
          return "False(Albatross – Three strokes under par)";
        if (!listFormat2.ListPictureBullet.IsPictureBullet)
          return "False(Triple bogey – three strokes over par)";
        if (listFormat1.ListTemplate.ListLevels[1].NumberPosition.ToString() != "28.8")
          return "False(Number Position)";
        if (listFormat1.ListTemplate.ListLevels[1].TextPosition.ToString() != "50.4")
          return "False(Text Position)";
        return (double) d.Paragraphs[Index - 1].LineSpacing != 18.0 || (double) d.Paragraphs[Index + 5].LineSpacing != 18.0 ? "False(1.5 line)" : "True";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
    }

    private static string Cau11(Application a, Document d)
    {
      try
      {
        if (d.Footnotes.Count != 1)
          return "False(number of footnote)";
        if (d.Footnotes[1].Range.Text != "Also known as double eagle")
          return "False(Also known as double eagle)";
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Eagle – Two strokes under par") && Index < d.Paragraphs.Count)
          ++Index;
        return d.Paragraphs[Index - 1].Range.Text != "Albatross\u0002 – Three strokes under par\r" ? "False(Albatross)" : "True";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
    }

    private static string Cau12(Application a, Document d)
    {
      try
      {
        if (d.Hyperlinks.Count != 1)
          return "False(number hyperlink)";
        object Index = (object) 1;
        // ISSUE: variable of a compiler-generated type
        Hyperlink hyperlink = d.Hyperlinks[ref Index];
        if (hyperlink.TextToDisplay != "rules")
          return "False(rules)";
        return hyperlink.Name != "_Rules_and_Regulations" ? "False(_Rules_and_Regulations Heading)" : "True";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
    }

    private static string Cau13(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Match Play") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        return !d.Paragraphs[Index + 1].Range.Text.Contains("Two players (or two teams) play each hole as a separate contest against ") ? "False(insert text after 'Match Play' heading)" : "True";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
    }

    private static string Cau14(Application a, Document d)
    {
      try
      {
        int Index1 = 1;
        while (!d.Paragraphs[Index1].Range.Text.Contains("Play the ball as it lies, play the course as find it, and if you cannot do either, do what is fair.") && Index1 < d.Paragraphs.Count)
          ++Index1;
        if (Index1 < d.Paragraphs.Count)
          return "False(cut not copy)";
        object Index2 = (object) "Text Box 2";
        if (!d.Shapes[ref Index2].TextFrame.TextRange.Text.Contains("Play the ball as it lies, play the course as find it, and if you cannot do either, do what is fair."))
          return "False (Conten of textbox)";
        object Index3 = (object) "Text Box 2";
        if (d.Shapes[ref Index3].WrapFormat.Type != WdWrapType.wdWrapSquare)
          return "False(WrapText<>Square)";
        object Index4 = (object) "Text Box 2";
        if (d.Shapes[ref Index4].RelativeHorizontalSize != WdRelativeHorizontalSize.wdRelativeHorizontalSizeMargin)
          return "False(H_S)";
        object Index5 = (object) "Text Box 2";
        if (d.Shapes[ref Index5].RelativeVerticalSize != WdRelativeVerticalSize.wdRelativeVerticalSizeMargin)
          return "False(V_S)";
        object Index6 = (object) "Text Box 2";
        if (d.Shapes[ref Index6].Top.ToString() != "115.2")
          return "False(Top)";
        object Index7 = (object) "Text Box 2";
        if ((double) d.Shapes[ref Index7].Left != -999998.0)
          return "False(Left)";
        object Index8 = (object) "Text Box 2";
        if (d.Shapes[ref Index8].RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin)
          return "False(H_P)";
        object Index9 = (object) "Text Box 2";
        if (d.Shapes[ref Index9].RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin)
          return "False(V_P)";
        object Index10 = (object) "Text Box 2";
        if (d.Shapes[ref Index10].TextFrame.TextRange.Font.Name != "Calibri")
          return "False(Format font)";
        object Index11 = (object) "Text Box 2";
        return d.Shapes[ref Index11].TextFrame.TextRange.Font.TextColor.RGB != -738131969 ? "False(Format color)" : "True";
      }
      catch (Exception ex)
      {
        return "False(Inset Asting Qu)";
      }
    }

    private static string Cau15(Application a, Document d)
    {
      try
      {
        if (d.Bookmarks.Count != 1)
          return "False (Number of bookmark)";
        object Index1 = (object) 1;
        if (d.Bookmarks[ref Index1].Name != "Rules")
        {
          object Index2 = (object) 1;
          return "False (name=" + d.Bookmarks[ref Index2].Name + ")";
        }
        object Index3 = (object) 1;
        if (!d.Bookmarks[ref Index3].Range.Text.Contains("Rules and Regulations"))
        {
          object Index4 = (object) 1;
          return "False (text=" + d.Bookmarks[ref Index4].Range.Text + ")";
        }
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau16(Application a, Document d)
    {
      try
      {
        int Index1 = 1;
        while (Index1 < d.Paragraphs.Count && !d.Paragraphs[Index1].Range.Text.Contains("World Wide Sport"))
          ++Index1;
        if (Index1 >= d.Paragraphs.Count)
          return "False (heading World Wide Sport was edited)";
        int Index2 = Index1 + 1;
        while (Index2 < Index2 + 8 && !d.Paragraphs[Index2].Range.Text.Contains("Number of Courses\r\a"))
          ++Index2;
        if (Index2 >= Index2 + 8)
          return "False (corvet to table use comma)";
        if (d.Tables[d.Tables.Count].Columns.PreferredWidthType != WdPreferredWidthType.wdPreferredWidthPercent)
          return "False (Not auto fit Windows)";
        if (d.Tables[d.Tables.Count].Columns.Count != 3)
          return "False (number of columns<>3)";
        if (d.Tables[d.Tables.Count].Rows.Count != 13)
          return "False (number of raws<>13)";
        if (d.Tables[d.Tables.Count].Rows[1].Cells[1].Shading.BackgroundPatternColorIndex != WdColorIndex.wdGray50 || d.Tables[d.Tables.Count].Rows[2].Cells[1].Shading.BackgroundPatternColorIndex != WdColorIndex.wdWhite || d.Tables[d.Tables.Count].Rows[2].Cells[1].Borders[WdBorderType.wdBorderRight].ColorIndex != WdColorIndex.wdAuto)
          return "False (Not correct table style)";
        if (d.Tables[d.Tables.Count].Rows[1].Cells[2].Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
          return "False(Cente colum 2)";
        if (d.Tables[d.Tables.Count].Rows[1].Cells[3].Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
          return "False(Cente colum 3)";
        if (d.Tables[d.Tables.Count].Rows[10].Cells[2].Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
          return "False(Cente colum 2)";
        if (d.Tables[d.Tables.Count].Rows[10].Cells[3].Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
          return "False(Cente colum 3)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau17(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Major Championships"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False (heading Major Championships was edited)";
        if (!d.Paragraphs[Index - 1].Range.Text.Contains("Table 1-Golf courses by country."))
          return "False(Table 1-Golf courses by country.)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau18(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count - 1 && !d.Paragraphs[Index].Range.Text.Contains("Major Championships"))
          ++Index;
        if (Index >= d.Paragraphs.Count - 1)
          return "False (Major Championships heading not found)";
        if (d.Paragraphs[Index + 2].Range.InlineShapes.Count != 1)
          return "False (number of SmartArt)";
        if (d.Paragraphs[Index + 2].Range.InlineShapes[1].Type != WdInlineShapeType.wdInlineShapeSmartArt)
          return "False (Not SmartArt)";
        if (d.Paragraphs[Index + 2].Range.InlineShapes[1].SmartArt.Layout.Name != "Lined List")
          return "False (" + d.Paragraphs[Index + 2].Range.InlineShapes[1].SmartArt.Layout.Name + ")";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau19(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count - 1 && !d.Paragraphs[Index].Range.Text.Contains("Major Championships"))
          ++Index;
        if (Index >= d.Paragraphs.Count - 1)
          return "False (Major Championships heading not found)";
        if (d.Paragraphs[Index + 2].Range.InlineShapes.Count != 1)
          return "False (number of SmartArt)";
        if (d.Paragraphs[Index + 2].Range.InlineShapes[1].Height.ToString() != "180")
          return "False (Height)";
        if (d.Paragraphs[Index + 2].Range.InlineShapes[1].Width.ToString() != "360")
          return "False (Width)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau20(Application a, Document d)
    {
      string str = "True";
      object documentProperties = d.BuiltInDocumentProperties;
      object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder) null, documentProperties, new object[1]
      {
        (object) "Title"
      });
      return target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder) null, target, new object[0]).ToString() != "Golf" ? "False" : str;
    }

    private static string Cau21(Application a, Document d)
    {
      try
      {
        if (a.Options.SaveInterval != 7)
          return "False (Save:" + (object) a.Options.SaveInterval + ")";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau22(Application a, Document d)
    {
      string str = "True";
      return d.ShowGrammaticalErrors ? "False(Hide grammar error)" : str;
    }

    private static string Cau23(Application a, Document d)
    {
      try
      {
        if (!d.ActiveWindow.DocumentMap)
          return "False";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau24(Application a, Document d)
    {
      string str = "True";
      return a.Options.SaveInterval != 13 || !d.EmbedTrueTypeFonts ? "False" : str;
    }

    private static string Cau25(Application a, Document d)
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
          // ISSUE: reference to a compiler-generated field
          if (CheckWork3.\u003CCau25\u003Eo__SiteContainer18.\u003C\u003Ep__Site19 == null)
          {
            // ISSUE: reference to a compiler-generated field
            CheckWork3.\u003CCau25\u003Eo__SiteContainer18.\u003C\u003Ep__Site19 = CallSite<Func<CallSite, object, TableStyle>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (TableStyle), typeof (CheckWork3)));
          }
          // ISSUE: reference to a compiler-generated field
          // ISSUE: reference to a compiler-generated field
          // ISSUE: variable of a compiler-generated type
          TableStyle tableStyle = CheckWork3.\u003CCau25\u003Eo__SiteContainer18.\u003C\u003Ep__Site19.Target((CallSite) CheckWork3.\u003CCau25\u003Eo__SiteContainer18.\u003C\u003Ep__Site19, table.Style);
          str = "True";
        }
      }
      return str;
    }

    private static string Cau26(Application a, Document d) => "";

    private static string Cau27(Application a, Document d) => "";

    private static string Cau28(Application a, Document d) => "";

    private static string Cau29(Application a, Document d) => "";
  }
}
