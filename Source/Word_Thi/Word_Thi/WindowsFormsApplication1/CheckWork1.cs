// Decompiled with JetBrains decompiler
// Type: MOS_WORD_TEST.CheckWork1
// Assembly: MOS_WORD_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using Microsoft.Office.Interop.Word;
using System;
using System.Reflection;

namespace MOS_WORD_TEST
{
  public static class CheckWork1
  {
    public static string Check(int cauLon, Application a, Document d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork1.Cau0(a, d);
        case 1:
          return CheckWork1.Cau1(a, d);
        case 2:
          return CheckWork1.Cau2(a, d);
        case 3:
          return CheckWork1.Cau3(a, d);
        case 4:
          return CheckWork1.Cau4(a, d);
        case 5:
          return CheckWork1.Cau5(a, d);
        case 6:
          return CheckWork1.Cau6(a, d);
        case 7:
          return CheckWork1.Cau7(a, d);
        case 8:
          return CheckWork1.Cau8(a, d);
        case 9:
          return CheckWork1.Cau9(a, d);
        case 10:
          return CheckWork1.Cau10(a, d);
        case 11:
          return CheckWork1.Cau11(a, d);
        case 12:
          return CheckWork1.Cau12(a, d);
        case 13:
          return CheckWork1.Cau13(a, d);
        case 14:
          return CheckWork1.Cau14(a, d);
        case 15:
          return CheckWork1.Cau15(a, d);
        case 16:
          return CheckWork1.Cau16(a, d);
        case 17:
          return CheckWork1.Cau17(a, d);
        case 18:
          return CheckWork1.Cau18(a, d);
        case 19:
          return CheckWork1.Cau19(a, d);
        case 20:
          return CheckWork1.Cau20(a, d);
        case 21:
          return CheckWork1.Cau21(a, d);
        case 22:
          return CheckWork1.Cau22(a, d);
        case 23:
          return CheckWork1.Cau23(a, d);
        case 24:
          return CheckWork1.Cau24(a, d);
        case 25:
          return CheckWork1.Cau25(a, d);
        case 26:
          return CheckWork1.Cau26(a, d);
        case 27:
          return CheckWork1.Cau27(a, d);
        case 28:
          return CheckWork1.Cau28(a, d);
        case 29:
          return CheckWork1.Cau29(a, d);
        default:
          return "";
      }
    }

    private static string Cau0(Application a, Document d)
    {
      try
      {
        if (d.Name != "Memo.txt")
          return "False(luu dang text)";
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
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Cultivated Land Summary"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False (heading was edited)";
        if (d.Paragraphs[Index + 4].Range.Tables.Count != 1)
          return "False (place table wrong place)";
        // ISSUE: variable of a compiler-generated type
        Microsoft.Office.Interop.Word.Table table = d.Paragraphs[Index + 4].Range.Tables[1];
        if (table.Rows.Count != 8)
          return "False(rows)";
        if (table.Columns.Count != 4)
          return "False(cols)";
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
        if (d.Tables[1].Rows[d.Tables[1].Rows.Count].Cells.Count != 1)
          return "False(Merge last row)";
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
        if (d.Tables[1].Rows[1].Cells[1].Width.ToString() != "116.85")
          return "False(Distribute column width)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau4(Application a, Document d)
    {
      try
      {
        object Index = (object) "Text Box 2";
        return d.Shapes[ref Index].TextFrame.TextRange.Text != "Requires committee review\r" ? "False(Requires committee review)" : "True";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
    }

    private static string Cau5(Application a, Document d)
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

    private static string Cau6(Application a, Document d)
    {
      try
      {
        if (!a.ActiveWindow.View.ShowTabs)
          return "False(Tab)";
        if (a.ActiveWindow.View.ShowSpaces)
          return "False(Spaces)";
        if (a.ActiveWindow.View.ShowAll)
          return "False(turn off show All)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau7(Application a, Document d)
    {
      string str = "True";
      return d.Sections[2].PageSetup.TextColumns.Count != 2 ? "False (must section 2 have two colums)" : str;
    }

    private static string Cau8(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "An eruption\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count)
        return "False (text was modified)";
      // ISSUE: variable of a compiler-generated type
      ListFormat listFormat1 = d.Paragraphs[Index].Range.ListFormat;
      // ISSUE: variable of a compiler-generated type
      ListFormat listFormat2 = d.Paragraphs[Index + 2].Range.ListFormat;
      // ISSUE: variable of a compiler-generated type
      ListFormat listFormat3 = d.Paragraphs[Index + 6].Range.ListFormat;
      // ISSUE: variable of a compiler-generated type
      ListFormat listFormat4 = d.Paragraphs[Index + 8].Range.ListFormat;
      if (listFormat1.ListPictureBullet.AlternativeText != "*")
        return "False (" + d.Paragraphs[Index].Range.Text + ")";
      if (listFormat2.ListPictureBullet.AlternativeText != "*")
        return "False (" + d.Paragraphs[Index + 2].Range.Text + ")";
      if (listFormat3.ListPictureBullet.AlternativeText != "*")
        return "False (" + d.Paragraphs[Index + 6].Range.Text + ")";
      if (listFormat4.ListPictureBullet.AlternativeText != "*")
        return "False (" + d.Paragraphs[Index + 8].Range.Text + ")";
      if (!listFormat1.ListPictureBullet.IsPictureBullet || !listFormat2.ListPictureBullet.IsPictureBullet || !listFormat3.ListPictureBullet.IsPictureBullet || !listFormat4.ListPictureBullet.IsPictureBullet)
        return "False (PictureBullet)";
      return (double) listFormat1.ListTemplate.ListLevels[1].NumberPosition != 36.0 ? "False (Position of bullet)" : str;
    }

    private static string Cau9(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Stages of Volcanic Activity\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count)
        return "False (text was modified)";
      if (d.Paragraphs[Index + 1].Range.Text == "In Suspension\r\a")
        return "False(In Suspension)";
      if (d.Paragraphs[Index + 2].Range.Text == "Son\r\a")
        return "False(Son)";
      if (d.Paragraphs[Index + 3].Range.Text == "Gas and steam\r\a")
        return "False(Gas and steam)";
      if (d.Paragraphs[Index + 4].Range.Text == "Pyroclastic\r\a")
        return "False(Pyroclastic)";
      if (d.Paragraphs[Index + 5].Range.Text == "Slag\r\a")
        return "False(Slag)";
      if (d.Paragraphs[Index + 6].Range.Text == "White Island\r\a")
        return "False(White Island)";
      if (d.Paragraphs[Index + 7].Range.Text == "Mt. Hekla\r\a")
        return "False(Mt. Hekla)";
      if (d.Paragraphs[Index + 8].Range.Text == "Italy\r\a")
        return "False(Italy)";
      if (d.Paragraphs[Index + 9].Range.Text == "New Zealand\r\a")
        return "False(New Zealand)";
      foreach (Microsoft.Office.Interop.Word.Table table in d.Tables)
      {
        if (table.Columns.Count == 1)
          return "False (col=" + (object) table.Columns.Count + ")";
        if (table.Rows.Count == 9)
          return "False (raw=" + (object) table.Rows.Count + ")";
      }
      return str;
    }

    private static string Cau10(Application a, Document d)
    {
      string str = "True";
      if (d.Paragraphs[1].DropCap.Position != WdDropPosition.wdDropNormal)
        return "False (Position)";
      return d.Paragraphs[1].DropCap.LinesToDrop != 3 ? "False (LinesToDrop)" : str;
    }

    private static string Cau11(Application a, Document d)
    {
      string str = "True";
      if (d.InlineShapes.Count != 1)
        return "False (number of InlineShapes)";
      if (d.InlineShapes[1].SmartArt.Nodes.Count != 2)
        return "False (Number of Node)";
      if (d.InlineShapes[1].SmartArt.Nodes[(object) 1].TextFrame2.TextRange.Text != "In Suspension")
        return "False(Text 1)";
      if ((int) d.InlineShapes[1].Width != 283)
        return "False(Width)";
      if ((int) d.InlineShapes[1].Height != 113)
        return "False(Height)";
      if (d.InlineShapes[1].SmartArt.Nodes[(object) 2].TextFrame2.TextRange.Text != "White Island")
        return "False (Text 2)";
      if (d.InlineShapes[1].SmartArt.Nodes[(object) 2].Nodes.Count != 3 || d.InlineShapes[1].SmartArt.Nodes[(object) 1].Nodes.Count != 4)
        return "False (number of subtext)";
      if (d.InlineShapes[1].SmartArt.Nodes[(object) 1].Nodes[(object) 1].TextFrame2.TextRange.Text != "Son" || d.InlineShapes[1].SmartArt.Nodes[(object) 1].Nodes[(object) 2].TextFrame2.TextRange.Text != "Gas and steam" || d.InlineShapes[1].SmartArt.Nodes[(object) 1].Nodes[(object) 3].TextFrame2.TextRange.Text != "Pyroclastic" || d.InlineShapes[1].SmartArt.Nodes[(object) 1].Nodes[(object) 4].TextFrame2.TextRange.Text != "Slag")
        return "False (subtext 1)";
      if (d.InlineShapes[1].SmartArt.Nodes[(object) 2].Nodes[(object) 1].TextFrame2.TextRange.Text != "Mt. Hekla" || d.InlineShapes[1].SmartArt.Nodes[(object) 2].Nodes[(object) 2].TextFrame2.TextRange.Text != "Italy")
        return "False (subtext 2)";
      return d.InlineShapes[1].SmartArt.Nodes[(object) 2].Nodes[(object) 3].TextFrame2.TextRange.Text != "New Zealand" ? "False  (subtext 2)" : str;
    }

    private static string Cau12(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (!d.Paragraphs[Index].Range.Text.Contains("Figure 1-Some contain iron, magnesium, silica, or") && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count)
        return "False (Figure 1-Some contain iron, magnesium, silica, or aluminum)";
      if (d.Paragraphs[Index - 2].Range.Text != "Stages of Volcanic Activity\r")
        return "False(Stages of Volcanic Activity)";
      if (d.Paragraphs[Index - 1].Range.Text != "/\r")
        return "False (below smartArt)";
      return d.Paragraphs[Index + 1].Range.Text != "\f" ? "False (above break section)" : str;
    }

    private static string Cau13(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Cinder Cones\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count)
        return "False (Text was modified)";
      return !d.Paragraphs[Index + 2].Range.Text.Contains("Composite volcanoes are very large and are formed from alternating explosive and quiet eruptions") ? "False(paragaph " + (object) (Index + 2) + ")" : str;
    }

    private static string Cau14(Application a, Document d)
    {
      foreach (Microsoft.Office.Interop.Word.Table table in d.Tables)
      {
        if (table.Columns.Count == 8)
        {
          if (table.PreferredWidthType != WdPreferredWidthType.wdPreferredWidthAuto)
            return "False(WidthAuto)";
          // ISSUE: reference to a compiler-generated method
          if (table.Cell(1, 1).Range.Text != "Pops\r\a")
            return "False(wrong content)";
          if (table.AutoFormatType != 1)
            return "False(AutoFormatType)";
          if (table.Rows.Count != 11)
            return "False(Rows)";
          // ISSUE: reference to a compiler-generated method
          float width = table.Cell(1, 1).Width;
          // ISSUE: reference to a compiler-generated method
          if ((int) table.Cell(1, 1).Width != 32)
            return "False(Width)";
          if (table.Borders.Count != 8)
            return "False (Borders)";
          if (table.Borders[WdBorderType.wdBorderBottom].ColorIndex != WdColorIndex.wdGray25)
            return "False (border color)";
          if (table.AutoFormatType != 1)
            return "False(AutoFormatType)";
          // ISSUE: reference to a compiler-generated method
          // ISSUE: reference to a compiler-generated method
          if (table.Cell(2, 2).Range.Text != "GRANDVILLE CHRISTIAN SCH\r\a" || table.Cell(11, 2).Range.Text != "Martin, MI\r\a")
            return "False (content)";
        }
      }
      return "True";
    }

    private static string Cau15(Application a, Document d)
    {
      string str = "True";
      int Index = 1;
      while (d.Paragraphs[Index].Range.Text != "Table 1-Types of volcanoes 2008\r" && Index < d.Paragraphs.Count)
        ++Index;
      if (Index >= d.Paragraphs.Count)
        return "False (Table 1-Types of volcanoes 2008)";
      if (d.Paragraphs[Index - 2].Range.Text != "8\r\a" || d.Paragraphs[Index - 1].Range.Text != "\r\a")
        return "False(below table)";
      return !d.Paragraphs[Index + 1].Range.Text.Contains("Types of Volcanoes") ? "False(above heading Types of Volcanoes)" : str;
    }

    private static string Cau16(Application a, Document d)
    {
      string str = "True";
      object Index = (object) "Picture 1";
      // ISSUE: variable of a compiler-generated type
      Shape shape = d.Shapes[ref Index];
      if (shape.WrapFormat.Type != WdWrapType.wdWrapThrough)
        return "False(Through)";
      if ((double) shape.Left != -999996.0)
        return "False(HorizontalRelativeWith)";
      if ((double) shape.Top != -999997.0)
        return "False(VerticalRelativeWith)";
      return shape.RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin || shape.RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin ? "False(Margin)" : str;
    }

    private static string Cau17(Application a, Document d)
    {
      string str = "True";
      for (int Index = 1; Index < d.Sections.Count; ++Index)
      {
        if (d.Sections[Index].PageSetup.Orientation == WdOrientation.wdOrientLandscape)
          return "False(not Landscape section 1)";
      }
      return d.Sections[d.Sections.Count].PageSetup.Orientation != WdOrientation.wdOrientLandscape ? "False(OrientLandscape)" : str;
    }

    private static string Cau18(Application a, Document d)
    {
      string str = "True";
      try
      {
        object Index1 = (object) "Group 198";
        // ISSUE: variable of a compiler-generated type
        Shape shape = d.Shapes[ref Index1];
        if (shape.RelativeHorizontalSize != WdRelativeHorizontalSize.wdRelativeHorizontalSizeMargin)
          return "False(Horizontal)";
        if (shape.RelativeVerticalSize != WdRelativeVerticalSize.wdRelativeVerticalSizeMargin)
          return "False(Vertical)";
        if ((double) shape.Top != -999997.0)
          return "False(Top)";
        if ((double) shape.Left != -999998.0)
          return "False(Left)";
        if (shape.RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin || shape.RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin)
          return "False(Margin)";
        object Index2 = (object) "Text Box 200";
        // ISSUE: variable of a compiler-generated type
        Shape groupItem = shape.GroupItems[ref Index2];
        return groupItem.TextFrame.TextRange.Text != "ANDESITE IS ONE OF THE MOST COMMON VOLCANIC ROCKS AND CAN CONTAIN OLIVINE\r" ? "False(ANDESITE IS ONE OF THE MOST COMMON VOLCANIC ROCKS AND CAN CONTAIN OLIVINE)" : str;
      }
      catch (Exception ex)
      {
        return "False (Wrong text box)";
      }
    }

    private static string Cau19(Application a, Document d)
    {
      string str = "True";
      if (d.Footnotes.Count != 1)
        return "False(number of Footnote)";
      return d.Footnotes[1].Range.Text != "We can search it online from 2008." ? "False(We can search it online from 2008.)" : str;
    }

    private static string Cau20(Application a, Document d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Hyperlink hyperlink1 = (Hyperlink) null;
      foreach (Hyperlink hyperlink2 in d.Hyperlinks)
      {
        if (hyperlink2.Name == "http://www.mos1000.wordpress.com/")
          hyperlink1 = hyperlink2;
      }
      if (hyperlink1 == null)
        return "False(http://www.mos1000.wordpress.com/)";
      return !hyperlink1.TextToDisplay.Contains("fragments") ? "False(fragments)" : str;
    }

    private static string Cau21(Application a, Document d)
    {
      string str = "True";
      if (d.Bookmarks.Count != 1)
        return "False(number of bookmark)";
      object Index = (object) "Scoring";
      return !d.Bookmarks[ref Index].Range.Text.Contains("Cinder Cones") ? "False(Cinder Cones)" : str;
    }

    private static string Cau22(Application a, Document d)
    {
      string str = "True";
      return !a.ActiveWindow.ActivePane.DisplayRulers ? "False(DisplayRulers)" : str;
    }

    private static string Cau23(Application a, Document d)
    {
      string str = "True";
      object documentProperties = d.BuiltInDocumentProperties;
      object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (Binder) null, documentProperties, new object[1]
      {
        (object) "Subject"
      });
      return target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (Binder) null, target, new object[0]).ToString() != "Volcanoes" ? "False(Volcanoes)" : str;
    }

    private static string Cau24(Application a, Document d)
    {
      string str = "True";
      if (a.Options.SaveInterval != 13)
        return "False(" + (object) a.Options.SaveInterval + ")";
      return !d.EmbedTrueTypeFonts ? "False(Embed Font)" : str;
    }

    private static string Cau25(Application a, Document d)
    {
      foreach (Microsoft.Office.Interop.Word.Table table in d.Tables)
      {
        if (table.Columns.Count == 8)
        {
          if (table.PreferredWidthType != WdPreferredWidthType.wdPreferredWidthAuto)
            return "False(fix content)";
          // ISSUE: reference to a compiler-generated method
          if (table.Cell(1, 1).Range.Text != "Michigan 2016 Region 4 Aesop Gone Viral Div 1\r\a")
            return "False(cell[1,1])";
          if (table.AutoFormatType != 1)
            return "False(AutoFormat)";
          if (table.Rows.Count != 10)
            return "False(row)";
          // ISSUE: reference to a compiler-generated method
          float width = table.Cell(1, 1).Width;
          // ISSUE: reference to a compiler-generated method
          if ((int) table.Cell(1, 1).Width != 506)
            return "False(auto fix content)";
          if (table.Borders.Count != 8 || table.Borders[WdBorderType.wdBorderBottom].ColorIndex != WdColorIndex.wdGray25)
            return "False(Table style)";
        }
      }
      return "True";
    }

    private static string Cau26(Application a, Document d) => "";

    private static string Cau27(Application a, Document d) => "";

    private static string Cau28(Application a, Document d) => "";

    private static string Cau29(Application a, Document d) => "";
  }
}
