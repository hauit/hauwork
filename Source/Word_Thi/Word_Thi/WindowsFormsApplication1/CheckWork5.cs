// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.CheckWork5
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
  public static class CheckWork5
  {
    public static string Check(int cauLon, Application a, Document d)
    {
      switch (cauLon)
      {
        case 0:
          return CheckWork5.Cau1(a, d);
        case 1:
          return CheckWork5.Cau0(a, d);
        case 2:
          return CheckWork5.Cau2(a, d);
        case 3:
          return CheckWork5.Cau3(a, d);
        case 4:
          return CheckWork5.Cau4(a, d);
        case 5:
          return CheckWork5.Cau5(a, d);
        case 6:
          return CheckWork5.Cau6(a, d);
        case 7:
          return CheckWork5.Cau7(a, d);
        case 8:
          return CheckWork5.Cau8(a, d);
        case 9:
          return CheckWork5.Cau9(a, d);
        case 10:
          return CheckWork5.Cau10(a, d);
        case 11:
          return CheckWork5.Cau11(a, d);
        case 12:
          return CheckWork5.Cau12(a, d);
        case 13:
          return CheckWork5.Cau13(a, d);
        case 14:
          return CheckWork5.Cau14(a, d);
        case 15:
          return CheckWork5.Cau15(a, d);
        case 16:
          return CheckWork5.Cau16(a, d);
        case 17:
          return CheckWork5.Cau17(a, d);
        case 18:
          return CheckWork5.Cau18(a, d);
        case 19:
          return CheckWork5.Cau19(a, d);
        case 20:
          return CheckWork5.Cau20(a, d);
        case 21:
          return CheckWork5.Cau21(a, d);
        case 22:
          return CheckWork5.Cau22(a, d);
        case 23:
          return CheckWork5.Cau23(a, d);
        case 24:
          return CheckWork5.Cau24(a, d);
        case 25:
          return CheckWork5.Cau25(a, d);
        case 26:
          return CheckWork5.Cau26(a, d);
        case 27:
          return CheckWork5.Cau27(a, d);
        case 28:
          return CheckWork5.Cau28(a, d);
        case 29:
          return CheckWork5.Cau29(a, d);
        default:
          return "";
      }
    }

    private static string Cau0(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Focus on the text you want") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(không chỉnh sửa văn bản)";
        if (d.Paragraphs[Index].Range.ListFormat.ListString != "▸")
          return "Falas(25B8)";
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
        if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Characters[3].Font.TextColor.RGB != -738131969)
          return "False";
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
        if (d.Paragraphs[4].Range.Text.Contains("in the embed code"))
          return "False(in the embed code)";
        if (!d.Paragraphs[6].Range.Text.Contains("new look"))
          return "False(new look)";
        if (d.Paragraphs[8].Range.Text.Contains("and SmartArt"))
          return "False(and SmartArt)";
        if (d.Paragraphs[10].Range.Characters[5].Font.Bold == -1)
          return "False(không chấp nhận định dạng)";
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

    private static string Cau4(Application a, Document d)
    {
      try
      {
        if (d.Bookmarks.Count != 1)
          return "False (Number of bookmark)";
        object Index1 = (object) 1;
        if (d.Bookmarks[ref Index1].Name != "BeEntertaining")
        {
          object Index2 = (object) 1;
          return "False (name=" + d.Bookmarks[ref Index2].Name + ")";
        }
        object Index3 = (object) 1;
        if (!d.Bookmarks[ref Index3].Range.Text.Contains("Be Entertaining"))
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

    private static string Cau5(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Training software") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        return (double) d.Paragraphs[Index].Format.SpaceAfter != 132.0 ? "False(" + (object) d.Paragraphs[Index].Format.SpaceAfter + ")" : "True";
      }
      catch (Exception ex)
      {
        return "False(somthing wrong)";
      }
    }

    private static string Cau6(Application a, Document d)
    {
      try
      {
        object Index1 = (object) 1;
        if (!(d.Hyperlinks[ref Index1].Name != "http://www.mos1000.wordpress.com/"))
          return "True";
        object Index2 = (object) 1;
        return "False(" + d.Hyperlinks[ref Index2].Name + ")";
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
        return d.Sections[3].PageSetup.TextColumns.Count != 2 ? "False(Column<>2)" : "True";
      }
      catch (Exception ex)
      {
        return "False(section 3)";
      }
    }

    private static string Cau8(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Microsoft Word") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        // ISSUE: variable of a compiler-generated type
        ListFormat listFormat = d.Paragraphs[Index + 1].Range.ListFormat;
        if (listFormat.ListTemplate.ListLevels[1].NumberPosition.ToString() != "25.2")
          return "False(" + listFormat.ListTemplate.ListLevels[1].NumberPosition.ToString() + ")";
        if (listFormat.ListString != "\uF0B7")
          return "False(Filled circle)";
        return (double) d.Paragraphs[Index + 1].LineSpacing != 18.0 ? "False(" + (object) d.Paragraphs[Index + 1].LineSpacing + ")" : "True";
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
        int Index = 1;
        while (d.Paragraphs[Index].Range.Text != "Nome\r" && Index < d.Paragraphs.Count)
          ++Index;
        return Index >= d.Paragraphs.Count ? "False(Convert to table with paragraph not tab)" : "True";
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
        if (d.Footnotes.Count != 1)
          return "False(number of footnote)";
        return d.Footnotes[1].Range.Text != "We can search it online from 2008." ? "False(We can search it online from 2008.)" : "True";
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
        object Index1 = (object) "Picture 1";
        float num = d.Shapes[ref Index1].Left;
        if (num.ToString() != "-999996")
          return "False(H)";
        object Index2 = (object) "Picture 1";
        num = d.Shapes[ref Index2].Top;
        if (num.ToString() != "-999997")
          return "False(V)";
        object Index3 = (object) "Picture 1";
        if (d.Shapes[ref Index3].RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin)
          return "Falas(RelativeHorizontalPositionMargin)";
        object Index4 = (object) "Picture 1";
        if (d.Shapes[ref Index4].RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin)
          return "Falas(RelativeVerticalPositionMargin)";
        object Index5 = (object) "Picture 1";
        return d.Shapes[ref Index5].WrapFormat.Type != WdWrapType.wdWrapThrough ? "False(WrapThrough)" : "True";
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
        int Index = 1;
        while (d.Paragraphs[Index].Range.Text != "An eruption\r" && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modyfied)";
        // ISSUE: reference to a compiler-generated field
        if (CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sitec == null)
        {
          // ISSUE: reference to a compiler-generated field
          CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sitec = CallSite<Func<CallSite, object, Style>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Style), typeof (CheckWork5)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Style style1 = CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sitec.Target((CallSite) CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sitec, d.Paragraphs[Index].Style);
        if (style1.NameLocal != "Heading 2")
          return "False(An eruption)";
        while (d.Paragraphs[Index].Range.Text != "Example\r" && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modyfied)";
        // ISSUE: reference to a compiler-generated field
        if (CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sited == null)
        {
          // ISSUE: reference to a compiler-generated field
          CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sited = CallSite<Func<CallSite, object, Style>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Style), typeof (CheckWork5)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Style style2 = CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sited.Target((CallSite) CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sited, d.Paragraphs[Index].Style);
        if (style2.NameLocal != "Heading 2")
          return "False(Example)";
        while (d.Paragraphs[Index].Range.Text != "Location\r" && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modyfied)";
        // ISSUE: reference to a compiler-generated field
        if (CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sitee == null)
        {
          // ISSUE: reference to a compiler-generated field
          CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sitee = CallSite<Func<CallSite, object, Style>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Style), typeof (CheckWork5)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Style style3 = CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sitee.Target((CallSite) CheckWork5.\u003CCau12\u003Eo__SiteContainerb.\u003C\u003Ep__Sitee, d.Paragraphs[Index].Style);
        return style3.NameLocal != "Heading 2" ? "False(Location)" : "True";
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
        while (d.Paragraphs[Index].Range.Text != "An eruption\r" && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        return !d.Paragraphs[Index + 1].Range.Text.Contains("There are many") ? "False(insert text after 'An eruption' heading)" : "True";
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
        int Index = 1;
        while (d.Paragraphs[Index].Range.Text != "Example\r" && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        return !d.Paragraphs[Index + 1].Range.Text.Contains("Composite volcanoes are") ? "False(insert text after 'Example' heading)" : "True";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
    }

    private static string Cau15(Application a, Document d)
    {
      try
      {
        foreach (Microsoft.Office.Interop.Word.Table table in d.Tables)
        {
          // ISSUE: reference to a compiler-generated field
          if (CheckWork5.\u003CCau15\u003Eo__SiteContainerf.\u003C\u003Ep__Site10 == null)
          {
            // ISSUE: reference to a compiler-generated field
            CheckWork5.\u003CCau15\u003Eo__SiteContainerf.\u003C\u003Ep__Site10 = CallSite<Func<CallSite, object, Style>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Style), typeof (CheckWork5)));
          }
          // ISSUE: reference to a compiler-generated field
          // ISSUE: reference to a compiler-generated field
          // ISSUE: variable of a compiler-generated type
          Style style = CheckWork5.\u003CCau15\u003Eo__SiteContainerf.\u003C\u003Ep__Site10.Target((CallSite) CheckWork5.\u003CCau15\u003Eo__SiteContainerf.\u003C\u003Ep__Site10, table.Style);
          if (style.NameLocal == "Grid Table 4")
          {
            if (table.PreferredWidthType != WdPreferredWidthType.wdPreferredWidthAuto)
              return "False (autofix content)";
            // ISSUE: reference to a compiler-generated method
            string text = table.Cell(1, 1).Range.Text;
            if (!text.Contains("Name"))
              return "False(" + text + ")";
            if (table.AutoFormatType != 1)
              return "False(FormatType)";
            if (table.Columns.Count != 2)
              return "False(Columns)";
            // ISSUE: reference to a compiler-generated method
            float width = table.Cell(1, 1).Width;
            return (double) width > 200.0 ? "False(" + (object) width + ")" : "True";
          }
        }
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "False(Grid Table 4)";
    }

    private static string Cau16(Application a, Document d)
    {
      try
      {
        foreach (Microsoft.Office.Interop.Word.Table table in d.Tables)
        {
          // ISSUE: reference to a compiler-generated field
          if (CheckWork5.\u003CCau16\u003Eo__SiteContainer11.\u003C\u003Ep__Site12 == null)
          {
            // ISSUE: reference to a compiler-generated field
            CheckWork5.\u003CCau16\u003Eo__SiteContainer11.\u003C\u003Ep__Site12 = CallSite<Func<CallSite, object, Style>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Style), typeof (CheckWork5)));
          }
          // ISSUE: reference to a compiler-generated field
          // ISSUE: reference to a compiler-generated field
          // ISSUE: variable of a compiler-generated type
          Style style = CheckWork5.\u003CCau16\u003Eo__SiteContainer11.\u003C\u003Ep__Site12.Target((CallSite) CheckWork5.\u003CCau16\u003Eo__SiteContainer11.\u003C\u003Ep__Site12, table.Style);
          if (style.NameLocal == "Grid Table 4")
          {
            if (table.PreferredWidthType != WdPreferredWidthType.wdPreferredWidthAuto)
              return "False (autofix content)";
            // ISSUE: reference to a compiler-generated method
            string text1 = table.Cell(1, 1).Range.Text;
            if (!text1.Contains("Name"))
              return "False(" + text1 + ")";
            if (table.Rows.Count != 9)
              return "False(rows)";
            // ISSUE: reference to a compiler-generated method
            string text2 = table.Cell(4, 2).Range.Text;
            return !text2.Contains("132") ? "False(" + text2 + ")" : "True";
          }
        }
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "False(Grid Table 4)";
    }

    private static string Cau17(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("200\r\a"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False (conten was edited)";
        if (!d.Paragraphs[Index + 2].Range.Text.Contains("Table 1 Time of training software"))
          return "False(Table 1 Time of training software)";
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
        if (d.Bookmarks.Count != 1)
          return "False(number of bookmark)";
        object Index = (object) "Scoring";
        if (!d.Bookmarks[ref Index].Range.Text.Contains("Schedule"))
          return "False(Schedule)";
      }
      catch (Exception ex)
      {
        return "False (add bookmark name Scoring)";
      }
      return "True";
    }

    private static string Cau19(Application a, Document d)
    {
      try
      {
        object Index1 = (object) "Group 198";
        // ISSUE: variable of a compiler-generated type
        Shape shape = d.Shapes[ref Index1];
        if (shape.ID != 198 || shape.GroupItems.Count != 2)
          return "False (Not currect type)";
        object Index2 = (object) 2;
        string text = shape.GroupItems[ref Index2].TextFrame.TextRange.Text;
        if (!text.ToLower().Contains("andesite is one of the most common volcanic rocks and can contain olivine"))
          return "False (" + text + ")";
        object Index3 = (object) 2;
        if (shape.GroupItems[ref Index3].TextFrame.TextRange.Font.TextColor.RGB != -738131969)
          return "False (keep format)";
        for (int Index4 = 1; Index4 < d.Paragraphs.Count; ++Index4)
        {
          if (d.Paragraphs[Index4].Range.Text.Contains("Andesite is one of the most common "))
            return "False (Cut not copy)";
        }
        if ((double) shape.LeftRelative != -999999.0)
          return "False (H right)";
        if ((double) shape.Left != -999998.0)
          return "False(H left)";
        if ((double) shape.Top != -999997.0 || (double) shape.TopRelative != -999999.0)
          return "False V bottom";
        if (shape.RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin)
          return "False (H Margin)";
        if (shape.RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin)
          return "False (V Margin)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau20(Application a, Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count - 1 && !d.Paragraphs[Index].Range.Text.Contains("Cooling"))
          ++Index;
        if (Index >= d.Paragraphs.Count - 1)
          return "False (Cooling heading not found)";
        if (d.Paragraphs[Index + 1].Range.InlineShapes.Count != 1)
          return "False (number of SmartArt)";
        if (d.Paragraphs[Index + 1].Range.InlineShapes[1].Type != WdInlineShapeType.wdInlineShapeSmartArt)
          return "False (Not SmartArt)";
        if (d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Layout.Name != "Segmented Process")
          return "False (" + d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Layout.Name + ")";
        if (d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Nodes.Count != 2)
          return "False (Two level)";
        if ((double) d.Paragraphs[Index + 1].Range.InlineShapes[1].Height != 113.75)
          return "False (Height)";
        if ((double) d.Paragraphs[Index + 1].Range.InlineShapes[1].Width != 283.64999389648438)
          return "False (Width)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string Cau21(Application a, Document d)
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

    private static string Cau22(Application a, Document d)
    {
      string str = "True";
      object documentProperties = d.BuiltInDocumentProperties;
      object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder) null, documentProperties, new object[1]
      {
        (object) "Title"
      });
      return target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder) null, target, new object[0]).ToString() != "software" ? "False(software)" : str;
    }

    private static string Cau23(Application a, Document d)
    {
      try
      {
        if (a.Options.SaveInterval != 12)
          return "False (Save:" + (object) a.Options.SaveInterval + ")";
        if (!d.EmbedTrueTypeFonts)
          return "False(EmbedTrueTypeFonts)";
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
          if (CheckWork5.\u003CCau25\u003Eo__SiteContainer17.\u003C\u003Ep__Site18 == null)
          {
            // ISSUE: reference to a compiler-generated field
            CheckWork5.\u003CCau25\u003Eo__SiteContainer17.\u003C\u003Ep__Site18 = CallSite<Func<CallSite, object, TableStyle>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (TableStyle), typeof (CheckWork5)));
          }
          // ISSUE: reference to a compiler-generated field
          // ISSUE: reference to a compiler-generated field
          // ISSUE: variable of a compiler-generated type
          TableStyle tableStyle = CheckWork5.\u003CCau25\u003Eo__SiteContainer17.\u003C\u003Ep__Site18.Target((CallSite) CheckWork5.\u003CCau25\u003Eo__SiteContainer17.\u003C\u003Ep__Site18, table.Style);
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
