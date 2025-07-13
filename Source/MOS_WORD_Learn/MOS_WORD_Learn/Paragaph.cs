// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Paragaph
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.Office.Interop.Word;
using System;

namespace WindowsFormsApplication1
{
  internal class Paragaph
  {
    public static string CheckCau(int causo, Application a, _Document d)
    {
      switch (causo)
      {
        case 1:
          return Paragaph.cau26(a, d);
        case 2:
          return Paragaph.cau22(a, d);
        case 3:
          return Paragaph.cau23(a, d);
        case 4:
          return Paragaph.cau7(a, d);
        case 5:
          return Paragaph.cau27(a, d);
        case 6:
          return Paragaph.cau28(a, d);
        case 7:
          return Paragaph.cau24(a, d);
        case 8:
          return Paragaph.cau21(a, d);
        case 9:
          return Paragaph.cau29(a, d);
        case 10:
          return Paragaph.cau25(a, d);
        case 11:
          return Paragaph.cau17(a, d);
        case 12:
          return Paragaph.cau14(a, d);
        case 13:
          return Paragaph.cau10(a, d);
        case 14:
          return Paragaph.cau14(a, d);
        case 15:
          return Paragaph.cau15(a, d);
        case 16:
          return Paragaph.cau16(a, d);
        case 17:
          return Paragaph.cau17(a, d);
        case 18:
          return Paragaph.cau18(a, d);
        case 19:
          return Paragaph.cau19(a, d);
        case 20:
          return Paragaph.cau20(a, d);
        case 21:
          return Paragaph.cau21(a, d);
        case 22:
          return Paragaph.cau22(a, d);
        case 23:
          return Paragaph.cau23(a, d);
        case 24:
          return Paragaph.cau24(a, d);
        case 25:
          return Paragaph.cau25(a, d);
        case 26:
          return Paragaph.cau26(a, d);
        case 27:
          return Paragaph.cau27(a, d);
        case 28:
          return Paragaph.cau28(a, d);
        case 29:
          return Paragaph.cau29(a, d);
        default:
          return "False(default)";
      }
    }

    private static string cau1(Application a, _Document d)
    {
      try
      {
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph1 = (Paragraph) null;
        int Index;
        for (Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Now you can"))
          {
            paragraph1 = d.Paragraphs[Index];
            break;
          }
        }
        if (paragraph1 == null)
          return "False (List edited)";
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph2 = d.Paragraphs[Index + 1];
        if (paragraph2.Range.ListFormat.ListLevelNumber != 2)
          return "False(Level 2)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau2(Application a, _Document d)
    {
      try
      {
        for (int index = 1; index < d.Paragraphs.Count; ++index)
        {
          if ((double) d.Paragraphs[1].LineSpacing != 13.800000190734863)
            return "False(1.15 line)";
        }
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
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph1 = (Paragraph) null;
        int Index;
        for (Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Frequency of guard patrols"))
          {
            paragraph1 = d.Paragraphs[Index];
            break;
          }
        }
        if (paragraph1 == null)
          return "False (Contents edited)";
        if (paragraph1.Range.ListFormat.ListString == "")
          return "False (not bullet)";
        if (paragraph1.Range.ListFormat.ListLevelNumber != 1)
          return "False(Level 1)";
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph2 = d.Paragraphs[Index + 3];
        if (paragraph2.Range.ListFormat.ListString == "")
          return "False (not bullet)";
        if (paragraph2.Range.ListFormat.ListLevelNumber != 1)
          return "False(Level 1)";
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
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph = (Paragraph) null;
        for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("DO NOT COPY OR REDISTRIBUTE"))
          {
            paragraph = d.Paragraphs[Index];
            break;
          }
        }
        if (paragraph == null)
          return "False (text edited)";
        if (paragraph.Range.HighlightColorIndex != WdColorIndex.wdPink)
          return "False (Highligh Pink)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau5(Application a, _Document d)
    {
      try
      {
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph = (Paragraph) null;
        for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Handicap"))
          {
            paragraph = d.Paragraphs[Index];
            break;
          }
        }
        if (paragraph == null)
          return "False (List edited)";
      }
      catch (Exception ex)
      {
        return "False (not picture)";
      }
      return "True";
    }

    private static string cau6(Application a, _Document d)
    {
      try
      {
        for (int index = 1; index < d.Paragraphs.Count; ++index)
        {
          if ((double) d.Paragraphs[1].LineSpacing != 24.0)
            return "False(double line)";
        }
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau7(Application a, _Document d)
    {
      try
      {
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph = (Paragraph) null;
        for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("100% organic"))
          {
            paragraph = d.Paragraphs[Index];
            break;
          }
        }
        if (paragraph == null)
          return "False (List edited)";
        if (!paragraph.Range.ListFormat.ListString.Contains("A"))
          return "False(" + paragraph.Range.ListFormat.ListString + ")";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau8(Application a, _Document d)
    {
      try
      {
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph = (Paragraph) null;
        for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Vegetables"))
          {
            paragraph = d.Paragraphs[Index + 7];
            break;
          }
        }
        if (paragraph == null)
          return "False (List edited)";
        if (paragraph.Range.ListFormat.ListString != "1071.")
          return "False(" + paragraph.Range.ListFormat.ListString + ")";
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
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph1 = (Paragraph) null;
        int Index;
        for (Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Prepare sinful cakes, pies, and pastries"))
          {
            paragraph1 = d.Paragraphs[Index];
            break;
          }
        }
        if (paragraph1 == null)
          return "False (List edited)";
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph2 = d.Paragraphs[Index + 1];
        if (paragraph2.Range.ListFormat.ListLevelNumber != 2)
          return "False(Level 2)";
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
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph = (Paragraph) null;
        int Index;
        for (Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Thorp"))
          {
            paragraph = d.Paragraphs[Index];
            break;
          }
        }
        if (paragraph == null)
          return "False (text was edited)";
        if (paragraph.Range.HighlightColorIndex != WdColorIndex.wdTurquoise || d.Paragraphs[Index + 1].Range.HighlightColorIndex != WdColorIndex.wdTurquoise || d.Paragraphs[Index + 2].Range.HighlightColorIndex != WdColorIndex.wdTurquoise)
          return "False (Highligh Turquoise)";
        if (d.Paragraphs[Index + 3].Range.HighlightColorIndex != WdColorIndex.wdTurquoise)
          return "False (Highligh Turquoise)";
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
        // ISSUE: variable of a compiler-generated type
        Paragraph paragraph = (Paragraph) null;
        for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Handicap"))
          {
            paragraph = d.Paragraphs[Index];
            break;
          }
        }
        if (paragraph == null)
          return "False (List edited)";
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
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Know Your Topic") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Know Your Topic text)";
        if (d.Paragraphs[Index].Range.ListFormat.ListString != "\uF0B7")
          return "False(Bullet to Know Your Topic)";
        while (!d.Paragraphs[Index].Range.Text.Contains("Be Excited About the Topic") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Be Excited About the Topic text)";
        if (d.Paragraphs[Index].Range.ListFormat.ListString != "\uF0B7")
          return "False(Bullet to Be Excited About the Topic)";
        while (!d.Paragraphs[Index].Range.Text.Contains("Be Entertaining") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Be Entertaining text)";
        if (d.Paragraphs[Index].Range.ListFormat.ListString != "\uF0B7")
          return "False(Bullet to Be Entertaining)";
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
        for (int index = 1; index < d.Paragraphs.Count; ++index)
        {
          if ((double) d.Paragraphs[1].LineSpacing != 13.800000190734863)
            return "False(1.15 line)";
        }
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
        while (!d.Paragraphs[Index].Range.Text.Contains("On the Insert tab, the galleries include items ") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        if (d.Paragraphs[Index].Range.Font.Bold != 0)
          return "Falas(paste only value)";
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
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Three-String") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        if (d.Paragraphs[Index + 1].Range.ListFormat.ListLevelNumber != 2)
          return "Falas(chon level 2)";
        if (d.Paragraphs[Index + 4].Range.ListFormat.ListLevelNumber != 2)
          return "Falas(chon level 2)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau16(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Three-String") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        if (d.Paragraphs[Index + 1].Range.ListFormat.ListLevelNumber != 2)
          return "Falas(chon level 2)";
        if (d.Paragraphs[Index + 4].Range.ListFormat.ListLevelNumber != 2)
          return "Falas(chon level 2)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau17(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Mandolin Style") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        if (d.Paragraphs[Index].Range.ListFormat.ListValue != 11)
          return "Falas(dung format painter)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau18(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Bass") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        if (d.Paragraphs[Index].Range.ListFormat.ListValue != 5)
          return "Falas(click chuot phai chon continue q)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau19(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Standard") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        if (d.Paragraphs[Index].Range.ListFormat.ListString != "1)")
          return "Falas(numbering kieu 1)...)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau20(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Standard") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        if (d.Paragraphs[Index].Range.ListFormat.ListValue != 500)
          return "Falas(click chuot phai chon set value)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau21(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Thimos1000.wordpress.com") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        if ((double) d.Paragraphs[Index].Range.ParagraphFormat.LineSpacing != 18.0)
          return "Falas(Thimos1000.wordpress.com)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau22(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Standard") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Text was modified)";
        if (d.Paragraphs[Index].Range.ListFormat.ListType != WdListType.wdListPictureBullet)
          return "Falas(bullet bang picture)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau23(Application a, _Document d)
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

    private static string cau24(Application a, _Document d)
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

    private static string cau25(Application a, _Document d)
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

    private static string cau26(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Corporate events") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(không sửa nọi dung)";
        if (d.Paragraphs[Index].Range.ListFormat.ListString != "\uF0B7")
          return "Falas(bullet)";
        while (!d.Paragraphs[Index].Range.Text.Contains("Online events") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(không sửa nọi dung)";
        if (d.Paragraphs[Index].Range.ListFormat.ListString != "\uF0B7")
          return "Falas(bullet)";
        while (!d.Paragraphs[Index].Range.Text.Contains("weddings") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(không sửa nọi dung)";
        if (d.Paragraphs[Index].Range.ListFormat.ListString != "")
          return "Falas(không bullet weddings)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau27(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Velociraptor") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(không sửa nọi dung)";
        if (d.Paragraphs[Index].Range.ListFormat.ListLevelNumber != 3)
          return "Falas(chuyển Velociraptor level 3)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau28(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("margie@margiestravel.com") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(không chỉnh sửa văn bản)";
        if (d.Paragraphs[Index].Range.ParagraphFormat.LineSpacing.ToString() != "14")
          return "Falas(exactly 14 pt)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau29(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (!d.Paragraphs[Index].Range.Text.Contains("Algebra II") && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(không chỉnh sửa văn bản)";
        if (d.Paragraphs[Index].Range.ListFormat.ListString != "111.")
          return "Falas(111.)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }
  }
}
