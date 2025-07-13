// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.PageSetup
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.Office.Interop.Word;
using System;

namespace WindowsFormsApplication1
{
  internal class PageSetup
  {
    public static string CheckCau(int causo, Application a, _Document d)
    {
      switch (causo)
      {
        case 1:
          return PageSetup.cau0(a, d);
        case 2:
          return PageSetup.cau16(a, d);
        case 3:
          return PageSetup.cau4(a, d);
        case 4:
          return PageSetup.cau7(a, d);
        case 5:
          return PageSetup.cau8(a, d);
        case 6:
          return PageSetup.cau3(a, d);
        case 7:
          return PageSetup.cau15(a, d);
        case 8:
          return PageSetup.cau14(a, d);
        case 9:
          return PageSetup.cau6(a, d);
        case 10:
          return PageSetup.cau11(a, d);
        case 11:
          return PageSetup.cau13(a, d);
        case 12:
          return PageSetup.chiaCot(a, d);
        case 13:
          return PageSetup.cau12(a, d);
        case 14:
          return PageSetup.cau13(a, d);
        case 15:
          return PageSetup.cau14(a, d);
        case 16:
          return PageSetup.cau15(a, d);
        case 17:
          return PageSetup.cau16(a, d);
        default:
          return "False";
      }
    }

    private static string cau0(Application a, _Document d)
    {
      try
      {
        if (!a.ActiveWindow.View.ShowAll)
          return "False";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau1(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (d.Paragraphs[Index].Range.Text != "\u000EFour-String\r" && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Column break before 'Four-String')";
        if (d.Paragraphs[Index].Range.Text != "\u000EFour-String\r")
          return "False()";
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
        if (d.Sections[2].PageSetup.TextColumns.Count != 2)
          return "False(section2 <>2 Column)";
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
        int Index = 1;
        while (d.Paragraphs[Index].Range.Text != "\u000EFour-String\r" && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Column break before 'Four-String')";
        if (d.Paragraphs[Index].Range.Text != "\u000EFour-String\r")
          return "False()";
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
        int Index = 1;
        while (d.Paragraphs[Index].Range.Text != "Vegetables\r" && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False('Vegetables' not found)";
        if (d.Paragraphs[Index - 1].Range.Text != "\f")
          return "False(Break section)";
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
        if (d.Sections[2].PageSetup.TextColumns.Count != 2)
          return "False(section2 <>2 Column)";
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
        if (d.Sections[2].PageSetup.Orientation != WdOrientation.wdOrientLandscape)
          return "False(Landscape section 2)";
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
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Tuning"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Tieu de Tuning khong tim thay)";
        if (d.Paragraphs[Index - 1].Range.Text != "\f\r")
          return "False(chen chi ngat trang (pageBreak))";
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
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Barstow College"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Tieu de Barstow College khong tim thay)";
        if (d.Paragraphs[Index].Range.Text != "Barstow College \vObservation Project\r")
          return "False(chen chi ngat dong xac tu Observation)";
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
        if (d.Sections.Count != 3)
          return "False (number of section)";
        if (d.Sections[2].PageSetup.TextColumns.Count != 2)
          return "False(section2 <>2 Column)";
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
        int Index = 1;
        while (d.Paragraphs[Index].Range.Text != "\u000EFour-String\r" && Index < d.Paragraphs.Count)
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False(Column break before 'Four-String')";
        if (d.Paragraphs[Index].Range.Text != "\u000EFour-String\r")
          return "False()";
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
        if (d.Sections.Count != 3)
          return "False (number of section)";
        float leftMargin = d.Sections[1].PageSetup.LeftMargin;
        if (leftMargin.ToString() != "54")
          return "False(apply kieu trong Build In)";
        leftMargin = d.Sections[3].PageSetup.LeftMargin;
        if (leftMargin.ToString() != "54")
          return "False(apply cho toan bo cac section)";
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
        if (d.Sections.Count != 2)
          return "False (number of section)";
        if (d.Sections[2].PageSetup.Orientation != WdOrientation.wdOrientLandscape)
          return "False(chinh huong giay ngan cho section 2)";
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
        if (d.Sections.Count != 1)
          return "False (number of section)";
        if ((double) d.Sections[1].PageSetup.TopMargin != 54.0)
          return "False(lề trên)";
        if ((double) d.Sections[1].PageSetup.BottomMargin != 54.0)
          return "False(lề dưới)";
        if ((double) d.Sections[1].PageSetup.LeftMargin != 36.0)
          return "False(lề trái)";
        if ((double) d.Sections[1].PageSetup.RightMargin != 36.0)
          return "False(lề phải)";
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
        if (d.Sections.Count != 3)
          return "False (number of section)";
        if (d.Sections[2].PageSetup.Orientation != WdOrientation.wdOrientLandscape)
          return "False(section 2 hường giấy ngan)";
        if (d.Sections[1].PageSetup.Orientation != WdOrientation.wdOrientPortrait)
          return "False(section 1 hường giấy đứng)";
        if (d.Sections[3].PageSetup.Orientation != WdOrientation.wdOrientPortrait)
          return "False(section 3 hường giấy đứng)";
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
        if (d.Sections.Count != 3)
          return "False (number of section)";
        if (d.Sections[2].PageSetup.TextColumns.Count != 2)
          return "False(section2 <>2 Column)";
        if (d.Sections[2].PageSetup.TextColumns.Spacing.ToString() != "21.6")
          return "False(sai khoảng cách giửa 2 cột)";
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

    private static string chiaCot(Application a, _Document d)
    {
      try
      {
        if (d.Sections.Count != 3)
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
  }
}
