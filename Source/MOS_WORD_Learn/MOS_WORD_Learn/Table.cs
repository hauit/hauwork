// Decompiled with JetBrains decompiler
// Type: MOS_WORD_LEARN.Table
// Assembly: MOS_WORD_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.Office.Interop.Word;
using System;

namespace MOS_WORD_LEARN
{
  internal class Table
  {
    public static string CheckCau(int causo, Application a, _Document d)
    {
      switch (causo)
      {
        case 1:
          return Table.cau33(a, d);
        case 2:
          return Table.cau28(a, d);
        case 3:
          return Table.cau29(a, d);
        case 4:
          return Table.cau3(a, d);
        case 5:
          return Table.cau26(a, d);
        case 6:
          return Table.cau31(a, d);
        case 7:
          return Table.cau32(a, d);
        case 8:
          return Table.cau30(a, d);
        case 9:
          return Table.cau27(a, d);
        case 10:
          return Table.cau12(a, d);
        case 11:
          return Table.cau48(a, d);
        case 12:
          return Table.cau12(a, d);
        case 13:
          return Table.cau13(a, d);
        case 14:
          return Table.cau14(a, d);
        case 15:
          return Table.cau15(a, d);
        case 16:
          return Table.cau16(a, d);
        case 17:
          return Table.cau17(a, d);
        case 18:
          return Table.cau18(a, d);
        case 19:
          return Table.cau19(a, d);
        case 20:
          return Table.cau20(a, d);
        case 21:
          return Table.cau21(a, d);
        case 22:
          return Table.cau22(a, d);
        case 23:
          return Table.cau23(a, d);
        case 24:
          return Table.cau24(a, d);
        case 25:
          return Table.cau25(a, d);
        case 26:
          return Table.cau26(a, d);
        case 27:
          return Table.cau27(a, d);
        case 28:
          return Table.cau28(a, d);
        case 29:
          return Table.cau29(a, d);
        case 30:
          return Table.cau30(a, d);
        case 31:
          return Table.cau31(a, d);
        case 32:
          return Table.cau32(a, d);
        case 33:
          return Table.cau33(a, d);
        case 34:
          return Table.cau34(a, d);
        default:
          return "Default Table";
      }
    }

    private static string cau1(Application a, _Document d)
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
        return "False";
      }
      return "True";
    }

    private static string cau2(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
          return "False(not insert or delete table)";
        if (d.Tables[1].Rows[d.Tables[1].Rows.Count].Cells.Count != 1)
          return "False(Merge last row)";
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
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Columns.Width.ToString() != "89.3")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau4(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
          return "False(convert text to table)";
        if (d.Tables[1].Rows.Count != 4)
          return "False(Wrong text)";
        if ((double) d.Tables[1].Columns.Width > 300.0)
          return "False (fix conten)";
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
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Registration Dates"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False (heading was edited)";
        if (d.Paragraphs[Index + 1].Range.Tables.Count != 1)
          return "False (place table wrong place)";
        // ISSUE: variable of a compiler-generated type
        Microsoft.Office.Interop.Word.Table table = d.Paragraphs[Index + 1].Range.Tables[1];
        if (table.Columns.Count != 3)
          return "False(3 cols must use tab)";
        if (table.Rows.Count != 7)
          return "False(rows)";
        if ((double) table.PreferredWidth != 9999999.0)
          return "False(Fit Windows)";
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
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Game Locations and Times"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False (heading was edited)";
        if (d.Paragraphs[Index + 1].Range.Tables.Count != 1)
          return "False (place table wrong place)";
        // ISSUE: variable of a compiler-generated type
        Microsoft.Office.Interop.Word.Table table = d.Paragraphs[Index + 1].Range.Tables[1];
        if (table.Rows[2].Cells[1].Range.Text != "Softball\r\a")
          return "False(sort)";
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
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Game Locations and Times"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False (heading was edited)";
        if (d.Paragraphs[Index + 1].Range.Tables.Count != 1)
          return "False (place table wrong place)";
        // ISSUE: variable of a compiler-generated type
        Microsoft.Office.Interop.Word.Table table = d.Paragraphs[Index + 1].Range.Tables[1];
        if (table.Rows.HeadingFormat != 9999999)
          return "False(Heading repeat)";
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
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("The table below lists the summary counts."))
          ++Index;
        if (Index >= d.Paragraphs.Count)
          return "False (Text was edited)";
        if (d.Paragraphs[Index + 1].Range.Tables.Count != 1)
          return "False (place table wrong place)";
        // ISSUE: variable of a compiler-generated type
        Microsoft.Office.Interop.Word.Table table = d.Paragraphs[Index + 1].Range.Tables[1];
        if (table.Rows[9].Cells.Count != 1)
          return "False(merge last row)";
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
        if (d.Tables.Count != 1)
          return "False(not insert or delete table)";
        if (d.Tables[1].Columns.Width.ToString() != "116.9")
          return "False(Distribute column width)";
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
        if (d.Tables.Count != 1)
          return "False(convert text to table)";
        if (d.Tables[1].Rows.Count != 4)
          return "False(Wrong text)";
        if ((double) d.Tables[1].Columns.Width > 300.0)
          return "False (fix conten)";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau11(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
          return "False(not insert or delete table)";
        if (d.Tables[1].Title != "Book Selections")
          return "False()";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau12(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Borders[WdBorderType.wdBorderBottom].Color.ToString() != "-721354906")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau13(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Rows[9].Cells.Count != 1)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau14(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
        if ((double) d.Tables[1].Columns.Width == 9999999.0)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau15(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[7].Range.Tables.Count != 1)
                    return "False";
                if (d.Paragraphs[7].Range.Tables[1].Columns.Count != 5)
                    return "False";
                if (d.Paragraphs[7].Range.Tables[1].Rows.Count != 9)
                    return "False";
                if ((double) d.Paragraphs[7].Range.Tables[1].Columns.Width == 9999999.0)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau16(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau17(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Title != "Species Count")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau18(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Title != "Species")
                    return "False";
                if (d.Tables[1].Descr != "Count of Species")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau19(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Rows.HeadingFormat != 9999999)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau20(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Columns[5].Cells.Count != 1)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau21(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if ((double) d.Tables[1].Rows.Height == 9999999.0)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau22(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (!d.Tables[1].Rows[8].Cells[4].Range.Text.Contains("27"))
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau23(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 0)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau24(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if ((double) d.Tables[1].Columns[1].Width > 110.0)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau25(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Rows.Count != 8)
                    return "False";
                if (d.Tables[1].Columns.Count != 4)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau26(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Rows[1].Cells.Count != 1)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau27(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 2)
                    return "False";
                if (d.Tables[1].Spacing.ToString() != "1.4")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau28(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Columns.Count != 2)
                    return "False";
                if (d.Tables[1].Rows.Count != 5)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau29(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 2)
                    return "False";
                if (d.Tables[1].Columns.Width.ToString() != "158.4")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau30(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Range.Text != "Geologic period \r\aDinosaur\r\a\r\aBooks\r\a1\r\a\r\aHighlighter\r\a2 colors\r\a\r\aMagazines\r\a3\r\a\r\aNotebooks\r\a1\r\a\r\aPaper pads\r\a1 \r\a\r\aPencils\r\a2\r\a\r\aPens\r\a3\r\a\r\aScissors\r\a1 pair\r\a\r\a")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau31(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Rows[1].Cells.Count != 1)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau32(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Rows.HeadingFormat != 9999999)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau33(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Rows.Count != 6)
                    return "False";
                if (d.Tables[1].Columns.Count != 2)
                    return "False";
                if (!d.Tables[1].Rows[1].Cells[1].Range.Text.ToLower().Contains("grade"))
                    return "False";
                if (!d.Tables[1].Rows[1].Cells[2].Range.Text.ToLower().Contains("score range"))
                    return "False";
                if (!d.Tables[1].AllowAutoFit)
                    return "False";
                if ((double) d.Tables[1].Columns[1].Width >= (double) d.Tables[1].Columns[2].Width)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau34(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Rows.Count != 6)
                    return "False";
                if (d.Tables[1].Columns.Count != 2)
                    return "False";
                if (!d.Tables[1].Rows[1].Cells[1].Range.Text.ToLower().Contains("grade"))
                    return "False";
                if (!d.Tables[1].Rows[1].Cells[2].Range.Text.ToLower().Contains("score range"))
                    return "False";
                if (!d.Tables[1].AllowAutoFit)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau48(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Rows[1].Cells.Count != 1)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }
  }
}
