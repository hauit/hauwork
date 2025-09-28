// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.CheckWork6
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
{
    public static class CheckWork6
    {
        public static string Check(int cauLon, Application a, Document d)
        {
            switch (cauLon)
            {
                case 0:
                    return CheckWork6.Cau0(a, d);
                case 1:
                    return CheckWork6.Cau1(a, d);
                case 2:
                    return CheckWork6.Cau2(a, d);
                case 3:
                    return CheckWork6.Cau3(a, d);
                case 4:
                    return CheckWork6.Cau4(a, d);
                case 5:
                    return CheckWork6.Cau5(a, d);
                case 6:
                    return CheckWork6.Cau6(a, d);
                case 7:
                    return CheckWork6.Cau7(a, d);
                case 8:
                    return CheckWork6.Cau8(a, d);
                case 9:
                    return CheckWork6.Cau9(a, d);
                case 10:
                    return CheckWork6.Cau10(a, d);
                case 11:
                    return CheckWork6.Cau11(a, d);
                case 12:
                    return CheckWork6.Cau12(a, d);
                case 13:
                    return CheckWork6.Cau13(a, d);
                case 14:
                    return CheckWork6.Cau14(a, d);
                case 15:
                    return CheckWork6.Cau15(a, d);
                case 16:
                    return CheckWork6.Cau16(a, d);
                case 17:
                    return CheckWork6.Cau17(a, d);
                case 18:
                    return CheckWork6.Cau18(a, d);
                case 19:
                    return CheckWork6.Cau19(a, d);
                case 20:
                    return CheckWork6.Cau20(a, d);
                case 21:
                    return CheckWork6.Cau21(a, d);
                case 22:
                    return CheckWork6.Cau22(a, d);
                case 23:
                    return CheckWork6.Cau23(a, d);
                case 24:
                    return CheckWork6.Cau24(a, d);
                case 25:
                    return CheckWork6.Cau25(a, d);
                case 26:
                    return CheckWork6.Cau26(a, d);
                case 27:
                    return CheckWork6.Cau27(a, d);
                case 28:
                    return CheckWork6.Cau28(a, d);
                case 29:
                    return CheckWork6.Cau29(a, d);
                default:
                    return "";
            }
        }

        private static string Cau0(Application a, Document d)
        {
            try
            {
                if (d.Sections.Count != 1)
                    return "False (number of section)";
                if ((double)d.Sections[1].PageSetup.TopMargin != 54.0)
                    return "False(lề trên)";
                if ((double)d.Sections[1].PageSetup.BottomMargin != 54.0)
                    return "False(lề dưới)";
                if ((double)d.Sections[1].PageSetup.LeftMargin != 36.0)
                    return "False(lề trái)";
                if ((double)d.Sections[1].PageSetup.RightMargin != 36.0)
                    return "False(lề phải)";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string Cau1(Application a, Document d) => "True";

        private static string Cau2(Application a, Document d)
        {
            try
            {
                if (d.Tables.Count != 2)
                    return "False(không thêm xóa table)";
                if (d.Tables[1].Spacing.ToString() != "1.4")
                    return "False(0.02 inch)";
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
                int Index1 = 0;
                for (int Index2 = 1; Index2 < d.Paragraphs.Count; ++Index2)
                {
                    if (d.Paragraphs[Index2].Range.Text.Contains("Payment may be withheld under the following circumstances"))
                    {
                        Index1 = Index2;
                        break;
                    }
                }
                if (Index1 == 0)
                    return "False(khồng chỉnh sửa văn bảng)";
                if (d.Paragraphs[Index1].Range.Font.Bold == -1)
                    return "False(clear format)";
                if (d.Paragraphs[Index1].Range.Font.Italic == -1)
                    return "False(clear format)";
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

        private static string Cau5(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("Organize") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Text was modified)";
                return !d.Paragraphs[1].Range.Text.Contains("Prep Your") ? "False(import text)" : "True";
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
                int Index = 1;
                while (d.Paragraphs[Index].Range.Text != "Ingredients for Beef or Turkey Stew\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Ingredients for Beef or Turkey Stew)";
                if (d.Paragraphs[Index + 1].Range.Text.Contains("\a"))
                    return "False(Table Ingredients for Beef or Turkey Stew)";
                while (d.Paragraphs[Index].Range.Text != "Ingredients for Spanish omelet\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Ingredients for Spanish omelet)";
                return d.Paragraphs[Index + 1].Range.Text.Contains("\a") ? "False(Table Ingredients for Spanish omelet)" : "True";
            }
            catch (Exception ex)
            {
                return "False(something wrong)";
            }
        }

        private static string Cau7(Application a, Document d)
        {
            try
            {
                return d.Sections[2].PageSetup.TextColumns.Count != 2 ? "False(section 2 <> 2 column)" : "True";
            }
            catch (Exception ex)
            {
                return "False(something wrong)";
            }
        }

        private static string Cau8(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (d.Paragraphs[Index].Range.Text != "Ingredients for Beef or Turkey Stew\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Ingredients for Beef or Turkey Stew)";
                // ISSUE: variable of a compiler-generated type
                ListFormat listFormat1 = d.Paragraphs[Index + 2].Range.ListFormat;
                if (listFormat1.List == null)
                    return "False(Not bullet on Ingredients for Beef list)";
                while (d.Paragraphs[Index].Range.Text != "Ingredients for Spanish omelet\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Ingredients for Spanish omelet)";
                // ISSUE: variable of a compiler-generated type
                ListFormat listFormat2 = d.Paragraphs[Index + 2].Range.ListFormat;
                return listFormat2.List == null ? "False(Not bullet on Ingredients for Spanish omelet list)" : "True";
            }
            catch (Exception ex)
            {
                return "False(Not Bulet)";
            }
        }

        private static string Cau9(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (d.Paragraphs[Index].Range.Text != "Vegetables\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Vegetables)";
                // ISSUE: variable of a compiler-generated type
                ListFormat listFormat1 = d.Paragraphs[Index].Range.ListFormat;
                if (d.Paragraphs[Index].Range.Font.Bold != -1)
                    return "False(B)";
                if (d.Paragraphs[Index].Range.Font.Italic != -1)
                    return "False(I)";
                if (listFormat1.List == null)
                    return "False(Not bullet on Vegetables)";
                if (listFormat1.ListLevelNumber != 1)
                    return "False(Vegetables on level 1)";
                // ISSUE: variable of a compiler-generated type
                ListFormat listFormat2 = d.Paragraphs[Index + 1].Range.ListFormat;
                if (listFormat2.List == null)
                    return "False(Not bullet on level 2)";
                if (listFormat2.ListLevelNumber != 2)
                    return "False(level 2)";
                // ISSUE: variable of a compiler-generated type
                ListFormat listFormat3 = d.Paragraphs[Index + 2].Range.ListFormat;
                if (listFormat3.List == null)
                    return "False(Not bullet on level 2)";
                if (listFormat3.ListLevelNumber != 2)
                    return "False(level 2)";
                while (d.Paragraphs[Index].Range.Text != "Fruits\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Fruits)";
                // ISSUE: variable of a compiler-generated type
                ListFormat listFormat4 = d.Paragraphs[Index].Range.ListFormat;
                if (listFormat4.List == null)
                    return "False(Not bullet on Fruits)";
                if (listFormat4.ListLevelNumber != 1)
                    return "False(Fruits on level 1)";
                if (d.Paragraphs[Index].Range.Font.Bold != -1)
                    return "False(B)";
                if (d.Paragraphs[Index].Range.Font.Italic != -1)
                    return "False(I)";
                // ISSUE: variable of a compiler-generated type
                ListFormat listFormat5 = d.Paragraphs[Index + 1].Range.ListFormat;
                if (listFormat5.List == null)
                    return "False(Not bullet on level 2)";
                if (listFormat5.ListLevelNumber != 2)
                    return "False(level 2)";
                return listFormat5.ListString != "2.1." ? "False(Type of multiList)" : "True";
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
                object Index1 = (object)"Picture 1";
                if (d.Shapes[ref Index1].WrapFormat.Type != WdWrapType.wdWrapThrough)
                {
                    object Index2 = (object)"Picture 1";
                    return "False(" + (object)d.Shapes[ref Index2].WrapFormat.Type + ")";
                }
                object Index3 = (object)"Picture 1";
                if (d.Shapes[ref Index3].Left.ToString() != "28.8")
                    return "False(H)";
                object Index4 = (object)"Picture 1";
                return d.Shapes[ref Index4].RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin ? "False(Margin)" : "True";
            }
            catch (Exception ex)
            {
                return "False(SmartArt must Inline with text)";
            }
        }

        private static string Cau11(Application a, Document d)
        {
            try
            {
                if (d.Footnotes.Count != 1)
                    return "False(insert footnote)";
                return d.Footnotes[1].Range.Text != "" ? "True" : "False(insert text)";
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
                    return "False(inser 1 hyperlink)";
                object Index1 = (object)1;
                if (!(d.Hyperlinks[ref Index1].SubAddress != "_Cook_ahead"))
                    return "True";
                object Index2 = (object)1;
                return "False(" + d.Hyperlinks[ref Index2].SubAddress + ")";
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
                foreach (Table table in d.Tables)
                {
                    if (table.Rows.Count == 15)
                    {
                        if (table.PreferredWidthType != WdPreferredWidthType.wdPreferredWidthPercent)
                            return "False(Autofit content)";
                        if (table.AutoFormatType != 1)
                            return "False(AutoFormatType)";
                        if (table.Columns.Count != 2)
                            return "False(col <>2)";
                        if (table.Rows[1].Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                            return "False (row 1 center)";
                        if (table.Rows[1].Cells.Count != 1)
                            return "False(M row 1)";
                        if (table.Rows[9].Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                            return "False (row 9 center)";
                        if (table.Rows[9].Cells.Count != 1)
                            return "False(M row 9)";
                        if (table.Rows[13].Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                            return "False (row 13 center)";
                        return table.Rows[13].Cells.Count != 1 ? "False(M row 13)" : "True";
                    }
                }
                return "False(conver table)";
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
                string str = "True";
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("Table 1-Instant Pot Mashed Potatoes") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Table 1-Instant Pot Mashed Potatoes)";
                return d.Paragraphs[Index - 1].Range.Text != "\r\a" ? "False(below table)" : str;
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string Cau15(Application a, Document d)
        {
            try
            {
                if (d.Bookmarks.Count != 1)
                    return "False(Insert Bookmark)";
                object Index1 = (object)1;
                if (!(d.Bookmarks[ref Index1].Name != "Scoring"))
                    return "True";
                object Index2 = (object)1;
                return "False(" + d.Bookmarks[ref Index2].Name + ")";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string Cau16(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("Minimum of ingredients to prepare recipes") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Minimum of ingredients to prepare recipes)";
                if (d.Paragraphs[Index + 1].Range.InlineShapes.Count != 1)
                    return "False (number of SmartArt)";
                if (d.Paragraphs[Index + 1].Range.InlineShapes[1].Type != WdInlineShapeType.wdInlineShapeSmartArt)
                    return "False (Not SmartArt)";
                if (d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Layout.Name != "Vertical Bullet List")
                    return "False (" + d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Layout.Name + ")";
                if (d.Paragraphs[Index + 1].Range.InlineShapes[1].Width.ToString() != "283.65")
                    return "False(W)";
                return d.Paragraphs[Index + 1].Range.InlineShapes[1].Height.ToString() != "113.75" ? "False(H)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string Cau17(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("Figure 1-Some contain iron, magnesium, silica, or") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False (Figure 1-Some contain iron, magnesium, silica, or aluminum)";
                if (d.Paragraphs[Index - 2].Range.Text != "Minimum of ingredients to prepare recipes\r")
                    return "False(below smartArt)";
                return d.Paragraphs[Index - 1].Range.Text != "/\r" ? "False (below smartArt)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string Cau18(Application a, Document d)
        {
            try
            {
                object Index = (object)"Text Box 2";
                // ISSUE: variable of a compiler-generated type
                Shape shape = d.Shapes[ref Index];
                if (shape.RelativeHorizontalSize != WdRelativeHorizontalSize.wdRelativeHorizontalSizeMargin)
                    return "False(Horizontal)";
                if (shape.RelativeVerticalSize != WdRelativeVerticalSize.wdRelativeVerticalSizeMargin)
                    return "False(Vertical)";
                if ((double)shape.Top != -999997.0)
                    return "False(Top)";
                if ((double)shape.Left != -999998.0)
                    return "False(Left)";
                if (shape.RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin || shape.RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin)
                    return "False(Margin)";
                return shape.TextFrame.TextRange.Text.ToUpper() != "ANDESITE IS ONE OF THE MOST COMMON VOLCANIC ROCKS AND CAN CONTAIN OLIVINE.\r" ? "False(ANDESITE IS ONE OF THE MOST COMMON VOLCANIC ROCKS AND CAN CONTAIN OLIVINE.)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string Cau19(Application a, Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string Cau20(Application a, Document d)
        {
            try
            {
                string text = d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text;
                return !text.Contains("\r\r") ? "False(" + text + ")" : "True";
            }
            catch (Exception ex)
            {
                return "False (insert header)";
            }
        }

        private static string Cau21(Application a, Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string Cau22(Application a, Document d)
        {
            try
            {
                return d.ActiveWindow.ActivePane.View.Zoom.Percentage != 95 ? "False(" + (object)d.ActiveWindow.ActivePane.View.Zoom.Percentage + ")" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string Cau23(Application a, Document d)
        {
            try
            {
                string str = "True";
                return a.Options.SaveInterval != 8 || !d.EmbedTrueTypeFonts ? "False" : str;
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string Cau24(Application a, Document d)
        {
            string str = "True";
            return a.Options.SaveInterval != 13 || !d.EmbedTrueTypeFonts ? "False" : str;
        }

        private static string Cau25(Application a, Document d)
        {
            string str = "False";
            foreach (Table table in d.Tables)
            {
                if (table.Columns.Count == 8)
                {
                    if (table.PreferredWidthType != WdPreferredWidthType.wdPreferredWidthAuto
                        || table.Cell(1, 1).Range.Text != "Michigan 2016 Region 4 Aesop Gone Viral Div 1\r\a"
                        || table.AutoFormatType != 1
                        || table.Rows.Count != 10)
                        return "False";
                    if ((int)table.Cell(1, 1).Width != 506
                        || table.Borders.Count != 8
                        || table.Borders[WdBorderType.wdBorderBottom].ColorIndex != WdColorIndex.wdGray25)
                        return "False";
                    // Fixed: Use type-safe cast for TableStyle
                    var tableStyle = table.get_Style() as TableStyle;
                    // Optionally, check for a specific style name if needed
                    // if (tableStyle == null || tableStyle.NameLocal != "ExpectedStyleName")
                    //     return "False";
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
