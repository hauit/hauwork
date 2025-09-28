// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.CheckWork4
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
{
    public static class CheckWork4
    {
        public static string Check(int cauLon, Application a, Document d)
        {
            switch (cauLon)
            {
                case 0:
                    return CheckWork4.Cau5(a, d);
                case 1:
                    return CheckWork4.Cau1(a, d);
                case 2:
                    return CheckWork4.Cau2(a, d);
                case 3:
                    return CheckWork4.Cau3(a, d);
                case 4:
                    return CheckWork4.Cau4(a, d);
                case 5:
                    return CheckWork4.Cau0(a, d);
                case 6:
                    return CheckWork4.Cau6(a, d);
                case 7:
                    return CheckWork4.Cau7(a, d);
                case 8:
                    return CheckWork4.Cau8(a, d);
                case 9:
                    return CheckWork4.Cau9(a, d);
                case 10:
                    return CheckWork4.Cau10(a, d);
                case 11:
                    return CheckWork4.Cau11(a, d);
                case 12:
                    return CheckWork4.Cau12(a, d);
                case 13:
                    return CheckWork4.Cau13(a, d);
                case 14:
                    return CheckWork4.Cau14(a, d);
                case 15:
                    return CheckWork4.Cau15(a, d);
                case 16:
                    return CheckWork4.Cau16(a, d);
                case 17:
                    return CheckWork4.Cau17(a, d);
                case 18:
                    return CheckWork4.Cau18(a, d);
                case 19:
                    return CheckWork4.Cau19(a, d);
                case 20:
                    return CheckWork4.Cau20(a, d);
                case 21:
                    return CheckWork4.Cau21(a, d);
                case 22:
                    return CheckWork4.Cau22(a, d);
                case 23:
                    return CheckWork4.Cau23(a, d);
                case 24:
                    return CheckWork4.Cau24(a, d);
                case 25:
                    return CheckWork4.Cau25(a, d);
                case 26:
                    return CheckWork4.Cau26(a, d);
                case 27:
                    return CheckWork4.Cau27(a, d);
                case 28:
                    return CheckWork4.Cau28(a, d);
                case 29:
                    return CheckWork4.Cau29(a, d);
                default:
                    return "";
            }
        }

        private static string Cau0(Application a, Document d)
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

        private static string Cau1(Application a, Document d)
        {
            try
            {
                int num = -1;
                for (int index = 1; index < d.Shapes.Count; ++index)
                {
                    object Index = (object)index;
                    if (d.Shapes[ref Index].Name.Contains("Horizontal Scroll"))
                        num = index;
                }
                if (num == -1)
                    return "False(chèn Horizontal Scroll)";
                object Index1 = (object)num;
                if (d.Shapes[ref Index1].TextFrame.TextRange.Text.Trim() != "Remember your calculator!")
                    return "Fales(Remember your calculator!)";
                object Index2 = (object)num;
                if (d.Shapes[ref Index2].WrapFormat.Type != WdWrapType.wdWrapSquare)
                    return "False(Square)";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
            return "True";
        }

        private static string Cau2(Application a, Document d)
        {
            try
            {
                object Index = (object)"Picture 1";
                if (d.Shapes[ref Index].Line.ForeColor.RGB != 11957550)
                    return "False(sai màu)";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
            return "True";
        }

        private static string Cau3(Application a, Document d)
        {
            try
            {
                if (d.InlineShapes[1].SmartArt.Reverse == MsoTriState.msoTrue)
                    return "False(Reverse)";
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
                if (d.Tables.Count != 1)
                    return "False(chèn 1 table)";
                if (d.Tables[1].Rows.Count != 6)
                    return "False(6 dòng)";
                if (d.Tables[1].Columns.Count != 2)
                    return "False(2 cột)";
                if (!d.Tables[1].Rows[1].Cells[1].Range.Text.ToLower().Contains("grade"))
                    return "False(Grade)";
                if (!d.Tables[1].Rows[1].Cells[2].Range.Text.ToLower().Contains("score range"))
                    return "False(Score Range)";
                if (!d.Tables[1].AllowAutoFit)
                    return "False(auto fix conten)";
                if ((double)d.Tables[1].Columns[1].Width >= (double)d.Tables[1].Columns[2].Width)
                    return "False(auto fix conten)";
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
                if (d.Bookmarks.Count == 0)
                    return "False (add bookmark)";
                if (d.Bookmarks.Count > 1)
                    return "False (chỉ add 1 bookmark thôi)";
                object Index = (object)1;
                if (!d.Bookmarks[ref Index].Name.Contains("Cheating"))
                    return "False (Cheating)";
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
                string str = "True";
                return d.Sections[2].PageSetup.TextColumns.Count != 2 ? "False (section 2: Column<>2 )" : str;
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
                string str = "True";
                int Index = 1;
                while (d.Paragraphs[Index].Range.Text != "Fissure vent\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Fissure vent)";
                // ISSUE: variable of a compiler-generated type
                ListFormat listFormat1 = d.Paragraphs[Index + 1].Range.ListFormat;
                if (listFormat1.ListPictureBullet.AlternativeText != "*")
                    return "False(Fissure vent: AlternativeText)";
                if (!listFormat1.ListPictureBullet.IsPictureBullet)
                    return "False(Fissure vent: PictureBullet)";
                if ((double)listFormat1.ListTemplate.ListLevels[1].NumberPosition != 36.0)
                    return "False(Fissure vent: NumberPosition)";
                if ((double)listFormat1.ListTemplate.ListLevels[1].TextPosition != 54.0)
                    return "False(Fissure vent: TextPosition)";
                while (d.Paragraphs[Index].Range.Text != "Shield volcano\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Shield volcano)";
                // ISSUE: variable of a compiler-generated type
                ListFormat listFormat2 = d.Paragraphs[Index + 1].Range.ListFormat;
                if (listFormat2.ListPictureBullet.AlternativeText != "*")
                    return "False(Shield volcano: AlternativeText)";
                if (!listFormat2.ListPictureBullet.IsPictureBullet)
                    return "False(Shield volcano: PictureBullet)";
                if ((double)listFormat2.ListTemplate.ListLevels[1].NumberPosition != 36.0)
                    return "False(Shield volcano: NumberPosition)";
                if ((double)listFormat2.ListTemplate.ListLevels[1].TextPosition != 54.0)
                    return "False(Shield volcano: TextPosition)";
                while (d.Paragraphs[Index].Range.Text != "Lava dome\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Lava dome)";
                // ISSUE: variable of a compiler-generated type
                ListFormat listFormat3 = d.Paragraphs[Index + 1].Range.ListFormat;
                if (listFormat3.ListPictureBullet.AlternativeText != "*")
                    return "False(Lava dome: AlternativeText)";
                if (!listFormat3.ListPictureBullet.IsPictureBullet)
                    return "False(Lava dome: PictureBullet)";
                if ((double)listFormat3.ListTemplate.ListLevels[1].NumberPosition != 36.0)
                    return "False(Lava dome: NumberPosition)";
                if ((double)listFormat3.ListTemplate.ListLevels[1].TextPosition != 54.0)
                    return "False(Lava dome: TextPosition)";
                while (d.Paragraphs[Index].Range.Text != "Lava dome\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Strato-volcano)";
                // ISSUE: variable of a compiler-generated type
                ListFormat listFormat4 = d.Paragraphs[Index + 1].Range.ListFormat;
                if (listFormat4.ListPictureBullet.AlternativeText != "*")
                    return "False(Strato-volcano: AlternativeText)";
                if (!listFormat4.ListPictureBullet.IsPictureBullet)
                    return "False(Strato-volcano: PictureBullet)";
                if ((double)listFormat4.ListTemplate.ListLevels[1].NumberPosition != 36.0)
                    return "False(Strato-volcano: NumberPosition)";
                return (double)listFormat4.ListTemplate.ListLevels[1].TextPosition != 54.0 ? "False(Strato-volcano: TextPosition)" : str;
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
                while (d.Paragraphs[Index].Range.Text != "Components of a volcanic eruption\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Components of a volcanic eruption)";
                for (; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text == "Airborne\r\a" || d.Paragraphs[Index].Range.Text == "Lava\r\a")
                        return "False(convert table to text)";
                }
                return "True";
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
                string str = "True";
                if (d.Paragraphs[1].DropCap.Position != WdDropPosition.wdDropNormal)
                    return "False(Type)";
                return d.Paragraphs[1].DropCap.LinesToDrop != 3 ? "False(line<>3)" : str;
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
                string str = "True";
                if (d.InlineShapes[1].SmartArt.Nodes.Count != 2)
                    return "False(remove 1 shape)";
                if (d.InlineShapes[1].SmartArt.Nodes[(object)1].TextFrame2.TextRange.Text != "Airborne")
                    return "False(Airborne)";
                if ((int)d.InlineShapes[1].Width != 360)
                    return "False(Width)";
                if ((int)d.InlineShapes[1].Height != 180)
                    return "False(Height)";
                if (d.InlineShapes[1].SmartArt.Nodes[(object)2].TextFrame2.TextRange.Text != "Earthbound")
                    return "False(Earthbound)";
                if (d.InlineShapes[1].SmartArt.Nodes[(object)2].Nodes.Count != 3)
                    return "False(shape 1 : 3 sub)";
                return d.InlineShapes[1].SmartArt.Nodes[(object)1].Nodes.Count != 4 ? "False(shape 2 : 4 sub)" : str;
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
                string str = "True";
                int Index = 1;
                while (d.Paragraphs[Index].Range.Text != "Figure 1- Volcanic ejecta by type\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Figure 1- Volcanic ejecta by type)";
                return d.Paragraphs[Index - 1].Range.Text != "/\r" ? "False(below SmartArt)" : str;
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
                string str = "True";
                int Index = d.Paragraphs.Count - 5;
                while (d.Paragraphs[Index].Range.Text.Contains("The 1980 eruption of Mt. Saint Helens (On the Insert tab)") && Index < d.Paragraphs.Count)
                    ++Index;
                return Index >= d.Paragraphs.Count ? "False(import text to end of document)" : str;
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
                string str = "False";
                foreach (Microsoft.Office.Interop.Word.Table table in d.Tables)
                {
                    if (table.Columns.Count == 6)
                    {
                        if (table.PreferredWidthType != WdPreferredWidthType.wdPreferredWidthAuto)
                            return "False(Autofit content)";
                        // ISSUE: reference to a compiler-generated method
                        if (table.Cell(2, 2).Range.Text != "B1\r\a")
                            return "False(sort)";
                        if (table.AutoFormatType != 1)
                            return "False(AutoFormatType)";
                        if (table.Rows.Count != 11)
                            return "False(row<>11)";
                        if (table.Borders.Count != 8 || table.Borders[WdBorderType.wdBorderBottom].ColorIndex != WdColorIndex.wdWhite)
                            return "False(Style)";
                        // ISSUE: reference to a compiler-generated method
                        // ISSUE: reference to a compiler-generated method
                        if (table.Cell(10, 2).Range.Text != "RAR\r\a" || table.Cell(11, 2).Range.Text != "AB\r\a")
                            return "False(sort)";
                        str = "True";
                    }
                }
                return str;
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
                while (d.Paragraphs[Index].Range.Text != "Table 1- Eruptions ranked by death toll\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Table 1- Eruptions ranked by death toll)";
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
                string str = "True";
                object Index = (object)"Picture 1";
                // ISSUE: variable of a compiler-generated type
                Microsoft.Office.Interop.Word.Shape shape = d.Shapes[ref Index];
                return shape.WrapFormat.Type != WdWrapType.wdWrapThrough || (double)shape.Left != -999996.0 || (double)shape.Top != -999997.0 || shape.RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin || shape.RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin ? "False" : str;
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
                string str = "True";
                for (int Index = 1; Index < d.Sections.Count; ++Index)
                {
                    if (d.Sections[Index].PageSetup.Orientation == WdOrientation.wdOrientLandscape)
                        return "False";
                }
                return d.Sections[d.Sections.Count].PageSetup.Orientation != WdOrientation.wdOrientLandscape ? "False" : str;
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
                object Index1 = (object)"Group 198";
                // ISSUE: variable of a compiler-generated type
                Microsoft.Office.Interop.Word.Shape shape = d.Shapes[ref Index1];
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
                object Index2 = (object)"Text Box 200";
                // ISSUE: variable of a compiler-generated type
                Microsoft.Office.Interop.Word.Shape groupItem = shape.GroupItems[ref Index2];
                return groupItem.TextFrame.TextRange.Text != "AN ESTIMATED 500 MILLION PEOPLE LIVE NEAR ACTIVE VOLCANOES.\r" ? "False(AN ESTIMATED 500 MILLION PEOPLE NEAR ACTIVE VOLCANOES.)" : "True";
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
                if (d.Footnotes.Count != 1)
                    return "False(number of footnote)";
                return d.Footnotes[1].Range.Text != "European Space Agency, 2009." ? "False(European Space Agency, 2009.)" : "True";
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
                string str = "True";
                // ISSUE: variable of a compiler-generated type
                Hyperlink hyperlink1 = (Hyperlink)null;
                foreach (Hyperlink hyperlink2 in d.Hyperlinks)
                {
                    if (hyperlink2.Name == "http://www.adatum.com/")
                        hyperlink1 = hyperlink2;
                }
                if (hyperlink1 == null)
                    return "False(http://www.adatum.com/)";
                return !hyperlink1.TextToDisplay.Contains("1500 active volcanoes") ? "False(1500 active volcanoes)" : str;
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
                string str = "True";
                if (d.Bookmarks.Count != 1)
                    return "False(number of bookmark)";
                object Index = (object)"Super";
                return !d.Bookmarks[ref Index].Range.Text.Contains("Super-volcano: the great devastator") ? "False(Super-volcano: the great devastator)" : str;
            }
            catch (Exception ex)
            {
                return "False (Super bookmart not found)";
            }
        }

        private static string Cau21(Application a, Document d)
        {
            try
            {
                string str = "True";
                return !a.ActiveWindow.ActivePane.DisplayRulers ? "False(show ruler)" : str;
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
                string str = "True";
                object documentProperties = d.BuiltInDocumentProperties;
                object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder)null, documentProperties, new object[1]
                {
          (object) "Subject"
                });
                return target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder)null, target, new object[0]).ToString() != "Volcanoes" ? "False" : str;
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
                return a.Options.SaveInterval != 15 || !d.EmbedTrueTypeFonts ? "False" : str;
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
            foreach (Microsoft.Office.Interop.Word.Table table in d.Tables)
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
