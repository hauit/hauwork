// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.CheckWork2
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
{
    public static class CheckWork2
    {
        public static string Check(int cauLon, Application a, Document d)
        {
            switch (cauLon)
            {
                case 0:
                    return CheckWork2.Cau0(a, d);
                case 1:
                    return CheckWork2.Cau1(a, d);
                case 2:
                    return CheckWork2.Cau2(a, d);
                case 3:
                    return CheckWork2.Cau3(a, d);
                case 4:
                    return CheckWork2.Cau4(a, d);
                case 5:
                    return CheckWork2.Cau5(a, d);
                case 6:
                    return CheckWork2.Cau6(a, d);
                case 7:
                    return CheckWork2.Cau7(a, d);
                case 8:
                    return CheckWork2.Cau8(a, d);
                case 9:
                    return CheckWork2.Cau9(a, d);
                case 10:
                    return CheckWork2.Cau10(a, d);
                case 11:
                    return CheckWork2.Cau11(a, d);
                case 12:
                    return CheckWork2.Cau12(a, d);
                case 13:
                    return CheckWork2.Cau13(a, d);
                case 14:
                    return CheckWork2.Cau14(a, d);
                case 15:
                    return CheckWork2.Cau15(a, d);
                case 16:
                    return CheckWork2.Cau16(a, d);
                case 17:
                    return CheckWork2.Cau17(a, d);
                case 18:
                    return CheckWork2.Cau18(a, d);
                case 19:
                    return CheckWork2.Cau19(a, d);
                case 20:
                    return CheckWork2.Cau20(a, d);
                case 21:
                    return CheckWork2.Cau21(a, d);
                case 22:
                    return CheckWork2.Cau22(a, d);
                case 23:
                    return CheckWork2.Cau23(a, d);
                case 24:
                    return CheckWork2.Cau24(a, d);
                case 25:
                    return CheckWork2.Cau25(a, d);
                case 26:
                    return CheckWork2.Cau26(a, d);
                case 27:
                    return CheckWork2.Cau27(a, d);
                case 28:
                    return CheckWork2.Cau28(a, d);
                case 29:
                    return CheckWork2.Cau29(a, d);
                default:
                    return "";
            }
        }

        private static string Cau0(Application a, Document d)
        {
            try
            {
                object Index = (object)"Text Box 2";
                if (!d.Shapes[ref Index].TextFrame.TextRange.Text.Contains("bicycles"))
                    return "False";
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
                object Index = (object)"Picture 4";
                if (d.Shapes[ref Index].WrapFormat.Type != WdWrapType.wdWrapSquare)
                    return "False(thay đổi Wraptext cho ảnh thành Square)";
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
                if (d.Tables.Count != 1)
                    return "False(không thêm xóa table)";
                if (d.Tables[1].Rows[1].Cells.Count != 1)
                    return "False(trọng ô dòng đầu)";
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
                int num1 = 0;
                int num2 = 0;
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.ToLower().Contains("city"))
                        ++num1;
                    if (d.Paragraphs[Index].Range.Text.ToLower().Contains("community"))
                        ++num2;
                }
                if (num1 > 0)
                    return "False";
                if (num2 != 6)
                    return "False";
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
                if (d.Comments.Count != 1)
                    return "False(không thêm xóa comment)";
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

        private static string Cau6(Application a, Document d)
        {
            try
            {
                if (!a.ActiveWindow.View.ShowTabs)
                    return "False(show Tab)";
                if (a.ActiveWindow.View.ShowSpaces)
                    return "False(don't show Spaces)";
                if (a.ActiveWindow.View.ShowHiddenText)
                    return "False(don't show HiddenText)";
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
            if (d.Sections.Count != 2)
                return "False(number section<>2)";
            return d.Sections[1].PageSetup.TextColumns.Count != 2 ? "False(<>2 Columns)" : str;
        }

        private static string Cau8(Application a, Document d)
        {
            string str = "True";
            int Index = 1;
            while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("1,Mercury") && !d.Paragraphs[Index].Range.Text.Contains("1, Mercury"))
                ++Index;
            if (Index >= d.Paragraphs.Count)
                return "False (1,Mercury)";
            if (!d.Paragraphs[Index + 1].Range.Text.Contains("2,Venus") && !d.Paragraphs[Index + 1].Range.Text.Contains("2, Venus"))
                return "False(2,Venus)";
            return !d.Paragraphs[Index + 2].Range.Text.Contains("3,Earth") && !d.Paragraphs[Index + 2].Range.Text.Contains("3, Earth") ? "False(3,Earth)" : str;
        }

        private static string Cau9(Application a, Document d)
        {
            string str = "True";
            if (d.InlineShapes.Count != 1)
                return "False(smartArt layout must 'inline with text)";
            if (d.InlineShapes[1].SmartArt.Nodes.Count != 2)
                return "False(<>2 section)";
            if (d.InlineShapes[1].SmartArt.Nodes[(object)1].TextFrame2.TextRange.Text != "In Suspension")
                return "False(In Suspension)";
            if ((int)d.InlineShapes[1].Width != 432)
                return "False(Width)";
            if ((int)d.InlineShapes[1].Height != 115)
                return "False(Height)";
            if (d.InlineShapes[1].SmartArt.Nodes[(object)2].TextFrame2.TextRange.Text != "White Island")
                return "False(White Island)";
            if (d.InlineShapes[1].SmartArt.Nodes[(object)2].Nodes.Count != 2)
                return "False(1)";
            if (d.InlineShapes[1].SmartArt.Nodes[(object)1].Nodes.Count != 2)
                return "False(2)";
            if (d.InlineShapes[1].SmartArt.Nodes[(object)1].Nodes[(object)1].TextFrame2.TextRange.Text != "Son")
                return "False(Son)";
            if (d.InlineShapes[1].SmartArt.Nodes[(object)1].Nodes[(object)2].TextFrame2.TextRange.Text != "Slag")
                return "False(Slag)";
            if (d.InlineShapes[1].SmartArt.Nodes[(object)2].Nodes[(object)1].TextFrame2.TextRange.Text != "Italy")
                return "False(Italy)";
            return d.InlineShapes[1].SmartArt.Nodes[(object)2].Nodes[(object)2].TextFrame2.TextRange.Text != "New Zealand" ? "False(New Zealand)" : str;
        }

        private static string Cau10(Application a, Document d)
        {
            int Index1 = 5;
            while (Index1 < d.Paragraphs.Count && d.Paragraphs[Index1].Range.Text != "Mercury\r")
                ++Index1;
            if (Index1 >= d.Paragraphs.Count)
                return "False(heading was modified)";
            for (int Index2 = Index1 + 1; Index2 < Index1 + 4; ++Index2)
            {
                if (d.Paragraphs[Index2].Range.Text.Contains("Figure 1-Some contain iron, magnesium, silica, or aluminum"))
                    return "True";
            }
            return "False(Figure 1-Some contain iron, magnesium, silica, or aluminum)";
        }

        private static string Cau11(Application a, Document d)
        {
            string str = "True";
            int Index = 5;
            while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Earth is the third planet from the earth"))
                ++Index;
            if (Index >= d.Paragraphs.Count)
                return "False(import text)";
            return !d.Paragraphs[Index - 1].Range.Text.Contains("Earth") ? "False(wrong position)" : str;
        }

        private static string Cau12(Application a, Document d)
        {
            foreach (Table table in d.Tables)
            {
                if (table.Columns.Count == 5)
                {
                    if (table.PreferredWidthType != WdPreferredWidthType.wdPreferredWidthPercent)
                        return "False (autofix)";
                    if (!table.Cell(1, 1).Range.Text.Contains("Name"))
                        return "False(cell[1,1])";
                    if (table.AutoFormatType != 1)
                        return "False(FormatType)";
                    if (table.Rows.Count != 9)
                        return "False(row)";
                    if ((int)table.Cell(1, 1).Width != 93)
                        return "False(fix window)";
                    var style = table.get_Style() as Style;
                    if (style == null || style.NameLocal != "List Table 4 - Accent 1")
                        return "False(" + (style != null ? style.NameLocal : "null") + ")";
                    if (!table.Cell(2, 5).Range.Text.Contains("2") || !table.Cell(9, 5).Range.Text.Contains("30"))
                        return "False(sort)";
                }
            }
            return "True";
        }

        private static string Cau13(Application a, Document d)
        {
            for (int Index = d.Paragraphs.Count - 20; Index < d.Paragraphs.Count; ++Index)
            {
                if (d.Paragraphs[Index].Range.Text.Contains("Table 1-Some information of eight planets"))
                    return "True";
            }
            return "False(Table 1-Some information of eight planets)";
        }

        private static string Cau14(Application a, Document d)
        {
            string str = "True";
            object Index = (object)"Picture 1";
            // ISSUE: variable of a compiler-generated type
            Shape shape = d.Shapes[ref Index];
            if (shape.WrapFormat.Type != WdWrapType.wdWrapTight)
                return "False(WrapTight)";
            if (shape.RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin)
                return "False(HorizontalPositionMargin)";
            if (shape.RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin)
                return "False(VerticalPositionMargin)";
            if ((double)shape.Left != -999996.0)
                return "False(RelativeHorizontal)";
            return (double)shape.Top != -999997.0 ? "False(RelativeVertical)" : str;
        }

        private static string Cau15(Application a, Document d)
        {
            foreach (Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Text Box") && shape.TextFrame.TextRange.Font.TextColor.RGB == -738131969 && shape.TextFrame.TextRange.Text.Contains("Andesite is one of the most common volcanic rocks and can contain olivine"))
                {
                    if ((double)shape.Left != -999998.0)
                        return "False(H position)";
                    if ((double)shape.Top != -999997.0)
                        return "False(V position)";
                    if (shape.RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin)
                        return "False(H Margin)";
                    return shape.RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin ? "False(V Margin)" : "True";
                }
            }
            return "False(Andesite is one of the most common volcanic rocks and can contain olivine)";
        }

        private static string Cau16(Application a, Document d)
        {
            string str = "True";
            if (d.Footnotes.Count != 1)
                return "False (number footnote)";
            return d.Footnotes[1].Range.Text != "we can search it online from 2008." ? "False(we can search it online from 2008.)" : str;
        }

        private static string Cau17(Application a, Document d)
        {
            string str = "True";
            if (d.Hyperlinks.Count != 1)
                return "False(number hyperlink)";
            object Index = (object)1;
            // ISSUE: variable of a compiler-generated type
            Hyperlink hyperlink = d.Hyperlinks[ref Index];
            if (hyperlink.TextToDisplay != "about planets")
                return "False(about planets)";
            return hyperlink.Name != "_About_Planets" ? "False(About Planets Heading)" : str;
        }

        private static string Cau18(Application a, Document d)
        {
            string str = "True";
            return a.ActiveWindow.View.Zoom.Percentage != 110 ? "False(zoom 110%)" : str;
        }

        private static string Cau19(Application a, Document d)
        {
            string str = "True";
            return d.ShowSpellingErrors ? "False(Hide Spelling)" : str;
        }

        private static string Cau20(Application a, Document d)
        {
            string str = "True";
            return a.Options.SaveInterval != 12 ? "False(12 minute)" : str;
        }

        private static string Cau21(Application a, Document d) => "";

        private static string Cau22(Application a, Document d) => "";

        private static string Cau23(Application a, Document d) => "";

        private static string Cau24(Application a, Document d) => "";

        private static string Cau25(Application a, Document d) => "";

        private static string Cau26(Application a, Document d) => "";

        private static string Cau27(Application a, Document d) => "";

        private static string Cau28(Application a, Document d) => "";

        private static string Cau29(Application a, Document d) => "";
    }
}
