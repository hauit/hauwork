// Decompiled with JetBrains decompiler
// Type: MOS_WORD_TEST.CheckWork22
// Assembly: MOS_WORD_TEST, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 59EF86EC-8F2D-419F-A3E8-5D9A0E352B91
// Assembly location: E:\hau_Work\Git_HauWork\hauwork\WE_19\TEST_WORD_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace MOS_WORD_TEST
{
    public static class CheckWork22
    {
        public static string Check(int cauLon, Application a, Document d)
        {
            switch (cauLon)
            {
                case 0:
                    return CheckWork22.Cau0(a, d);
                case 1:
                    return CheckWork22.Cau1(a, d);
                case 2:
                    return CheckWork22.Cau2(a, d);
                case 3:
                    return CheckWork22.Cau3(a, d);
                case 4:
                    return CheckWork22.Cau4(a, d);
                case 5:
                    return CheckWork22.Cau5(a, d);
                case 6:
                    return CheckWork22.Cau6(a, d);
                case 7:
                    return CheckWork22.Cau7(a, d);
                case 8:
                    return CheckWork22.Cau8(a, d);
                case 9:
                    return CheckWork22.Cau9(a, d);
                case 10:
                    return CheckWork22.Cau10(a, d);
                case 11:
                    return CheckWork22.Cau11(a, d);
                case 12:
                    return CheckWork22.Cau12(a, d);
                case 13:
                    return CheckWork22.Cau13(a, d);
                case 14:
                    return CheckWork22.Cau14(a, d);
                case 15:
                    return CheckWork22.Cau15(a, d);
                case 16:
                    return CheckWork22.Cau16(a, d);
                case 17:
                    return CheckWork22.Cau17(a, d);
                case 18:
                    return CheckWork22.Cau18(a, d);
                case 19:
                    return CheckWork22.Cau19(a, d);
                case 20:
                    return CheckWork22.Cau20(a, d);
                case 21:
                    return CheckWork22.Cau21(a, d);
                case 22:
                    return CheckWork22.Cau22(a, d);
                case 23:
                    return CheckWork22.Cau23(a, d);
                case 24:
                    return CheckWork22.Cau24(a, d);
                case 25:
                    return CheckWork22.Cau25(a, d);
                case 26:
                    return CheckWork22.Cau26(a, d);
                case 27:
                    return CheckWork22.Cau27(a, d);
                case 28:
                    return CheckWork22.Cau28(a, d);
                case 29:
                    return CheckWork22.Cau29(a, d);
                default:
                    return "";
            }
        }

        private static string Cau0(Application a, Document d)
        {
            try
            {
                if (d.Tables.Count != 1)
                    return "False(not insert or delete table)";
                if (d.Tables[1].Descr != "Book Selections")
                    return "False(Book Selections cho Description)";
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
                if (d.Tables.Count != 1)
                    return "False(not insert or delete table)";
                if (d.Tables[1].Borders[WdBorderType.wdBorderBottom].Color.ToString() != "-721354906")
                    return "False(table style)";
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
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("As we promised, you can use these galleries and Friends Book Club®"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False (insert ®)";
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
                return !d.RemovePersonalInformation ? "False(RemovePersonal)" : "True";
            }
            catch (Exception ex)
            {
                return "False(only RemovePersonal)";
            }
        }

        private static string Cau4(Application a, Document d)
        {
            try
            {
                return !d.RemovePersonalInformation ? "False(RemovePersonal)" : "True";
            }
            catch (Exception ex)
            {
                return "False(only RemovePersonal)";
            }
        }

        private static string Cau5(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (Index < 50 && !d.Paragraphs[Index].Range.Text.Contains("Boating Vacations"))
                    ++Index;
                if (Index >= 50)
                    return "False";
                var style1 = d.Paragraphs[Index].get_Style() as Style;
                if (style1 == null || style1.NameLocal != "Title")
                    return "False";
                while (Index < 50 && !d.Paragraphs[Index].Range.Text.Contains("Living and Dining Areas"))
                    ++Index;
                if (Index >= 50)
                    return "False";
                var style2 = d.Paragraphs[Index].get_Style() as Style;
                if (style2 == null || style2.NameLocal != "Heading 1")
                    return "False";
                while (Index < 50 && !d.Paragraphs[Index].Range.Text.Contains("Sleeping Areas"))
                    ++Index;
                if (Index >= 50)
                    return "False";
                var style3 = d.Paragraphs[Index].get_Style() as Style;
                if (style3 == null || style3.NameLocal != "Heading 1")
                    return "False";
                while (Index < 50 && !d.Paragraphs[Index].Range.Text.Contains("Outdoor Facilities"))
                    ++Index;
                if (Index >= 50)
                    return "False";
                var style4 = d.Paragraphs[Index].get_Style() as Style;
                if (style4 == null || style4.NameLocal != "Heading 1")
                    return "False";
                while (Index < 50 && !d.Paragraphs[Index].Range.Text.Contains("Good Weather All Year"))
                    ++Index;
                if (Index >= 50)
                    return "False";
                var style5 = d.Paragraphs[Index].get_Style() as Style;
                if (style5 == null || style5.NameLocal != "Heading 1")
                    return "False";
                while (Index < 150 && !d.Paragraphs[Index].Range.Text.Contains("Simple Planning"))
                    ++Index;
                if (Index >= 150)
                    return "False";
                var style6 = d.Paragraphs[Index].get_Style() as Style;
                if (style6 == null || style6.NameLocal != "Heading 1")
                    return "False";
                while (Index < 150 && !d.Paragraphs[Index].Range.Text.Contains("Contact Us"))
                    ++Index;
                if (Index >= 150)
                    return "False";
                var style7 = d.Paragraphs[Index].get_Style() as Style;
                if (style7 == null || style7.NameLocal != "Heading 1")
                    return "False";
            }
            catch
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string Cau6(Application a, Document d)
        {
            string str = "True";
            object Index = (object)"Heading 1";
            return (double)d.Styles[ref Index].ParagraphFormat.SpaceBefore != 0.0 ? "False" : str;
        }

        private static string Cau7(Application a, Document d)
        {
            string str = "True";
            // ISSUE: reference to a compiler-generated method
            return d.Tables.Count == 3 || d.Tables.Count == 4 || d.Tables.Count == 0 || d.Tables[1].Cell(1, 1).Range.Text.Contains("Luxury") ? "False" : str;
        }

        private static string Cau8(Application a, Document d)
        {
            string str = "True";
            int Index = 1;
            while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Luxury lounge area with couches"))
                ++Index;
            return Index >= d.Paragraphs.Count || d.Paragraphs[Index].Range.ListFormat.ListString != "o" ? "False" : str;
        }

        private static string Cau9(Application a, Document d)
        {
            string str = "True";
            int Index = 1;
            while (Index < 50 && !d.Paragraphs[Index].Range.Text.Contains("Click Insert and then choose the elements"))
                ++Index;
            return Index >= 50 || !d.Paragraphs[Index].Range.Text.Contains("Anywhere you want.\f") ? "False" : str;
        }

        private static string Cau10(Application a, Document d)
        {
            string str = "True";
            if (d.Sections.Count != 3)
                return "False(number of section)";
            if (d.Sections[2].PageSetup.TextColumns.Count != 2)
                return "False(number of column)";
            return d.Sections[2].PageSetup.TextColumns.Spacing.ToString() != "50.4" ? "False(Spacing)" : str;
        }

        private static string Cau11(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Picture") && (double)shape.Height * 100.0 == 16775.0 && (double)shape.Width == 612.0 && shape.WrapFormat.Type == WdWrapType.wdWrapSquare && (double)shape.Left == -999998.0 && (double)shape.Top == -999999.0 && shape.RelativeHorizontalPosition == WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage && shape.RelativeVerticalPosition == WdRelativeVerticalPosition.wdRelativeVerticalPositionPage)
                    return "True";
            }
            return str;
        }

        private static string Cau12(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Picture") && shape.SoftEdge.Type == MsoSoftEdgeType.msoSoftEdgeTypeMixed)
                    return "True";
            }
            return str;
        }

        private static string Cau13(Application a, Document d) => "True";

        private static string Cau14(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Picture") && shape.Rotation.ToString() == "90" && shape.WrapFormat.Type == WdWrapType.wdWrapTight && shape.RelativeHorizontalPosition == WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn && shape.RelativeVerticalPosition == WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph)
                    return "True";
            }
            return str;
        }

        private static string Cau15(Application a, Document d)
        {
            string str = "False";
            foreach (Hyperlink hyperlink in d.Hyperlinks)
            {
                if (hyperlink.Name == "http://wikipedia.org/wiki/Houseboat" && hyperlink.TextToDisplay == "houseboat")
                    return "True";
            }
            return str;
        }

        private static string Cau16(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Table table in d.Tables)
            {
                // ISSUE: reference to a compiler-generated method
                if (table.Cell(1, 1).Range.Text.Contains("Month") && table.Columns.Count == 4 && table.Rows.Count == 13 && table.AllowAutoFit && table.Borders[WdBorderType.wdBorderBottom].ColorIndex == WdColorIndex.wdWhite)
                    return "True";
            }
            return str;
        }

        private static string Cau17(Application a, Document d)
        {
            string str = "True";
            int Index = 90;
            while (Index < 100 && !d.Paragraphs[Index].Range.Text.Contains("Table 1 Average Temperatures"))
                ++Index;
            return Index >= 100 || !d.Paragraphs[Index - 2].Range.Text.Contains("53\r\a") ? "False" : str;
        }

        private static string Cau18(Application a, Document d)
        {
            string str = "False";
            foreach (Bookmark bookmark in d.Bookmarks)
            {
                if (bookmark.Name == "goContact" && bookmark.Range.Text == "Contact Us")
                    return "True";
            }
            return str;
        }

        private static string Cau19(Application a, Document d)
        {
            string str = "True";
            int Index = 50;
            while (Index < 150 && !d.Paragraphs[Index].Range.Text.Contains("Houseboat Lake Vacations, Inc"))
                ++Index;
            return Index >= 150 || (double)d.Paragraphs[Index].SpaceAfter != 2.0 || (double)d.Paragraphs[Index + 1].SpaceAfter != 2.0 || (double)d.Paragraphs[Index + 2].SpaceAfter != 2.0 ? "False" : str;
        }

        private static string Cau20(Application a, Document d)
        {
            string str = "True";
            // ISSUE: variable of a compiler-generated type
            Microsoft.Office.Interop.Word.Shape shape1 = (Microsoft.Office.Interop.Word.Shape)null;
            foreach (Microsoft.Office.Interop.Word.Shape shape2 in d.Shapes)
            {
                if (shape2.Name.Contains("Diagram"))
                    shape1 = shape2;
            }
            if (shape1 == null)
                return "False";
            SmartArt smartArt = shape1.SmartArt;
            return smartArt.QuickStyle.Name != "Intense Effect" || smartArt.Layout.Name != "Basic Process" || smartArt.Nodes.Count != 3 || smartArt.Nodes[(object)1].TextFrame2.TextRange.Text != "Book a houseboat" || smartArt.Nodes[(object)2].TextFrame2.TextRange.Text != "Pack your bag" || smartArt.Nodes[(object)3].TextFrame2.TextRange.Text != "Have fun" ? "False" : str;
        }

        private static string Cau21(Application a, Document d)
        {
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Text Box") && shape.TextFrame.TextRange.Text.Contains("My family spent a week on a houseboat this past summer. It was so much fun") && shape.WrapFormat.Type == WdWrapType.wdWrapTight && shape.TextFrame.TextRange.Font.TextColor.RGB == -738131969 && shape.RelativeVerticalPosition == WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph && shape.RelativeHorizontalPosition == WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage && (double)shape.Width == 252.0 && shape.Height.ToString() == "79.2")
                    return "True";
            }
            return "Fasle(My family spent a week on a houseboat this past summer. It was so much fun)";
        }

        private static string Cau22(Application a, Document d)
        {
            string str = "False";
            foreach (Footnote footnote in d.Footnotes)
            {
                if (footnote.Range.Text.Contains("Fishing license required") && footnote.Reference.FootnoteOptions.NumberStyle == WdNoteNumberStyle.wdNoteNumberStyleSymbol && footnote.Reference.FootnoteOptions.StartingNumber == 2)
                    return "True";
            }
            return str;
        }

        private static string Cau23(Application a, Document d)
        {
            string str = "True";
            return d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count != 0 || d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.TextColor.RGB != -738131969 ? "Fales" : str;
        }

        private static string Cau24(Application a, Document d) => "True";

        private static string Cau25(Application a, Document d)
        {
            string str = "True";
            return d.ActiveWindow.ActivePane.View.Zoom.PageRows != 1 || d.ActiveWindow.ActivePane.View.Zoom.PageColumns != 2 ? "False" : str;
        }

        private static string Cau26(Application a, Document d)
        {
            string str = "True";
            object documentProperties = d.BuiltInDocumentProperties;
            Type type = documentProperties.GetType();
            object target1 = type.InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder)null, documentProperties, new object[1]
            {
        (object) "Title"
            });
            if (target1.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder)null, target1, new object[0]).ToString() != "Boating Vacations")
                return "False";
            object target2 = type.InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder)null, documentProperties, new object[1]
            {
        (object) "Comments"
            });
            return target2.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder)null, target2, new object[0]).ToString() != "Please review yearly" ? "False" : str;
        }

        private static string Cau27(Application a, Document d)
        {
            string str = "True";
            return a.Options.SaveInterval != 8 || !d.EmbedTrueTypeFonts || !a.Options.PrintBackgrounds ? "False" : str;
        }

        private static string Cau28(Application a, Document d) => "";

        private static string Cau29(Application a, Document d) => "";
    }
}
