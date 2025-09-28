// Decompiled with JetBrains decompiler
// Type: MOS_WORD_TEST.CheckWork8
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
    public static class CheckWork8
    {
        public static string Check(int cauLon, Application a, Document d)
        {
            switch (cauLon)
            {
                case 0:
                    return CheckWork8.Cau0(a, d);
                case 1:
                    return CheckWork8.Cau1(a, d);
                case 2:
                    return CheckWork8.Cau2(a, d);
                case 3:
                    return CheckWork8.Cau3(a, d);
                case 4:
                    return CheckWork8.Cau4(a, d);
                case 5:
                    return CheckWork8.Cau5(a, d);
                case 6:
                    return CheckWork8.Cau6(a, d);
                case 7:
                    return CheckWork8.Cau7(a, d);
                case 8:
                    return CheckWork8.Cau8(a, d);
                case 9:
                    return CheckWork8.Cau9(a, d);
                case 10:
                    return CheckWork8.Cau10(a, d);
                case 11:
                    return CheckWork8.Cau11(a, d);
                case 12:
                    return CheckWork8.Cau12(a, d);
                case 13:
                    return CheckWork8.Cau13(a, d);
                case 14:
                    return CheckWork8.Cau14(a, d);
                case 15:
                    return CheckWork8.Cau15(a, d);
                case 16:
                    return CheckWork8.Cau16(a, d);
                case 17:
                    return CheckWork8.Cau17(a, d);
                case 18:
                    return CheckWork8.Cau18(a, d);
                case 19:
                    return CheckWork8.Cau19(a, d);
                case 20:
                    return CheckWork8.Cau20(a, d);
                case 21:
                    return CheckWork8.Cau21(a, d);
                case 22:
                    return CheckWork8.Cau22(a, d);
                case 23:
                    return CheckWork8.Cau23(a, d);
                case 24:
                    return CheckWork8.Cau24(a, d);
                case 25:
                    return CheckWork8.Cau25(a, d);
                case 26:
                    return CheckWork8.Cau26(a, d);
                case 27:
                    return CheckWork8.Cau27(a, d);
                case 28:
                    return CheckWork8.Cau28(a, d);
                case 29:
                    return CheckWork8.Cau29(a, d);
                case 30:
                    return CheckWork8.Cau30(a, d);
                case 31:
                    return CheckWork8.Cau31(a, d);
                case 32:
                    return CheckWork8.Cau32(a, d);
                case 33:
                    return CheckWork8.Cau33(a, d);
                case 34:
                    return CheckWork8.Cau34(a, d);
                case 35:
                    return CheckWork8.Cau35(a, d);
                case 36:
                    return CheckWork8.Cau36(a, d);
                case 37:
                    return CheckWork8.Cau37(a, d);
                case 38:
                    return CheckWork8.Cau38(a, d);
                case 39:
                    return CheckWork8.Cau39(a, d);
                default:
                    return "";
            }
        }

        private static string Cau0(Application a, Document d)
        {
            try
            {
                if (d.CompatibilityMode == 11)
                    return "False(File==>Convert)";
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
                if ((double)d.Paragraphs[2].Range.InlineShapes[1].Line.Weight != 15.0)
                    return "False(Matal Frame)";
            }
            catch (Exception ex)
            {
                return "False (Convert Text to WordArt)";
            }
            return "True";
        }

        private static string Cau2(Application a, Document d)
        {
            try
            {
            }
            catch (Exception ex)
            {
                return "False (loi khong xac dinh)";
            }
            return "True";
        }

        private static string Cau3(Application a, Document d)
        {
            try
            {
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("class"))
                        return "False(class)";
                }
                int Index1 = 1;
                while (Index1 < d.Paragraphs.Count && !d.Paragraphs[Index1].Range.Text.Contains("seminar"))
                    ++Index1;
                if (Index1 >= d.Paragraphs.Count)
                    return "False (seminar)";
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
                // ISSUE: variable of a compiler-generated type
                Paragraph paragraph1 = (Paragraph)null;
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

        private static string Cau5(Application a, Document d)
        {
            string str = "True";
            int Index = 1;
            while (!d.Paragraphs[Index].Range.Text.Contains("When you click Online Video, you can paste in the embed code") && Index < 100)
                ++Index;
            if (Index >= 100)
                return "False";
            // ISSUE: variable of a compiler-generated type
            ListFormat listFormat1 = d.Paragraphs[Index + 1].Range.ListFormat;
            // ISSUE: variable of a compiler-generated type
            ListFormat listFormat2 = d.Paragraphs[Index + 6].Range.ListFormat;
            return listFormat2.ListValue != 5 || (double)d.Paragraphs[Index + 4].LeftIndent != 54.0 ? "False" : str;
        }

        private static string Cau6(Application a, Document d)
        {
            string str = "True";
            try
            {
                object Index = (object)"Picture 1";
                // ISSUE: variable of a compiler-generated type
                Microsoft.Office.Interop.Word.Shape shape = d.Shapes[ref Index];
                if ((double)shape.Left != -999996.0 || (double)shape.Top != -999999.0 || shape.RelativeHorizontalPosition != WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin)
                    return "False";
                if (shape.RelativeVerticalPosition != WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return str;
        }

        private static string Cau7(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Smiley Face"))
                    return "True";
            }
            return str;
        }

        private static string Cau8(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Text Box") && shape.TextFrame.TextRange.Text == "When you aren't toxic, you are happy!\r")
                    return "True";
            }
            return str;
        }

        private static string Cau9(Application a, Document d)
        {
            string str = "True";
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            return d.Tables.Count != 1 || !d.Tables[1].Cell(1, 1).Range.Text.Contains("Toxins") || !d.Tables[1].Cell(11, 4).Range.Text.Contains("Some nausea") ? "False" : str;
        }

        private static string Cau10(Application a, Document d)
        {
            string str = "True";
            // ISSUE: reference to a compiler-generated method
            return d.Tables.Count != 1 || !d.Tables[1].Cell(6, 4).Range.Text.Contains("Constipation") ? "False" : str;
        }

        private static string Cau11(Application a, Document d)
        {
            string str = "True";
            return d.Tables.Count != 1 ? "False" : str;
        }

        private static string Cau12(Application a, Document d)
        {
            string str = "True";
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            return d.Tables[1].Cell(1, 1).VerticalAlignment != WdCellVerticalAlignment.wdCellAlignVerticalCenter || d.Tables[1].Cell(1, 4).VerticalAlignment != WdCellVerticalAlignment.wdCellAlignVerticalCenter || d.Tables[1].Cell(1, 1).Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphCenter || d.Tables[1].Cell(1, 4).Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphCenter ? "False" : str;
        }

        private static string Cau13(Application a, Document d)
        {
            string str = "True";
            object Index1 = (object)"Intense Emphasis";
            if (d.Styles[ref Index1].Font.TextColor.RGB != -587137063)
                return "False";
            object Index2 = (object)"Intense Emphasis";
            return (double)d.Styles[ref Index2].Font.Size != 12.0 ? "False" : str;
        }

        private static string Cau14(Application a, Document d)
        {
            string str = "True";
            int Index = 1;
            while (!d.Paragraphs[Index].Range.Text.Contains("you okay, Sparky") && Index < 100)
                ++Index;
            if (Index >= 100)
                return "False";
            var style1 = d.Paragraphs[Index].Range.ParagraphStyle as Style;
            if (style1 == null || style1.NameLocal != "Title")
                return "False";
            while (!d.Paragraphs[Index].Range.Text.Contains("What about our pets?") && Index < 100)
                ++Index;
            if (Index >= 100)
                return "False";
            var style2 = d.Paragraphs[Index].Range.CharacterStyle as Style;
            return (style2 == null || style2.NameLocal != "Intense Emphasis") ? "False" : str;
        }

        private static string Cau15(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Picture") && (int)shape.Width == 224 && (int)shape.Height == 202 && shape.ShapeStyle == MsoShapeStyleIndex.msoShapeStyleNotAPreset && shape.WrapFormat.Type == WdWrapType.wdWrapTight)
                    return "True";
            }
            return str;
        }

        private static string Cau16(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Text") && shape.TextFrame.TextRange.Text == "A\r")
                    return "True";
            }
            return str;
        }

        private static string Cau17(Application a, Document d)
        {
            string str = "True";
            int Index = 1;
            while (!d.Paragraphs[Index].Range.Text.Contains("lukewarm water") && Index < 100)
                ++Index;
            if (Index >= 100)
                return "False";
            int num = d.Paragraphs[Index].Range.Text.IndexOf("lukewarm water");
            // ISSUE: variable of a compiler-generated type
            Range character = d.Paragraphs[Index].Range.Characters[num + 3];
            return character.Underline != WdUnderline.wdUnderlineDouble ? "False" : str;
        }

        private static string Cau18(Application a, Document d)
        {
            string str = "True";
            int Index = 1;
            while (!d.Paragraphs[Index].Range.Text.Contains("Question & Answer with Dr. Nelson") && Index < 100)
                ++Index;
            if (Index >= 100)
                return "False";
            var style = d.Paragraphs[Index].Range.ParagraphStyle as Style;
            return (style == null || style.NameLocal != "Heading 1") ? "False" : str;
        }

        private static string Cau19(Application a, Document d)
        {
            string str = "True";
            return (double)d.Sections[d.Sections.Count].PageSetup.LeftMargin != 144.0 || (double)d.Sections[d.Sections.Count].PageSetup.RightMargin != 144.0 || (double)d.Sections[d.Sections.Count].PageSetup.TopMargin != 72.0 || (double)d.Sections[d.Sections.Count].PageSetup.BottomMargin != 72.0 ? "False" : str;
        }

        private static string Cau20(Application a, Document d)
        {
            string str = "True";
            return d.Footnotes.Count < 1 ? "False" : str;
        }

        private static string Cau21(Application a, Document d)
        {
            string str = "True";
            // ISSUE: variable of a compiler-generated type
            Hyperlink hyperlink1 = (Hyperlink)null;
            foreach (Hyperlink hyperlink2 in d.Hyperlinks)
            {
                if (hyperlink2.Name == "mailto:williamnelson@hypotheticaluniversity.com")
                    hyperlink1 = hyperlink2;
            }
            return hyperlink1 == null || !hyperlink1.TextToDisplay.Contains("Emaill all of your questions") ? "False" : str;
        }

        private static string Cau22(Application a, Document d) => "True";

        private static string Cau23(Application a, Document d)
        {
            string str = "False";
            foreach (Paragraph paragraph in d.Paragraphs)
            {
                if (paragraph.Range.Text.Contains("Toxins can reach a critical"))
                    return paragraph.Range.Characters[1].Bold == 0 || paragraph.Range.Characters[1].Underline == WdUnderline.wdUnderlineNone ? "False" : "True";
            }
            return str;
        }

        private static string Cau24(Application a, Document d)
        {
            string str = "True";
            foreach (Field field in d.Fields)
            {
                if (!(field.Code.Text == ""))
                    ;
            }
            return str;
        }

        private static string Cau25(Application a, Document d)
        {
            string str = "True";
            if (d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.Count != 0)
                return "False (PageNumbers)";
            if (d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count != 1)
                return "False(number of shape)";
            object Index = (object)1;
            return !d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[ref Index].Name.Contains("Curved Down Ribbon") ? "False(shape name)" : str;
        }

        private static string Cau26(Application a, Document d)
        {
            string str = "True";
            object documentProperties = d.BuiltInDocumentProperties;
            object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder)null, documentProperties, new object[1]
            {
        (object) "Title"
            });
            return target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder)null, target, new object[0]).ToString() != "Health Newsletter" ? "False" : str;
        }

        private static string Cau27(Application a, Document d)
        {
            string str = "True";
            return d.ProtectionType != WdProtectionType.wdAllowOnlyRevisions ? "False" : str;
        }

        private static string Cau28(Application a, Document d)
        {
            string str = "True";
            return !a.Options.PrintProperties || !a.Options.AllowOpenInDraftView || a.Options.CheckSpellingAsYouType ? "False" : str;
        }

        private static string Cau29(Application a, Document d)
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
                    if ((int)table.Cell(1, 1).Width != 506 || table.Borders.Count != 8 || table.Borders[WdBorderType.wdBorderBottom].ColorIndex != WdColorIndex.wdGray25)
                        return "False";
                    str = "True";
                }
            }
            return str;
        }

        private static string Cau30(Application a, Document d)
        {
            string str = "False";
            int Index1 = 1;
            while (d.Paragraphs[Index1].Range.Text != "Scoring Examples\r" && Index1 < 100)
                ++Index1;
            if (Index1 >= 100 || !d.Paragraphs[Index1 + 1].Range.Text.Contains("Michigan 2016 Region 4 Aesop Gone Viral Div"))
                return "False";
            for (int Index2 = Index1 + 80; Index2 < Index1 + 88; ++Index2)
            {
                if (d.Paragraphs[Index2].Range.Text.Contains("Table 1-Michigan Odyssey Web Site, Oct. 22, 2016"))
                    return "True";
            }
            return str;
        }

        private static string Cau31(Application a, Document d)
        {
            string str = "True";
            if (d.Bookmarks.Count != 1)
                return "False";
            object Index = (object)"Scoring";
            return d.Bookmarks[ref Index].Range.Text != "Scoring Examples" ? "False" : str;
        }

        private static string Cau32(Application a, Document d)
        {
            string str = "True";
            return d.InlineShapes.Count != 1 || d.InlineShapes[1].SmartArt.Nodes.Count != 5 || d.InlineShapes[1].SmartArt.Nodes[(object)1].TextFrame2.TextRange.Text != "Problem 1: No- Cycle Recycle" ? "False" : str;
        }

        private static string Cau33(Application a, Document d)
        {
            string str = "False";
            int count = d.Paragraphs.Count;
            for (int Index = count - 5; Index < count; ++Index)
            {
                if (d.Paragraphs[Index].Range.Text.Contains("Figure 2-OdysseyOfTheMind.com, Oct. 22, 2016"))
                    return "True";
            }
            return str;
        }

        private static string Cau34(Application a, Document d)
        {
            string str = "True";
            return d.InlineShapes.Count != 1 || d.InlineShapes[1].SmartArt.Nodes.Count != 5 || d.InlineShapes[1].SmartArt.Nodes[(object)1].TextFrame2.TextRange.Text != "Problem 1: No- Cycle Recycle" || d.InlineShapes[1].SmartArt.QuickStyle.Name != "Simple Fill" || d.InlineShapes[1].SmartArt.QuickStyle.Id != "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1" || (int)d.InlineShapes[1].Width != 525 || (int)d.InlineShapes[1].Height != 338 ? "False" : str;
        }

        private static string Cau35(Application a, Document d)
        {
            string str = "True";
            object documentProperties = d.BuiltInDocumentProperties;
            object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder)null, documentProperties, new object[1]
            {
        (object) "Keywords"
            });
            return target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder)null, target, new object[0]).ToString() != "MOS Word OM Test" ? "False" : str;
        }

        private static string Cau36(Application a, Document d)
        {
            string str = "True";
            return a.Options.SaveInterval != 4 || !d.EmbedTrueTypeFonts ? "False" : str;
        }

        private static string Cau37(Application a, Document d)
        {
            string str = "True";
            return d.ShowGrammaticalErrors ? "False" : str;
        }

        private static string Cau38(Application a, Document d)
        {
            string str = "True";
            return !a.ActiveWindow.DocumentMap ? "False" : str;
        }

        private static string Cau39(Application a, Document d) => "True";
    }
}
