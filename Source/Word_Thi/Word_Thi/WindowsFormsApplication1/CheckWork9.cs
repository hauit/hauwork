// Decompiled with JetBrains decompiler
// Type: MOS_WORD_TEST.CheckWork9
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
    public static class CheckWork9
    {
        public static string Check(int cauLon, Application a, Document d)
        {
            switch (cauLon)
            {
                case 0:
                    return CheckWork9.Cau0(a, d);
                case 1:
                    return CheckWork9.Cau1(a, d);
                case 2:
                    return CheckWork9.Cau2(a, d);
                case 3:
                    return CheckWork9.Cau3(a, d);
                case 4:
                    return CheckWork9.Cau4(a, d);
                case 5:
                    return CheckWork9.Cau5(a, d);
                case 6:
                    return CheckWork9.Cau6(a, d);
                case 7:
                    return CheckWork9.Cau7(a, d);
                case 8:
                    return CheckWork9.Cau8(a, d);
                case 9:
                    return CheckWork9.Cau9(a, d);
                case 10:
                    return CheckWork9.Cau10(a, d);
                case 11:
                    return CheckWork9.Cau11(a, d);
                case 12:
                    return CheckWork9.Cau12(a, d);
                case 13:
                    return CheckWork9.Cau13(a, d);
                case 14:
                    return CheckWork9.Cau14(a, d);
                case 15:
                    return CheckWork9.Cau15(a, d);
                case 16:
                    return CheckWork9.Cau16(a, d);
                case 17:
                    return CheckWork9.Cau17(a, d);
                case 18:
                    return CheckWork9.Cau18(a, d);
                case 19:
                    return CheckWork9.Cau19(a, d);
                case 20:
                    return CheckWork9.Cau20(a, d);
                case 21:
                    return CheckWork9.Cau21(a, d);
                case 22:
                    return CheckWork9.Cau22(a, d);
                case 23:
                    return CheckWork9.Cau23(a, d);
                case 24:
                    return CheckWork9.Cau24(a, d);
                case 25:
                    return CheckWork9.Cau25(a, d);
                case 26:
                    return CheckWork9.Cau26(a, d);
                case 27:
                    return CheckWork9.Cau27(a, d);
                case 28:
                    return CheckWork9.Cau28(a, d);
                case 29:
                    return CheckWork9.Cau29(a, d);
                default:
                    return "";
            }
        }

        private static string Cau0(Application a, Document d)
        {
            try
            {
                if (d.Name != "Notes.dotx")
                    return "False(luu lại kiểu template(Notes.dotx))";
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
                // ISSUE: variable of a compiler-generated type
                Microsoft.Office.Interop.Word.Shape shape = (Microsoft.Office.Interop.Word.Shape)null;
                for (int index = 1; index <= d.Shapes.Count; ++index)
                {
                    object Index1 = (object)index;
                    if (d.Shapes[ref Index1].Name.Contains("Text Box"))
                    {
                        object Index2 = (object)index;
                        shape = d.Shapes[ref Index2];
                    }
                }
                if (!shape.TextFrame.TextRange.Text.ToUpper().Contains("NEW EMPLOYEE TRAINING"))
                    return "False(NEW EMPLOYEE TRAINING)";
                if (shape.TextFrame.TextRange.Font.Fill.ForeColor.RGB.ToString() != "6051667")
                    return "False(wordArt type)";
            }
            catch (Exception ex)
            {
                return "False (convert text to wordArt)";
            }
            return "True";
        }

        private static string Cau2(Application a, Document d)
        {
            try
            {
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("Friday"))
                        return "False(Friday)";
                }
                int Index1 = 1;
                while (Index1 < d.Paragraphs.Count && !d.Paragraphs[Index1].Range.Text.Contains("Tuesday"))
                    ++Index1;
                if (Index1 >= d.Paragraphs.Count)
                    return "False (Tuesday)";
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
                // ISSUE: variable of a compiler-generated type
                Microsoft.Office.Interop.Word.Shape shape = (Microsoft.Office.Interop.Word.Shape)null;
                for (int index = 1; index <= d.Shapes.Count; ++index)
                {
                    object Index1 = (object)index;
                    if (d.Shapes[ref Index1].Name.Contains("Horizontal Scroll"))
                    {
                        object Index2 = (object)index;
                        shape = d.Shapes[ref Index2];
                    }
                }
                if (shape == null)
                    return "False(insert shape)";
                if (!shape.TextFrame.TextRange.Text.ToLower().Contains("class start at 9:00!"))
                    return "False(Class start at 9:00!)";
                float num = shape.Left;
                if (num.ToString() != "-999995")
                    return "False(H)";
                num = shape.Top;
                if (num.ToString() != "-999997")
                    return "False(V)";
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
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Broad-leaved Arrow-head"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(Broad-leaved Arrow-head)";
                if (!d.Paragraphs[Index].Range.Text.Contains("\u000EBroad-leaved Arrow-head\r"))
                    return "False(để con tro đâu đoạn Broad-leaved Arrow-head)";
            }
            catch (Exception ex)
            {
                return "False(Something wrong)";
            }
            return "True";
        }

        private static string Cau5(Application a, Document d)
        {
            string str = "True";
            return d.Sections.Count != 2 || d.Sections[1].PageSetup.TextColumns.Count != 1 || d.Sections[2].PageSetup.TextColumns.Count != 2 ? "False" : str;
        }

        private static string Cau6(Application a, Document d)
        {
            string str = "True";
            var style = d.Paragraphs[1].Range.ParagraphStyle as Style;
            return (style == null || style.NameLocal != "Title") ? "False" : str;
        }

        private static string Cau7(Application a, Document d)
        {
            string str = "True";
            object Index = (object)"Title";
            // ISSUE: variable of a compiler-generated type
            Style style = d.Styles[ref Index];
            return style.Font.TextColor.RGB != -704593921 ? "False" : str;
        }

        private static string Cau8(Application a, Document d)
        {
            string str = "True";
            int Index = 1;
            while (Index < 100 && !d.Paragraphs[Index].Range.Text.Contains("Broad-leaved Arrow-head"))
                ++Index;
            if (Index >= 100)
                return "False";
            var style1 = d.Paragraphs[Index].Range.ParagraphStyle as Style;
            if (style1 == null || style1.NameLocal != "Heading 1")
                return "False";
            while (Index < 100 && !d.Paragraphs[Index].Range.Text.Contains("Flowers--White"))
                ++Index;
            if (Index >= 100)
                return "False";
            var style2 = d.Paragraphs[Index].Range.Characters[2].CharacterStyle as Style;
            if (style2 == null || style2.NameLocal != "Intense Emphasis")
                return "False";
            while (Index < 100 && !d.Paragraphs[Index].Range.Text.Contains("Flower Season--July"))
                ++Index;
            if (Index >= 100)
                return "False";
            var style3 = d.Paragraphs[Index].Range.Characters[2].CharacterStyle as Style;
            return (style3 == null || style3.NameLocal != "Intense Emphasis") ? "False" : str;
        }

        private static string Cau9(Application a, Document d)
        {
            string str = "True";
            int Index = 1;
            while (Index < 10 && !d.Paragraphs[Index].Range.Text.Contains("Location"))
                ++Index;
            return Index >= 10 || d.Paragraphs[Index].Range.ListFormat.ListLevelNumber != 1 || d.Paragraphs[Index + 1].Range.ListFormat.ListLevelNumber != 2 || d.Paragraphs[Index + 2].Range.ListFormat.ListLevelNumber != 2 || d.Paragraphs[Index + 3].Range.ListFormat.ListLevelNumber != 1 || d.Paragraphs[Index + 4].Range.ListFormat.ListLevelNumber != 2 || d.Paragraphs[Index + 5].Range.ListFormat.ListLevelNumber != 2 || d.Paragraphs[Index + 6].Range.ListFormat.ListLevelNumber != 1 || d.Paragraphs[Index + 7].Range.ListFormat.ListLevelNumber != 2 || d.Paragraphs[Index + 8].Range.ListFormat.ListLevelNumber != 2 || d.Paragraphs[Index + 9].Range.ListFormat.ListLevelNumber != 2 || d.Paragraphs[Index + 10].Range.ListFormat.ListLevelNumber != 1 || d.Paragraphs[Index + 11].Range.ListFormat.ListLevelNumber != 2 ? "False" : str;
        }

        private static string Cau10(Application a, Document d)
        {
            string str = "False";
            foreach (Hyperlink hyperlink in d.Hyperlinks)
            {
                if (hyperlink.Name == "http://www.fusiontomo.com/bloomcolors" && hyperlink.TextToDisplay == "Color of the blooms")
                    return "True";
            }
            return str;
        }

        private static string Cau11(Application a, Document d) => "True";

        private static string Cau12(Application a, Document d)
        {
            string str = "True";
            int Index = 1;
            while (Index < 10 && !d.Paragraphs[Index].Range.Text.Contains("Location"))
                ++Index;
            return Index >= 10 || (double)d.Paragraphs[Index].Range.Font.Size != 12.0 || d.Paragraphs[Index].Range.Font.Name != "Cooper Black" || (double)d.Paragraphs[Index + 10].Range.Font.Size != 12.0 || d.Paragraphs[Index + 10].Range.Font.Name != "Cooper Black" || (double)d.Paragraphs[Index + 6].Range.Font.Size != 12.0 || d.Paragraphs[Index + 6].Range.Font.Name != "Cooper Black" ? "False" : str;
        }

        private static string Cau13(Application a, Document d)
        {
            string str = "True";
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            return d.Tables.Count != 1 || d.Tables[1].Rows.Count != 5 || d.Tables[1].Columns.Count != 3 || !d.Tables[1].Cell(1, 1).Range.Text.Contains("Flower Name") || !d.Tables[1].Cell(5, 3).Range.Text.Contains("Often picked") ? "False" : str;
        }

        private static string Cau14(Application a, Document d)
        {
            string str = "True";
            // ISSUE: reference to a compiler-generated method
            return d.Tables[1].AutoFormatType != 1 || d.Tables[1].Cell(1, 1).Range.Font.TextColor.RGB != -704593921 ? "False" : str;
        }

        private static string Cau15(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Picture") && (double)(int)shape.Width == 180.0 && (int)shape.Height == 240 && shape.ShapeStyle == MsoShapeStyleIndex.msoShapeStyleNotAPreset && shape.WrapFormat.Type == WdWrapType.wdWrapSquare)
                    return "True";
            }
            return str;
        }

        private static string Cau16(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Picture") && shape.Glow.Color.RGB == 53323 && (double)shape.Glow.Radius == 8.0 && (double)(shape.Glow.Transparency * 1000f) == 600.0)
                    return "True";
            }
            return str;
        }

        private static string Cau17(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Picture") && (double)shape.Left == -999998.0 && shape.RelativeVerticalPosition == WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph && shape.RelativeHorizontalPosition == WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn)
                    return "True";
            }
            return str;
        }

        private static string Cau18(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Picture") && (int)shape.Width == 198 && (int)shape.Height == 264 && shape.ShapeStyle == MsoShapeStyleIndex.msoShapeStyleNotAPreset && shape.WrapFormat.Type == WdWrapType.wdWrapTopBottom)
                    return "True";
            }
            return str;
        }

        private static string Cau19(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Picture") && (int)shape.Width == 198 && (int)shape.Height == 264 && shape.ShapeStyle == MsoShapeStyleIndex.msoShapeStyleNotAPreset && shape.WrapFormat.Type == WdWrapType.wdWrapTopBottom && shape.PictureFormat.TransparencyColor == int.MinValue && shape.ShapeStyle == MsoShapeStyleIndex.msoShapeStyleNotAPreset)
                    return "True";
            }
            return str;
        }

        private static string Cau20(Application a, Document d)
        {
            string str = "True";
            int Index = 20;
            while (Index < 100 && !d.Paragraphs[Index].Range.Text.Contains("ARUM FAMLY (Areceae)"))
                ++Index;
            if (Index >= 100)
                return "False";
            var style = d.Paragraphs[Index].get_Style() as Style;
            return (style == null || style.NameLocal != "Heading 1") ? "False" : str;
        }

        private static string Cau21(Application a, Document d)
        {
            string str = "False";
            foreach (Bookmark bookmark in d.Bookmarks)
            {
                if (bookmark.Name == "ARUM" && bookmark.Range.Text.Contains("ARUM FAMLY (Areceae)"))
                    return "True";
            }
            return str;
        }

        private static string Cau22(Application a, Document d)
        {
            string str = "True";
            return d.Endnotes.Count != 1 || d.Endnotes[1].Reference.EndnoteOptions.NumberStyle != WdNoteNumberStyle.wdNoteNumberStyleLowercaseLetter || d.Endnotes[1].Range.Text != "Also known as Indian turnip" ? "False" : str;
        }

        private static string Cau23(Application a, Document d)
        {
            return "True";
            //string str = "True";
            //int Index = 40;
            //while (Index < 100 && !d.Paragraphs[Index].Range.Text.Contains("Flowers-- When"))
            //    ++Index;
            //if (Index >= 100)
            //    return "False";
            //var character1 = d.Paragraphs[Index].Range.Characters[3];
            //var style1 = character1.Style as Style;
            //if (style1 == null || style1.NameLocal != "Intense Emphasis")
            //    return "False";
            //while (Index < 100 && !d.Paragraphs[Index].Range.Text.Contains("Leaves: Themes"))
            //    ++Index;
            //if (Index >= 100)
            //    return "False";
            //var character2 = d.Paragraphs[Index].Range.Characters[3];
            //var style2 = character2.Style as Style;
            //if (style2 == null || style2.NameLocal != "Intense Emphasis")
            //    return "False";
            //while (Index < 100 && !d.Paragraphs[Index].Range.Text.Contains("Fruit: For example,"))
            //    ++Index;
            //if (Index >= 100)
            //    return "False";
            //var character3 = d.Paragraphs[Index].Range.Characters[3];
            //var style3 = character3.Style as Style;
            //if (style3 == null || style3.NameLocal != "Intense Emphasis")
            //    return "False";
            //while (Index < 100 && !d.Paragraphs[Index].Range.Text.Contains("Preferred Habitat--Shallow"))
            //    ++Index;
            //if (Index >= 100)
            //    return "False";
            //var character4 = d.Paragraphs[Index].Range.Characters[3];
            //var style4 = character4.Style as Style;
            //if (style4 == null || style4.NameLocal != "Intense Emphasis")
            //    return "False";
            //while (Index < 100 && !d.Paragraphs[Index].Range.Text.Contains("Flowering Season--April"))
            //    ++Index;
            //if (Index >= 100)
            //    return "False";
            //var character5 = d.Paragraphs[Index].Range.Characters[3];
            //var style5 = character5.Style as Style;
            //if (style5 == null || style5.NameLocal != "Intense Emphasis")
            //    return "False";
            //while (Index < 100 && !d.Paragraphs[Index].Range.Text.Contains("Distribution--Reading"))
            //    ++Index;
            //if (Index >= 100)
            //    return "False";
            //var character6 = d.Paragraphs[Index].Range.Characters[3];
            //var style6 = character6.Style as Style;
            //return (style6 == null || style6.NameLocal != "Intense Emphasis") ? "False" : str;
        }

        private static string Cau24(Application a, Document d)
        {
            string str = "True";
            // ISSUE: variable of a compiler-generated type
            Microsoft.Office.Interop.Word.Shape shape1 = (Microsoft.Office.Interop.Word.Shape)null;
            foreach (Microsoft.Office.Interop.Word.Shape shape2 in d.Shapes)
            {
                if (shape2.Name.Contains("Group"))
                    shape1 = shape2;
            }
            return shape1 == null ? "False" : str;
        }

        private static string Cau25(Application a, Document d)
        {
            string str = "True";
            // ISSUE: variable of a compiler-generated type
            Microsoft.Office.Interop.Word.Shape shape1 = (Microsoft.Office.Interop.Word.Shape)null;
            foreach (Microsoft.Office.Interop.Word.Shape shape2 in d.Shapes)
            {
                if (shape2.Name.Contains("Diagram"))
                    shape1 = shape2;
            }
            return shape1 == null ? "False" : str;
        }

        private static string Cau26(Application a, Document d)
        {
            string str = "False";
            foreach (Microsoft.Office.Interop.Word.Shape shape in d.Shapes)
            {
                if (shape.Name.Contains("Text") && shape.TextFrame.TextRange.Text.Contains("Plants"))
                    return "True";
            }
            return str;
        }

        private static string Cau27(Application a, Document d)
        {
            string str = "False";
            if (d.Sections[2].PageSetup.FooterDistance.ToString() != "7.2")
                return "False(Distance)";
            return d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.InlineShapes.Count == 1 && d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.InlineShapes[1].Height.ToString() == "21.8" && d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.InlineShapes[1].Width.ToString() == "34.55" ? "True" : str;
        }

        private static string Cau28(Application a, Document d)
        {
            string str = "True";
            object documentProperties = d.BuiltInDocumentProperties;
            object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder)null, documentProperties, new object[1]
            {
        (object) "Keywords"
            });
            return target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder)null, target, new object[0]).ToString() != "Flowers, blooms, Broad-leaved Arrowhead, Jack-in-the-Pulpit" ? "False(Flowers, blooms, Broad-leaved Arrowhead, Jack-in-the-Pulpit)" : str;
        }

        private static string Cau29(Application a, Document d)
        {
            string str = "True";
            if (a.Options.EnableLivePreview)
                return "False(LivePreview)";
            if (!a.Options.CheckSpellingAsYouType)
                return "False(CheckSpelling)";
            return !a.Options.UpdateFieldsWithTrackedChangesAtPrint ? "False(TrackedChanges)" : str;
        }
    }
}
