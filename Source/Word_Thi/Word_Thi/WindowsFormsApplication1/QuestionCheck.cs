using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Shape = Microsoft.Office.Interop.Word.Shape;

namespace MOS_WORD_TEST
{
    public interface IQuestionCheck
    {
        string CheckAnswer(int questionNumber, Application a, Document d);
    }
    internal class QuestionCheck : IQuestionCheck
    {
        public string CheckAnswer(int questionNumber, Application a, Document d)
        {
            switch (questionNumber)
            {
                case 1: return Cau1(a, d);
                case 2: return Cau2(a, d);
                case 3: return Cau3(a, d);
                case 4: return Cau4(a, d);
                case 5: return Cau5(a, d);
                case 6: return Cau6(a, d);
                case 7: return Cau7(a, d);
                case 8: return Cau8(a, d);
                case 9: return Cau9(a, d);
                case 10: return Cau10(a, d);
                case 11: return Cau11(a, d);
                case 12: return Cau12(a, d);
                case 13: return Cau13(a, d);
                case 14: return Cau14(a, d);
                case 15: return Cau15(a, d);
                case 16: return Cau16(a, d);
                case 17: return Cau17(a, d);
                case 18: return Cau18(a, d);
                case 19: return Cau19(a, d);
                case 20: return Cau20(a, d);
                case 21: return Cau21(a, d);
                case 22: return Cau22(a, d);
                case 23: return Cau23(a, d);
                case 24: return Cau24(a, d);
                case 25: return Cau25(a, d);
                case 26: return Cau26(a, d);
                case 27: return Cau27(a, d);
                case 28: return Cau28(a, d);
                case 29: return Cau29(a, d);
                case 30: return Cau30(a, d);
                case 31: return Cau31(a, d);
                case 32: return Cau32(a, d);
                case 33: return Cau33(a, d);
                case 34: return Cau34(a, d);
                case 35: return Cau35(a, d);
                case 36: return Cau36(a, d);
                case 37: return Cau37(a, d);
                case 38: return Cau38(a, d);
                case 39: return Cau39(a, d);
                case 40: return Cau40(a, d);
                case 41: return Cau41(a, d);
                case 42: return Cau42(a, d);
                case 43: return Cau43(a, d);
                case 44: return Cau44(a, d);
                case 45: return Cau45(a, d);
                case 46: return Cau46(a, d);
                case 47: return Cau47(a, d);
                case 48: return Cau48(a, d);
                case 49: return Cau49(a, d);
                case 50: return Cau50(a, d);
                case 51: return Cau51(a, d);
                case 52: return Cau52(a, d);
                case 53: return Cau53(a, d);
                case 54: return Cau54(a, d);
                case 55: return Cau55(a, d);
                case 56: return Cau56(a, d);
                case 57: return Cau57(a, d);
                case 58: return Cau58(a, d);
                case 59: return Cau59(a, d);
                case 60: return Cau60(a, d);
                case 61: return Cau61(a, d);
                case 62: return Cau62(a, d);
                case 63: return Cau63(a, d);
                case 64: return Cau64(a, d);
                case 65: return Cau65(a, d);
                case 66: return Cau66(a, d);
                case 67: return Cau67(a, d);
                case 68: return Cau68(a, d);
                case 69: return Cau69(a, d);
                case 70: return Cau70(a, d);
                case 71: return Cau71(a, d);
                case 72: return Cau72(a, d);
                case 73: return Cau73(a, d);
                case 74: return Cau74(a, d);
                case 75: return Cau75(a, d);
                case 76: return Cau76(a, d);
                case 77: return Cau77(a, d);
                case 78: return Cau78(a, d);
                case 79: return Cau79(a, d);
                case 80: return Cau80(a, d);
                case 81: return Cau81(a, d);
                case 82: return Cau82(a, d);
                case 83: return Cau83(a, d);
                case 84: return Cau84(a, d);
                case 85: return Cau85(a, d);
                case 86: return Cau86(a, d);
                case 87: return Cau87(a, d);
                case 88: return Cau88(a, d);
                case 89: return Cau89(a, d);
                case 90: return Cau90(a, d);
                case 91: return Cau91(a, d);
                case 92: return Cau92(a, d);
                case 93: return Cau93(a, d);
                case 94: return Cau94(a, d);
                case 95: return Cau95(a, d);
                case 96: return Cau96(a, d);
                case 97: return Cau97(a, d);
                case 98: return Cau98(a, d);
                case 99: return Cau99(a, d);
                case 100: return Cau100(a, d);
                case 101: return Cau101(a, d);
                case 102: return Cau102(a, d);
                case 103: return Cau103(a, d);
                case 104: return Cau104(a, d);
                case 105: return Cau105(a, d);
                case 106: return Cau106(a, d);
                case 107: return Cau107(a, d);
                case 108: return Cau108(a, d);
                case 109: return Cau109(a, d);
                case 110: return Cau110(a, d);
                case 111: return Cau111(a, d);
                case 112: return Cau112(a, d);
                case 113: return Cau113(a, d);
                case 114: return Cau114(a, d);
                case 115: return Cau115(a, d);
                case 116: return Cau116(a, d);
                case 117: return Cau117(a, d);
                case 118: return Cau118(a, d);
                case 119: return Cau119(a, d);
                case 120: return Cau120(a, d);
                case 121: return Cau121(a, d);
                case 122: return Cau122(a, d);
                case 123: return Cau123(a, d);
                case 124: return Cau124(a, d);
                case 125: return Cau125(a, d);
                default:
                    return "Invalid question number";
            }
        }

        private string Cau1(Application a, Document d)
        {
            try
            {
                if (!d.Name.StartsWith("Notes."))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau2(Application a, Document d)
        {

            try
            {
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Barstow College"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.Text != "Barstow College \vObservation Project\r")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau3(Application a, Document d)
        {
            try
            {
                if (d.DocumentTheme.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent1).RGB != 1250736)
                    return "False";
                object Index = (object)"Heading 1";
                if (d.Styles[ref Index].ParagraphFormat.Shading.BackgroundPatternColor.ToString() != "-738131969")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau4(Application a, Document d)
        {
            try
            {
                if (d.DocumentTheme.ThemeFontScheme.MajorFont.Item(MsoFontLanguageIndex.msoThemeLatin).Name != "Candara")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau5(Application a, Document d)
        {
            try
            {
                if (d.Sections.Count != 3)
                    return "False";
                if (d.Sections[2].PageSetup.Orientation != WdOrientation.wdOrientLandscape)
                    return "False";
                if (d.Sections[1].PageSetup.Orientation != WdOrientation.wdOrientPortrait)
                    return "False";
                if (d.Sections[3].PageSetup.Orientation != WdOrientation.wdOrientPortrait)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau6(Application a, Document d)
        {
            try
            {
                if (!d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains("pg. "))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau7(Application a, Document d)
        {
            try
            {
                if (!d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains("\r\r"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";

        }

        private string Cau8(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Once again,"))
                    ++Index;
                if (d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Layout.Name != "Vertical Bullet List")
                    return "False(" + d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Layout.Name + ")";
                if (d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Nodes[(object)1].TextFrame2.TextRange.Text != "Frank Miller")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau9(Application a, Document d)
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

        private string Cau10(Application a, Document d)
        {
            try
            {
                if (d.Bookmarks.Count != 0)
                    return "False";
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("NOTE TO REVIEW COMMITTEE"))
                        return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau11(Application a, Document d)
        {
            try
            {
                if (d.Hyperlinks.Count == 0)
                    return "False";
                if (d.Hyperlinks.Count > 1)
                    return "False";
                object Index1 = (object)1;
                if (!d.Hyperlinks[ref Index1].SubAddress.Contains("Introduction"))
                    return "False";
                object Index2 = (object)1;
                if (d.Hyperlinks[ref Index2].TextToDisplay != "Home")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau12(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Contents"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index + 3].Range.Text.Contains("Summary"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau13(Application a, Document d)
        {
            try
            {
                // ISSUE: reference to a compiler-generated method
                if (d.Bibliography.Sources[1].get_Field("Year") != "2001")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau14(Application a, Document d)
        {
            try
            {
                object Index = (object)"ui";
                if (a.AutoCorrect.Entries[ref Index].Value != "aut")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau15(Application a, Document d)
        {
            try
            {
                if (a.ActiveWindow.View.ShowTabs)
                    return "False";
                if (a.ActiveWindow.View.ShowSpaces)
                    return "False";
                if (!a.ActiveWindow.View.ShowHiddenText)
                    return "False";
                if (a.ActiveWindow.View.ShowAll)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau16(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (d.Paragraphs[Index].Range.Text != "Vegetables\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index - 1].Range.Text != "\f")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau17(Application a, Document d)
        {
            try
            {
                if (d.Background.Fill.ForeColor.RGB != 14282722)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau18(Application a, Document d)
        {
            try
            {
                if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text != "MOS\r\r")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau19(Application a, Document d)
        {
            try
            {
                if (d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle != WdPageNumberStyle.wdPageNumberStyleArabic)
                    return "False";
                if (!d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains(" | Page\r\r"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau20(Application a, Document d)
        {
            try
            {
                // ISSUE: variable of a compiler-generated type
                Paragraph paragraph = (Paragraph)null;
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("100% organic"))
                    {
                        paragraph = d.Paragraphs[Index];
                        break;
                    }
                }
                if (paragraph == null)
                    return "False";
                if (!paragraph.Range.ListFormat.ListString.Contains("A"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau21(Application a, Document d)
        {
            try
            {
                object Index = (object)"Picture 4";
                if (d.Shapes[ref Index].WrapFormat.Type != WdWrapType.wdWrapSquare)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau22(Application a, Document d)
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

        private string Cau23(Application a, Document d)
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
                return "False";
            }
            return "True";
        }

        private string Cau24(Application a, Document d)
        {
            try
            {
                if (d.Comments.Count != 1 || d.Comments[1].Done != true)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau25(Application a, Document d)
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
                    return "False";
                if (!d.Paragraphs[num + 2].Range.Text.ToLower().Contains("programs at other universities"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau26(Application a, Document d)
        {
            try
            {
                object Index = (object)"Text Box 2";
                if (!d.Shapes[ref Index].TextFrame.TextRange.Text.Contains("bicycles"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau27(Application a, Document d)
        {
            try
            {
                string subject = d.BuiltInDocumentProperties["Subject"].Value as string;

                if (!string.IsNullOrEmpty(subject) &&
                    subject.IndexOf("program", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "True";
                }
                else
                {
                    return "False";
                }
            }
            catch (Exception)
            {
                return "False";
            }
        }

        private string Cau28(Application a, Document d)
        {
            try
            {
                if ((double)d.Paragraphs[2].Range.InlineShapes[1].Line.Weight != 15.0)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau29(Application a, Document d)
        {
            try
            {
                if (d.Paragraphs[3].Range.Font.TextShadow.Type != MsoShadowType.msoShadow22)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau30(Application a, Document d)
        {
            try
            {
                int index1 = 1;

                // Tìm đoạn chứa "Information"
                while (index1 <= d.Paragraphs.Count &&
                       !d.Paragraphs[index1].Range.Text.Contains("Information"))
                {
                    index1++;
                }

                if (index1 > d.Paragraphs.Count)
                    return "False";

                var para1 = d.Paragraphs[index1];
                var style1 = para1.Range.get_Style() as Microsoft.Office.Interop.Word.Style;
                if (style1 == null || style1.NameLocal != "Heading 1")
                    return "False";

                // Tìm đoạn chứa "Games Times"
                int index2 = 1;
                while (index2 <= d.Paragraphs.Count &&
                       !d.Paragraphs[index2].Range.Text.Contains("Games Times"))
                {
                    index2++;
                }

                if (index2 > d.Paragraphs.Count)
                    return "False";

                var para2 = d.Paragraphs[index2];
                var style2 = para2.Range.get_Style() as Microsoft.Office.Interop.Word.Style;
                if (style2 == null || style2.NameLocal != "Heading 1")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private string Cau31(Application a, Document d)
        {
            try
            {
                int index1 = 1;

                // Tìm đoạn chứa "Online"
                while (index1 <= d.Paragraphs.Count &&
                       !d.Paragraphs[index1].Range.Text.Contains("Online"))
                {
                    index1++;
                }

                if (index1 > d.Paragraphs.Count)
                    return "False";

                var para1 = d.Paragraphs[index1];
                var style1 = para1.Range.get_Style() as Microsoft.Office.Interop.Word.Style;
                if (style1 == null || style1.NameLocal != "Heading 2")
                    return "False";

                int index2 = 1;

                // Tìm đoạn chứa "In Person"
                while (index2 <= d.Paragraphs.Count &&
                       !d.Paragraphs[index2].Range.Text.Contains("In Person"))
                {
                    index2++;
                }

                if (index2 > d.Paragraphs.Count)
                    return "False";

                var para2 = d.Paragraphs[index2];
                var style2 = para2.Range.get_Style() as Microsoft.Office.Interop.Word.Style;
                if (style2 == null || style2.NameLocal != "Heading 2")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private string Cau32(Application a, Document d)
        {
            try
            {
                if (!a.ActiveWindow.View.ShowTabs)
                    return "False";
                if (a.ActiveWindow.View.ShowSpaces)
                    return "False";
                if (a.ActiveWindow.View.ShowHiddenText)
                    return "False";
                if (a.ActiveWindow.View.ShowAll)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau33(Application a, Document d)
        {
            try
            {
                if (!a.ActiveWindow.View.ShowAll)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau34(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Tuning"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index - 1].Range.Text != "\f\r")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau35(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (d.Paragraphs[Index].Range.Text != "\u000EFour-String\r" && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.Text != "\u000EFour-String\r")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau36(Application a, Document d)
        {
            try
            {
                if (d.Sections.Count != 3)
                    return "False";
                float leftMargin = d.Sections[1].PageSetup.LeftMargin;
                if (leftMargin.ToString() != "54")
                    return "False";
                leftMargin = d.Sections[3].PageSetup.LeftMargin;
                if (leftMargin.ToString() != "54")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau37(Application a, Document d)
        {
            try
            {
                if (d.Paragraphs[6].Range.InlineShapes[1].Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectBackgroundRemoval)
                    return "False";
                if (double.Parse(d.Paragraphs[6].Range.InlineShapes[1].Fill.PictureEffects[1].EffectParameters[(object)1].Value.ToString()) >= 0.1)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau38(Application a, Document d)
        {
            try
            {
                if (d.Hyperlinks.Count == 0)
                    return "False";
                if (d.Hyperlinks.Count > 1)
                    return "False";
                object Index1 = (object)1;
                if (!d.Hyperlinks[ref Index1].Address.Contains("http://www.tailspintoys.com"))
                    return "False";
                object Index2 = (object)1;
                if (d.Hyperlinks[ref Index2].TextToDisplay != "tailspintoys.com")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau39(Application a, Document d)
        {
            try
            {
                if (d.Sections.Count != 3)
                    return "False";
                if (d.Sections[2].PageSetup.Orientation != WdOrientation.wdOrientLandscape)
                    return "False";
                if (d.Sections[1].PageSetup.Orientation != WdOrientation.wdOrientPortrait)
                    return "False";
                if (d.Sections[3].PageSetup.Orientation != WdOrientation.wdOrientPortrait)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau40(Application a, Document d)
        {
            try
            {
                int num = 0;
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("Description"))
                    {
                        num = Index;
                        break;
                    }
                }
                if (num == 0)
                    return "False";
                if (d.Paragraphs[num + 3].Range.InlineShapes.Count == 0)
                    return "Fasle";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau41(Application a, Document d)
        {
            try
            {
                //if (d.InlineShapes.Count != 1)
                //    return "False 1";
                if (d.InlineShapes[1].AlternativeText == "")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau42(Application a, Document d)
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

        private string Cau43(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Description"))
                    ++Index;
                if (Index == 0)
                    return "False";
                if (!d.Paragraphs[Index + 2].Range.Text.Contains("(Manufacturing1)"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau44(Application a, Document d)
        {
            try
            {
                foreach (Microsoft.Office.Interop.Word.Table tbl in d.Tables)
                {
                    bool isHeader = tbl.Rows.First.HeadingFormat == -1
                                    || tbl.Rows.First.Range.Bold == -1
                                    || tbl.Rows.First.Range.get_Style() is Microsoft.Office.Interop.Word.Style s && s.NameLocal.Contains("Heading");

                    //return isHeader.ToString();
                    if (isHeader)
                        return "True";

                    return "False";
                }
            }
            catch (Exception)
            {
                return "False";
            }

            return "False";
        }

        private string Cau45(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("On the Insert tab, the galleries include items ") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.Font.Bold != 0)
                    return "Falas";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau46(Application a, Document d)
        {
            try
            {
                if (d.Paragraphs[4].Range.InlineShapes[1].Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectPencilGrayscale)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau47(Application a, Document d)
        {
            try
            {
                Range FindRange(string t)
                {
                    Range r = d.Content;
                    if (r.Find.Execute(t)) return r.Duplicate;
                    return null;
                }

                Range r1 = FindRange("Contest");
                Range r2 = FindRange("The Dirty Details");
                if (r1 == null || r2 == null)
                    return "False";
                var c = r1.Font;
                var b = r2.Font;
                if (r1.Font.Name != r2.Font.Name ||
                       r1.Font.Size != r2.Font.Size ||
                       r1.Font.Bold != r2.Font.Bold ||
                       r1.Font.Italic != r2.Font.Italic ||
                       r1.Font.Underline != r2.Font.Underline ||
                       r1.Font.Color != r2.Font.Color ||
                       r1.HighlightColorIndex != r2.HighlightColorIndex)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau48(Application a, Document d)
        {
            try
            {
                int num = 0;
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("Kids love dinosaurs"))
                    {
                        num = Index;
                        break;
                    }
                }
                if (num == 0)
                    return "False";
                if (!d.Paragraphs[num + 2].Range.Text.Contains("Save time in Word with"))
                    return "False";
                if (d.Paragraphs[num + 2].Range.Font.Bold != 0)
                    return "False";
                if (d.Paragraphs[num + 2].Range.Font.TextColor.RGB != -16777216)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau49(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("Velociraptor") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.ListFormat.ListLevelNumber != 3)
                    return "Falas";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau50(Application a, Document d)
        {
            try
            {
                int num = 0;
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("Favorite dinosaurs"))
                    {
                        num = Index;
                        break;
                    }
                }
                if (num == 0)
                    return "False";
                if (d.Paragraphs[num + 3].Range.InlineShapes.Count == 0)
                    return "Fasle";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau51(Application a, Document d)
        {
            try
            {
                object Index1 = (object)"Picture 6";
                if (d.Shapes[ref Index1].Fill.PictureEffects.Count != 1)
                    return "False";
                object Index2 = (object)"Picture 6";
                if (d.Shapes[ref Index2].Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectPencilSketch)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau52(Application a, Document d)
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

        private string Cau53(Application a, Document d)
        {
            try
            {
                object Index = (object)"Text Box 3";
                if (!d.Shapes[ref Index].TextFrame.TextRange.Text.Contains("animals"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau54(Application a, Document d)
        {
            try
            {
                if (d.Sections[1].Borders[WdBorderType.wdBorderTop].LineWidth != WdLineWidth.wdLineWidth300pt)
                    return "False";
                if (d.Sections[1].Borders.Shadow)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau55(Application a, Document d)
        {
            try
            {
                if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count != 1)
                    return "False";
                object Index = (object)"Rectangle 197";
                if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[ref Index].TextFrame.TextRange.Text != "[DOCUMENT TITLE]\r")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau56(Application a, Document d)
        {
            try
            {
                if ((double)d.Paragraphs[3].Range.InlineShapes[1].SmartArt.Nodes[(object)1].Shapes.ThreeD.BevelTopDepth != 4.0)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau57(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Learning WareWolf™"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau58(Application a, Document d)
        {
            try
            {
                int indexRecycle = -1;
                int indexRework = -1;
                foreach (Shape shape in d.Shapes)
                {
                    if (shape.Type != Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                    {
                        continue;
                    }
                    var nodes = shape.SmartArt.AllNodes;
                    for (int i = 1; i <= nodes.Count; i++)
                    {
                        var node = nodes[i];
                        if (node.TextFrame2.TextRange.Text.Trim().Contains("Recycle"))
                        {
                            indexRecycle = i;
                            continue;
                        }

                        if (node.TextFrame2.TextRange.Text.Trim().Contains("Rework"))
                        {
                            indexRework = i;
                            continue;
                        }
                    }
                }

                if ((indexRecycle + indexRework) > 0 && (indexRecycle < indexRework))
                    return "True";

                return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private string Cau59(Application a, Document d)
        {
            try
            {
                int index = 1;

                // Tìm đoạn chứa "My Coffee"
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("My Coffee"))
                {
                    index++;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                // Dùng dynamic để truy cập CharacterStyle
                Paragraph para = d.Paragraphs[index];
                dynamic rng = para.Range;
                var style = rng.CharacterStyle as Microsoft.Office.Interop.Word.Style;

                if (style != null)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private string Cau60(Application a, Document d)
        {
            try
            {
                object Index = (object)"Picture 3";
                if (d.Shapes[ref Index].WrapFormat.Type == WdWrapType.wdWrapSquare &&
                    d.Shapes[ref Index].WrapFormat.Side == WdWrapSideType.wdWrapLeft)
                    return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "False";
        }

        private string Cau61(Application a, Document d)
        {
            try
            {
                object Index = (object)"Picture 3";
                string str = d.Shapes[ref Index].ThreeD.BevelTopInset.ToString();
                if (str != "12")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau62(Application a, Document d)
        {
            try
            {
                if (d.Paragraphs[7].Range.InlineShapes[1].SmartArt.Color.Name != "Transparent Gradient Range - Accent 1")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau63(Application a, Document d)
        {
            try
            {
                object Index = (object)"Diagram 1";
                if (d.Shapes[ref Index].SmartArt.Nodes[(object)2].TextFrame2.TextRange.Text != "organically grown")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau64(Application a, Document d)
        {
            try
            {
                if ((double)d.Paragraphs[1].Range.Font.Size != 11.0)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau65(Application a, Document d)
        {
            try
            {
                if (d.Paragraphs[1].Range.Font.Name != "Algerian")
                    return "False";
                if (d.Paragraphs[1].Range.Font.Underline != WdUnderline.wdUnderlineThick)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau66(Application a, Document d)
        {
            try
            {
                if (!d.Paragraphs[6].Range.Text.Contains("Fourth Coffee"))
                    return "False";
                if (d.Paragraphs[6].Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphLeft)
                    return "False";
                if (d.Paragraphs[6].Range.Font.Bold != 9999999)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau67(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("Focus on the text you want") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.ListFormat.ListString != "▸")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau68(Application a, Document d)
        {
            try
            {
                if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Characters[3].Font.TextColor.RGB != -738131969)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau69(Application a, Document d)
        {
            try
            {
                if (d.Paragraphs[4].Range.Text.Contains("in the embed code"))
                    return "False";
                if (!d.Paragraphs[6].Range.Text.Contains("new look"))
                    return "False";
                if (d.Paragraphs[8].Range.Text.Contains("and SmartArt"))
                    return "False";
                if (d.Paragraphs[10].Range.Characters[5].Font.Bold == -1)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau70(Application a, Document d)
        {
            try
            {
                if (d.Footnotes.Count != 1)
                    return "False";
                if (!d.Footnotes[1].Range.Text.Contains("Free to join"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau71(Application a, Document d)
        {
            try
            {
                if (d.Sections.Count != 3)
                    return "False";
                if (d.Sections[2].PageSetup.TextColumns.Count != 2)
                    return "False";
                if (d.Sections[2].PageSetup.TextColumns.Spacing.ToString() != "21.6")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau72(Application a, Document d)
        {
            //try
            //{
            //    int Index = 1;
            //    while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Barstow College"))
            //        ++Index;
            //    if (Index >= d.Paragraphs.Count)
            //        return "False 1";
            //    if (d.Paragraphs[Index].Range.Text != "Barstow College \vObservation Project\r")
            //        return "False 2";
            //}
            //catch (Exception ex)
            //{
            //    return "False" + ex.Message;
            //}
            //return "True";

            try
            {
                if (d.Sections[1].Borders[WdBorderType.wdBorderTop].LineWidth != WdLineWidth.wdLineWidth300pt)
                    return "False";
                if (d.Sections[1].Borders.Shadow)
                    return "False";
                if (d.Sections[1].Borders[WdBorderType.wdBorderTop].Color.ToString() != "-738131969")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau73(Application a, Document d)
        {
            try
            {
                int index = 1;

                // Tìm đoạn chứa câu cần kiểm tra
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("This event begins at 7 pm on "))
                {
                    index++;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                // Truy cập dynamic để lấy CharacterStyle
                var para = d.Paragraphs[index];
                dynamic rng = para.Range;
                var style = rng.CharacterStyle as Microsoft.Office.Interop.Word.Style;

                if (style == null || style.NameLocal != "Intense Emphasis")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private string Cau74(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("margie@margiestravel.com") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.ParagraphFormat.LineSpacing.ToString() != "14")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau75(Application a, Document d)
        {
            try
            {
                if (d.CompatibilityMode == 11)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau76(Application a, Document d)
        {
            try
            {
                if (d.RemovePersonalInformation)
                    return "False";
                if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count != 0)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau77(Application a, Document d)
        {
            try
            {
                if (d.Name != "Memo.txt")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau78(Application a, Document d)
        {
            try
            {
                // ISSUE: variable of a compiler-generated type
                Paragraph paragraph = (Paragraph)null;
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
                    return "False";
                if (paragraph.Range.HighlightColorIndex != WdColorIndex.wdTurquoise || d.Paragraphs[Index + 1].Range.HighlightColorIndex != WdColorIndex.wdTurquoise || d.Paragraphs[Index + 2].Range.HighlightColorIndex != WdColorIndex.wdTurquoise)
                    return "False";
                if (d.Paragraphs[Index + 3].Range.HighlightColorIndex != WdColorIndex.wdTurquoise)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau79(Application a, Document d)
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

        private string Cau80(Application a, Document d)
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
                    return "False";
                if (d.Paragraphs[num - 1].Range.Text != "\f")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau81(Application a, Document d)
        {
            try
            {
                if (d.Sections.Count != 3)
                    return "False";
                if (d.Sections[2].PageSetup.TextColumns.Count != 2)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau82(Application a, Document d)
        {
            try
            {
                object Index = (object)"Heading 1";
                if (d.Styles[ref Index].Borders[WdBorderType.wdBorderBottom].Color.ToString() != "-721354753")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau83(Application a, Document d)
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
                return "False";
            return d.Paragraphs[num + 1].Range.InlineShapes.Count != 1 ? "False" : "True";
        }

        private string Cau84(Application a, Document d)
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

        private string Cau85(Application a, Document d)
        {
            try
            {
                if (!d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains("\r\a [DOCUMENT TITLE]\a\a\r"))
                    return "False";
                if (d.Sections[1].PageSetup.DifferentFirstPageHeaderFooter != -1)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau86(Application a, Document d)
        {
            try
            {
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.ParagraphFormat.LineSpacing.ToString() != "16.8")
                        return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau87(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("The picture fits in your document") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.ListFormat.ListString != "5.")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau88(Application a, Document d)
        {
            try
            {
                object Index = (object)"Picture 11";
                if (d.Shapes[ref Index].WrapFormat.Type != WdWrapType.wdWrapSquare)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau89(Application a, Document d)
        {
            try
            {
                if (d.InlineShapes.Count != 2)
                    return "False";
                if ((double)d.InlineShapes[1].SmartArt.Nodes[(object)1].Shapes.ThreeD.BevelTopDepth != 4.0)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau90(Application a, Document d)
        {
            try
            {
                int Index1 = 0;
                for (int Index2 = 1; Index2 < d.Paragraphs.Count; ++Index2)
                {
                    if (d.Paragraphs[Index2].Range.Text.Contains("The muffin tray will still be hot"))
                    {
                        Index1 = Index2;
                        break;
                    }
                }
                if (Index1 == 0)
                    return "False";
                if (d.Paragraphs[Index1].Range.Characters[1].Text != "(")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau91(Application a, Document d)
        {
            try
            {
                if (d.Sections[2].PageSetup.Orientation != WdOrientation.wdOrientLandscape)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau92(Application a, Document d)
        {
            // ISSUE: variable of a compiler-generated type
            Microsoft.Office.Interop.Word.Shape shape;
            try
            {
                object Index = (object)"Group 193";
                shape = d.Shapes[ref Index];
            }
            catch (Exception ex)
            {
                return "False";
            }
            try
            {
                object Index = (object)"Rectangle 195";
                // ISSUE: variable of a compiler-generated type
                Microsoft.Office.Interop.Word.Shape groupItem = shape.GroupItems[ref Index];
                if (groupItem.TextFrame.TextRange.Text.Contains("Company address"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau93(Application a, Document d)
        {
            try
            {
                int index = 1;

                // Tìm đoạn có chứa đoạn văn bản chỉ định
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("These distractions could be on the"))
                {
                    index++;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                // Lấy style của đoạn văn bản đó
                Paragraph para = d.Paragraphs[index];
                var styleObj = para.Range.get_Style(); // hoặc .Style nếu không bị lỗi
                var style = styleObj as Microsoft.Office.Interop.Word.Style;

                if (style == null)
                    return "False";

                if (style.NameLocal != "Intense Emphasis")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private string Cau94(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Keep it Simple"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (!d.Paragraphs[Index + 1].Range.Text.Contains("you specify directly."))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau95(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Rehearse and Video Your Presentation"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (!d.Paragraphs[Index + 3].Range.Text.Contains("Summarize Main Points"))
                    return "False";
                if (!d.Paragraphs[Index + 5].Range.Text.Contains("You can easily change the formatting"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau96(Application a, Document d)
        {
            try
            {
                Range content = d.Content;
                Find findHeading = content.Find;
                findHeading.Text = "Summarize main points";
                if (!findHeading.Execute())
                    return "False";

                Range afterHeading = d.Range(content.End, d.Content.End);
                Paragraph firstPara = afterHeading.Paragraphs[1];
                if (firstPara == null)
                    return "False";

                //for(int i = 1; i <= afterHeading.Paragraphs.Count; i++)
                //{
                //    Range paraRange = afterHeading.Paragraphs[i].Range;
                //    Find findWord = paraRange.Find;
                //    findWord.Text = "look";
                //    if (!findWord.Execute())
                //    {
                //        continue;
                //    }
                //}
                Range paraRange = afterHeading.Paragraphs[2].Range;

                Find findWord = paraRange.Find;
                findWord.Text = "look";
                if (!findWord.Execute())
                    return "False";

                // Nếu tìm thấy từ "look", kiểm tra footnote "resource"
                bool hasResourceFootnote = false;
                foreach (Footnote fn in d.Footnotes)
                {
                    string fnText = fn.Range?.Text ?? string.Empty;
                    if (fnText.ToLower().IndexOf("resource") >= 0)
                    {
                        hasResourceFootnote = true;
                        break;
                    }
                }

                if (hasResourceFootnote)
                    return "True";
                else
                    return "False";
            }
            catch (Exception e)
            {
                return "False";
            }
        }

        private string Cau97(Application a, Document d)
        {
            try
            {
                object Index = (object)"Heading 1";
                if (d.Styles[ref Index].ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau98(Application a, Document d)
        {
            try
            {
                int index = 1;

                // Tìm đoạn chứa "Genre"
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("Genre"))
                {
                    index++;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                var para1 = d.Paragraphs[index];
                dynamic rng1 = para1.Range;
                var style1 = rng1.CharacterStyle as Microsoft.Office.Interop.Word.Style;

                if (style1 == null || style1.NameLocal != "Subtle Emphasis")
                    return "False";

                // Tìm đoạn chứa "External Parther"
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("External Parther"))
                {
                    index++;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                var para2 = d.Paragraphs[index];
                dynamic rng2 = para2.Range;
                var style2 = rng2.CharacterStyle as Microsoft.Office.Interop.Word.Style;

                if (style2 == null || style2.NameLocal != "Subtle Emphasis")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private string Cau99(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("Corporate events") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.ListFormat.ListString != "\uF0B7")
                    return "False";
                while (!d.Paragraphs[Index].Range.Text.Contains("Online events") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.ListFormat.ListString != "\uF0B7")
                    return "False";
                while (!d.Paragraphs[Index].Range.Text.Contains("weddings") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.ListFormat.ListString != "")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau100(Application a, Document d)
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

        private string Cau101(Application a, Document d)
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

        private string Cau102(Application a, Document d)
        {
            try
            {
                int Index1 = 0;
                int num = 0;
                int Index2 = 0;
                for (int Index3 = 1; Index3 < d.Paragraphs.Count; ++Index3)
                {
                    if (d.Paragraphs[Index3].Range.Text.Contains("powerful new"))
                        Index1 = Index3;
                    if (d.Paragraphs[Index3].Range.Text.Contains("Themes and styles also help keep your"))
                        num = Index3;
                    if (d.Paragraphs[Index3].Range.Text.Contains("1-2 videographers"))
                        Index2 = Index3;
                }
                if (Index1 == 0)
                    return "False";
                if (num == 0 || Index2 == 0)
                    return "False";
                if (!d.Paragraphs[Index1].Range.Text.Contains("powerful new"))
                    return "False";
                if (d.Paragraphs[49].Range.Text.Contains("Themes and styles also help keep your"))
                    return "False";
                if (d.Paragraphs[Index2].Range.Characters[5].Font.Bold == -1)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau103(Application a, Document d)
        {
            try
            {
                if (d.Footnotes.Count != 1)
                    return "False";
                if (!d.Footnotes[1].Range.Text.Contains("Includes digital files."))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau104(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("Algebra II") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.ListFormat.ListString != "111.")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau105(Application a, Document d)
        {
            try
            {
                if (d.Shapes.Count != 2)
                    return "False";
                object Index1 = (object)2;
                if (!d.Shapes[ref Index1].Name.Contains("Horizontal Scroll"))
                    return "False";
                object Index2 = (object)2;
                if (d.Shapes[ref Index2].TextFrame.TextRange.Text.Trim() != "Remember your calculator!")
                    return "False";
                object Index3 = (object)2;
                if (d.Shapes[ref Index3].WrapFormat.Type != WdWrapType.wdWrapSquare)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau106(Application a, Document d)
        {
            try
            {
                if (d.Shapes.Count != 1)
                    return "False";
                object Index = (object)"Picture 1";
                if (d.Shapes[ref Index].Line.ForeColor.RGB != 11957550)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau107(Application a, Document d)
        {
            try
            {
                if (d.InlineShapes[1].SmartArt.Reverse == MsoTriState.msoTrue)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau108(Application a, Document d)
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
                if ((double)d.Tables[1].Columns[1].Width >= (double)d.Tables[1].Columns[2].Width)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau109(Application a, Document d)
        {
            try
            {
                if (d.Bookmarks.Count == 0)
                    return "False";
                if (d.Bookmarks.Count > 1)
                    return "False";
                object Index = (object)1;
                if (!d.Bookmarks[ref Index].Name.Contains("Cheating"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau110(Application a, Document d)
        {
            try
            {
                if (d.Sections.Count != 1)
                    return "False";
                if ((double)d.Sections[1].PageSetup.TopMargin != 54.0)
                    return "False";
                if ((double)d.Sections[1].PageSetup.BottomMargin != 54.0)
                    return "False";
                if ((double)d.Sections[1].PageSetup.LeftMargin != 36.0)
                    return "False";
                if ((double)d.Sections[1].PageSetup.RightMargin != 36.0)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau111(Application a, Document d)
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

        private string Cau112(Application a, Document d)
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
                    return "False";
                if (d.Paragraphs[Index1].Range.Font.Bold == -1)
                    return "False";
                if (d.Paragraphs[Index1].Range.Font.Italic == -1)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau113(Application a, Document d)
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
                return "False";
            }
            return "True";
        }

        private string Cau114(Application a, Document d)
        {
            try
            {
                int okTables = 0;

                foreach (Microsoft.Office.Interop.Word.Table tbl in d.Tables)
                {
                    bool isHeader = tbl.Rows.First.HeadingFormat == -1
                                    || tbl.Rows.First.Range.Bold == -1
                                    || tbl.Rows.First.Range.get_Style() is Microsoft.Office.Interop.Word.Style s && s.NameLocal.Contains("Heading");

                    //return isHeader.ToString();
                    if (isHeader)
                        okTables++;
                }

                return okTables > 1 ? "True" : "False";
            }
            catch (Exception)
            {
                return "False";
            }
            //return "False";
        }

        private string Cau115(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("Standard") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.ListFormat.ListType != WdListType.wdListPictureBullet)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau116(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("Mos360.vn") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if ((double)d.Paragraphs[Index].Range.ParagraphFormat.LineSpacing != 18.0)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau117(Application a, Document d)
        {
            try
            {
                int Index = 1;
                while (!d.Paragraphs[Index].Range.Text.Contains("Mandolin Style") && Index < d.Paragraphs.Count)
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index].Range.ListFormat.ListValue != 11)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau118(Application a, Document d)
        {
            try
            {
                object count1 = (object)d.Shapes.Count;
                if (d.Shapes[ref count1].TextFrame.TextRange.Text.Contains("We Heard you..."))
                    return "False";
                object count2 = (object)d.Shapes.Count;
                if (d.Shapes[ref count2].TextFrame.TextRange.Font.TextColor.RGB.ToString() != "-16777216")
                    return "False";
                object count3 = (object)d.Shapes.Count;
                if (d.Shapes[ref count3].Left.ToString() != "-999995")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau119(Application a, Document d)
        {
            try
            {
                if (d.Tables.Count != 9)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau120(Application a, Document d)
        {
            try
            {
                object Index = (object)"Rectangle 4";
                if (d.Shapes[ref Index].TextFrame.TextRange.Text.Trim() != "ANYTIME ACCOUNT ACCESS")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau121(Application a, Document d)
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

        private string Cau122(Application a, Document d)
        {
            try
            {
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("automatic"))
                        return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau123(Application a, Document d)
        {
            try
            {
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.ToUpper().Contains("WOODGROVE PLUS SAVINGS"))
                        return "True";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "False";
        }

        private string Cau124(Application a, Document d)
        {
            try
            {
                if (d.Comments.Count != 0)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private string Cau125(Application a, Document d)
        {
            try
            {
                if (d.Paragraphs[2].Range.Text != "Contents\r")
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
