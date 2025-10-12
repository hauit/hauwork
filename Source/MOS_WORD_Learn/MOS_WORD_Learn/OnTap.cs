// Decompiled with JetBrains decompiler
// Type: MOS_WORD_LEARN.OnTap
// Assembly: MOS_WORD_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.CompilerServices;

namespace MOS_WORD_LEARN
{
    internal class OnTap
    {
        public static string CheckCau(int causo, Application a, _Document d)
        {
            switch (causo)
            {
                case 1:
                    return OnTap.cau1(a, d);
                case 2:
                    return OnTap.cau2(a, d);
                case 3:
                    return OnTap.cau3(a, d);
                case 4:
                    return OnTap.cau4(a, d);
                case 5:
                    return OnTap.cau5(a, d);
                case 6:
                    return OnTap.cau6(a, d);
                case 7:
                    return OnTap.cau7(a, d);
                case 8:
                    return OnTap.cau8(a, d);
                case 9:
                    return OnTap.cau9(a, d);
                case 10:
                    return OnTap.cau10(a, d);
                case 11:
                    return OnTap.cau11(a, d);
                case 12:
                    return OnTap.cau12(a, d);
                case 13:
                    return OnTap.cau13(a, d);
                case 14:
                    return OnTap.cau14(a, d);
                case 15:
                    return OnTap.cau15(a, d);
                case 16:
                    return OnTap.cau16(a, d);
                case 17:
                    return OnTap.cau17(a, d);
                case 18:
                    return OnTap.cau18(a, d);
                case 19:
                    return OnTap.cau19(a, d);
                case 20:
                    return OnTap.cau20(a, d);
                case 21:
                    return OnTap.cau21(a, d);
                case 22:
                    return OnTap.cau22(a, d);
                case 23:
                    return OnTap.cau23(a, d);
                case 24:
                    return OnTap.cau24(a, d);
                case 25:
                    return OnTap.cau25(a, d);
                case 26:
                    return OnTap.cau26(a, d);
                case 27:
                    return OnTap.cau27(a, d);
                case 28:
                    return OnTap.cau28(a, d);
                case 29:
                    return OnTap.cau29(a, d);
                case 30:
                    return OnTap.cau30(a, d);
                case 31:
                    return OnTap.cau31(a, d);
                case 32:
                    return OnTap.cau32(a, d);
                case 33:
                    return OnTap.cau33(a, d);
                case 34:
                    return OnTap.cau34(a, d);
                case 35:
                    return OnTap.cau35(a, d);
                case 36:
                    return OnTap.cau36(a, d);
                case 37:
                    return OnTap.cau37(a, d);
                case 38:
                    return OnTap.cau38(a, d);
                case 39:
                    return OnTap.cau39(a, d);
                case 40:
                    return OnTap.cau40(a, d);
                case 41:
                    return OnTap.cau41(a, d);
                case 42:
                    return OnTap.cau42(a, d);
                case 43:
                    return OnTap.cau43(a, d);
                case 44:
                    return OnTap.cau44(a, d);
                case 45:
                    return OnTap.cau45(a, d);
                case 46:
                    return OnTap.cau46(a, d);
                case 47:
                    return OnTap.cau47(a, d);
                case 48:
                    return OnTap.cau48(a, d);
                case 49:
                    return OnTap.cau49(a, d);
                case 50:
                    return OnTap.cau50(a, d);
                case 51:
                    return OnTap.cau51(a, d);
                case 52:
                    return OnTap.cau52(a, d);
                case 53:
                    return OnTap.cau53(a, d);
                case 54:
                    return OnTap.cau54(a, d);
                case 55:
                    return OnTap.cau55(a, d);
                case 56:
                    return OnTap.cau56(a, d);
                case 57:
                    return OnTap.cau57(a, d);
                case 58:
                    return OnTap.cau58(a, d);
                case 59:
                    return OnTap.cau59(a, d);
                case 60:
                    return OnTap.cau60(a, d);
                case 61:
                    return OnTap.cau61(a, d);
                case 62:
                    return OnTap.cau62(a, d);
                case 63:
                    return OnTap.cau63(a, d);
                case 64:
                    return OnTap.cau64(a, d);
                case 65:
                    return OnTap.cau65(a, d);
                case 66:
                    return OnTap.cau66(a, d);
                case 67:
                    return OnTap.cau67(a, d);
                case 68:
                    return OnTap.cau68(a, d);
                case 69:
                    return OnTap.cau69(a, d);
                case 70:
                    return OnTap.cau70(a, d);
                case 71:
                    return OnTap.cau71(a, d);
                case 72:
                    return OnTap.cau72(a, d);
                case 73:
                    return OnTap.cau73(a, d);
                case 74:
                    return OnTap.cau74(a, d);
                case 75:
                    return OnTap.cau75(a, d);
                case 76:
                    return OnTap.cau76(a, d);
                case 77:
                    return OnTap.cau77(a, d);
                case 78:
                    return OnTap.cau78(a, d);
                case 79:
                    return OnTap.cau79(a, d);
                case 80:
                    return OnTap.cau80(a, d);
                case 81:
                    return OnTap.cau81(a, d);
                case 82:
                    return OnTap.cau82(a, d);
                case 83:
                    return OnTap.cau83(a, d);
                case 84:
                    return OnTap.cau84(a, d);
                case 85:
                    return OnTap.cau85(a, d);
                case 86:
                    return OnTap.cau86(a, d);
                case 87:
                    return OnTap.cau87(a, d);
                case 88:
                    return OnTap.cau88(a, d);
                case 89:
                    return OnTap.cau89(a, d);
                case 90:
                    return OnTap.cau90(a, d);
                case 91:
                    return OnTap.cau91(a, d);
                case 92:
                    return OnTap.cau92(a, d);
                case 93:
                    return OnTap.cau93(a, d);
                case 94:
                    return OnTap.cau94(a, d);
                default:
                    return "default check";
            }
        }

        private static string cau1(Application a, _Document d)
        {
            try
            {
                // ISSUE: variable of a compiler-generated type
                Paragraph paragraph = (Paragraph)null;
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("Company name"))
                    {
                        paragraph = d.Paragraphs[Index];
                        break;
                    }
                }
                if (paragraph == null)
                    return "False";
                return !paragraph.Range.Text.Contains("Thắng®") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau2(Application a, _Document d)
        {
            try
            {
                if (d.Bookmarks.Count != 1)
                    return "False";
                object Index1 = (object)1;
                if (d.Bookmarks[ref Index1].Name != "HCM")
                    return "False";
                object Index2 = (object)1;
                return !d.Bookmarks[ref Index2].Range.Text.Contains("TP. HCM") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau3(Application a, _Document d)
        {
            try
            {
                return d.PageSetup.Orientation != WdOrientation.wdOrientLandscape ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau4(Application a, _Document d)
        {
            try
            {
                int num = -1;
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("Đà Lạt"))
                        num = Index;
                }
                if (num == -1)
                    return "False";
                return !d.Paragraphs[num + 1].Range.Text.Contains("đà lạt") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau5(Application a, _Document d)
        {
            try
            {
                int num = -1;
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("Đà Lạt"))
                        num = Index;
                }
                if (num == -1)
                    return "False";
                if (d.Paragraphs[num + 2].Range.ListFormat.ListString != "\uF0A7")
                    return "False";
                return d.Paragraphs[num + 2].LineSpacing.ToString() != "18" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau6(Application a, _Document d)
        {
            try
            {
                if (d.Shapes.Count != 1)
                    return "False";
                object Index1 = (object)1;
                if (!d.Shapes[ref Index1].Name.Contains("Pic"))
                    return "False";
                object Index2 = (object)1;
                if (d.Shapes[ref Index2].WrapFormat.Type != WdWrapType.wdWrapSquare)
                    return "False";
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
                object Index1 = (object)"Picture 1";
                float num = d.Shapes[ref Index1].Top;
                if (num.ToString() != "144")
                    return "False";
                object Index2 = (object)"Picture 1";
                num = d.Shapes[ref Index2].Left;
                if (num.ToString() != "432")
                    return "False";
                object Index3 = (object)"Picture 1";
                if (d.Shapes[ref Index3].RelativeHorizontalPosition.ToString() != "wdRelativeHorizontalPositionPage")
                    return "False";
                object Index4 = (object)"Picture 1";
                if (d.Shapes[ref Index4].RelativeVerticalPosition.ToString() != "wdRelativeVerticalPositionPage")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private static string cau8(Application a, _Document d) => "True";

        private static string cau9(Application a, _Document d)
        {
            try
            {
                if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Title != "Mission")
                    return "False";
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
                    return "False";
                if (d.Tables[1].Borders[WdBorderType.wdBorderBottom].Color.ToString() != "-738132122")
                    return "False";
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
                    return "False";
                if (d.Tables[1].Columns.Width.ToString() != "226.55")
                    return "False";
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
                object Index1 = (object)"Picture 2";
                if (d.Shapes[ref Index1].Left.ToString() != "-999998")
                    return "False";
                object Index2 = (object)"Picture 2";
                if (d.Shapes[ref Index2].RelativeHorizontalPosition.ToString() != "wdRelativeHorizontalPositionPage")
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
                if (d.Shapes.Count != 3)
                    return "False";
                object Index1 = (object)3;
                if (!d.Shapes[ref Index1].Name.Contains("Text"))
                    return "False";
                object Index2 = (object)3;
                return d.Shapes[ref Index2].TextFrame.TextRange.Font.ColorIndex.ToString() != "wdYellow" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau14(Application a, _Document d)
        {
            try
            {
                if (d.InlineShapes[1].SmartArt.AllNodes.Count != 3)
                    return "False";
                return d.InlineShapes[1].SmartArt.AllNodes[(object)1].TextFrame2.TextRange.Text != "Coffee" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau15(Application a, _Document d) => "True";

        private static string cau16(Application a, _Document d)
        {
            try
            {
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Controlled Observation"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                return !d.Paragraphs[Index + 2].Range.Text.Contains("24 thg 6, 2020") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau17(Application a, _Document d)
        {
            try
            {
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("f you need their natural habitat"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                return d.Paragraphs[Index + 1].Range.Tables.Count != 1 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau18(Application a, _Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau19(Application a, _Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau20(Application a, _Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau21(Application a, _Document d)
        {
            try
            {
                return !d.Paragraphs[5].Range.Text.Contains("Đông Nai") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau22(Application a, _Document d)
        {
            try
            {
                if (d.Sections[2].PageSetup.Orientation != WdOrientation.wdOrientLandscape)
                    return "False";
                return d.Sections[1].PageSetup.Orientation == WdOrientation.wdOrientLandscape ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau23(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[1].Range.Text != "\r" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau24(Application a, _Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau25(Application a, _Document d)
        {
            try
            {
                object Index1 = (object)"Rectangle 472";
                if (d.Shapes[ref Index1].TextFrame.TextRange.Text != "Word\r")
                    return "False";
                object Index2 = (object)"Rectangle 16";
                string text = d.Shapes[ref Index2].TextFrame.TextRange.Text;
                if (!text.Contains("MOS2016"))
                    return "False";
                return text.Contains("abstract") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau26(Application a, _Document d)
        {
            try
            {
                if (!d.Paragraphs[5].Range.Text.Contains("Cooker"))
                    return "False";
                return d.Paragraphs[5].Range.Font.Bold != 9999999 || d.Paragraphs[5].Alignment != WdParagraphAlignment.wdAlignParagraphLeft ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau27(Application a, _Document d)
        {
            try
            {
                int index = 1;
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("Tips for Cooking with Pressure Cooker"))
                {
                    index++;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                // Lấy Style từ Paragraph.Range
                var styleObj = d.Paragraphs[index].Range.get_Style();

                if (styleObj is Microsoft.Office.Interop.Word.Style style)
                {
                    return style.NameLocal != "Intense Reference"
                        ? "False"
                        : "True";
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


        private static string cau28(Application a, _Document d)
        {
            try
            {
                if (d.Paragraphs[4].Range.InlineShapes.Count != 1)
                    return "False";
                return d.Paragraphs[4].Range.InlineShapes[1].Type != WdInlineShapeType.wdInlineShapeSmartArt ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau29(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[2].Range.Text.Contains("cookware") || !d.Paragraphs[2].Range.Text.Contains("pan") || d.Paragraphs[6].Range.Text.Contains("cookware") || !d.Paragraphs[6].Range.Text.Contains("pan") || d.Paragraphs[7].Range.Text.Contains("cookware") || !d.Paragraphs[7].Range.Text.Contains("pan") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau30(Application a, _Document d)
        {
            try
            {
                if (d.Paragraphs[6].Range.InlineShapes.Count != 1)
                    return "False";
                return d.Paragraphs[6].Range.InlineShapes[1].Type != WdInlineShapeType.wdInlineShapeSmartArt ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau31(Application a, _Document d)
        {
            try
            {
                Paragraph para = d.Paragraphs[1];
                object styleObj = para.Range.get_Style();

                // Kiểm tra nếu style lấy được là kiểu Style thì ép kiểu
                if (styleObj is Microsoft.Office.Interop.Word.Style charStyle)
                {
                    string name = charStyle.NameLocal ?? "";
                    return name != "Intense Emphasis" ? "False" : "True";
                }
                else
                {
                    return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
        }


        private static string cau32(Application a, _Document d)
        {
            try
            {
                return d.Background.Fill.ForeColor.RGB != 16181982 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau33(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Text Box 3";
                return d.Shapes[ref Index].TextFrame.TextRange.Text != "Petrolimex\r" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau34(Application a, _Document d)
        {
            try
            {
                if (d.Tables.Count != 2)
                    return "False";
                return d.Tables[2].Rows.Count != 4 || d.Tables[2].Columns.Count != 3 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau35(Application a, _Document d)
        {
            try
            {
                if (!d.Tables[1].Rows[2].Cells[1].Range.Text.Contains("Lúa"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private static string cau36(Application a, _Document d)
        {
            try
            {
                return d.Tables[1].Borders[WdBorderType.wdBorderBottom].Color.ToString() != "-738132071" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau37(Application a, _Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau38(Application a, _Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau39(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[1].Range.InlineShapes.Count != 1 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau40(Application a, _Document d)
        {
            try
            {
                if (d.PageSetup.DifferentFirstPageHeaderFooter != -1)
                    return "False(chọn different first page)";
                return d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count != 2 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau41(Application a, _Document d)
        {
            try
            {
                if (d.Paragraphs[3].Range.InlineShapes.Count != 1)
                    return "False";
                if (d.Paragraphs[3].Range.InlineShapes[1].Width.ToString() != "460.8")
                    return "False";
                return d.Paragraphs[3].Range.InlineShapes[1].Height.ToString() != "306.7" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau42(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[8].Range.Text != "\f" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau43(Application a, _Document d)
        {
            try
            {
                object Index = (object)2;
                return !d.Shapes[ref Index].Name.Contains("7") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau44(Application a, _Document d)
        {
            try
            {
                object Index1 = (object)2;
                object Index2 = (object)"Group 193";
                return d.Shapes[ref Index2].GroupItems[ref Index1].TextFrame.TextRange.Text.Contains("Company address") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau45(Application a, _Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau46(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[12].Range.Characters[5].HighlightColorIndex != WdColorIndex.wdBrightGreen ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau47(Application a, _Document d)
        {
            try
            {
                if (d.Bookmarks.Count != 1)
                    return "False";
                object Index1 = (object)1;
                if (d.Bookmarks[ref Index1].Name != "Postscrip")
                    return "False";
                object Index2 = (object)1;
                return !d.Bookmarks[ref Index2].Range.Text.Contains("Cảnh đẹp") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau48(Application a, _Document d)
        {
            try
            {
                return d.Footnotes.Count != 1 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau49(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Picture 1";
                return d.Shapes[ref Index].AlternativeText != "Dark cocoa & Dark sweet" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau50(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[3].Range.Text.ToLower().Contains("designs") || !d.Paragraphs[3].Range.Text.ToLower().Contains("caffe") || d.Paragraphs[5].Range.Text.ToLower().Contains("designs") || !d.Paragraphs[5].Range.Text.ToLower().Contains("caffe") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau51(Application a, _Document d)
        {
            try
            {
                return d.Shapes.Count != 2 || d.Paragraphs[1].Range.Text.ToLower().Contains("hcm") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau52(Application a, _Document d)
        {
            try
            {
                return !d.Paragraphs[1].Range.Text.Contains("TP. HCM©") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau53(Application a, _Document d)
        {
            try
            {
                if (!d.Paragraphs[3].Range.Text.Contains("Vủng Tàu"))
                    return "False";
                return !d.Paragraphs[10].Range.Text.Contains("To make your document") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau54(Application a, _Document d)
        {
            try
            {
                return d.Endnotes.Count != 1 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau55(Application a, _Document d)
        {
            try
            {
                if (!d.Paragraphs[10].Range.ListFormat.ListPictureBullet.IsPictureBullet)
                    return "False";
                return !d.Paragraphs[10].Range.ListFormat.ListPictureBullet.IsPictureBullet ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau56(Application a, _Document d)
        {
            try
            {
                return !a.ActiveWindow.View.ShowAll ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau57(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs.Count != 20 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau58(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[10].Range.ListFormat.ListString != "o" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau59(Application a, _Document d)
        {
            try
            {
                if (d.Tables.Count != 1)
                    return "False";
                if (d.Tables[1].Title != "Price list")
                    return "False";
                return d.Tables[1].Descr != "Price list iPhone" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau60(Application a, _Document d)
        {
            try
            {
                if (d.Paragraphs[3].Range.Text.Contains("different galleries"))
                    return "False";
                return !d.Paragraphs[3].Range.Text.Contains("samsum") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau61(Application a, _Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau62(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[8].LeftIndent.ToString() != "54" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau63(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[1].Range.Font.TextColor.RGB.ToString() != "-738148353" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau64(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[3].Range.Characters[d.Paragraphs[3].Range.Characters.Count - 3].HighlightColorIndex != WdColorIndex.wdBrightGreen ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau65(Application a, _Document d)
        {
            try
            {
                return !d.Paragraphs[10].Range.ListFormat.ListPictureBullet.IsPictureBullet ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau66(Application a, _Document d)
        {
            try
            {
                return !d.Paragraphs[1].Range.Text.Contains("IBM®") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau67(Application a, _Document d)
        {
            try
            {
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("Reading is easier"))
                        return "False";
                }
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau68(Application a, _Document d)
        {
            try
            {
                if (d.Sections[2].PageSetup.Orientation != WdOrientation.wdOrientLandscape)
                    return "False";
                return d.Sections[1].PageSetup.Orientation == WdOrientation.wdOrientLandscape ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau69(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Text Box 2";
                return !d.Shapes[ref Index].TextFrame.TextRange.Text.Contains("Microsoft Office") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau70(Application a, _Document d)
        {
            try
            {
                if (d.Tables.Count != 1)
                    return "False";
                return d.Tables[1].Rows[5].Cells.Count != 1 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau71(Application a, _Document d)
        {
            try
            {
                if (d.Tables[1].Title != "Price list")
                    return "False";
                return !d.Tables[1].Descr.Contains("Price list of Honda scooter in") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau72(Application a, _Document d)
        {
            try
            {
                object Index1 = (object)2;
                if (!d.Shapes[ref Index1].Name.Contains("8"))
                    return "False";
                object Index2 = (object)2;
                if (d.Shapes[ref Index2].Width.ToString() != "86.4")
                    return "False";
                object Index3 = (object)2;
                if (d.Shapes[ref Index3].Height.ToString() != "86.4")
                    return "False";
                object Index4 = (object)2;
                return d.Shapes[ref Index4].Left.ToString() != "-999998" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau73(Application a, _Document d)
        {
            try
            {
                object Index = (object)"5-Point Star 2";
                return !d.Shapes[ref Index].TextFrame.TextRange.Text.Contains("2 year") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau74(Application a, _Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau75(Application a, _Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau76(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Picture 1";
                return d.Shapes[ref Index].ThreeD.RotationY.ToString() != "-44" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau77(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[13].Range.Text != "\f" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau78(Application a, _Document d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau79(Application a, _Document d)
        {
            try
            {
                return d.Tables.Count != 1 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau80(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Rectangle 197";
                return !d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[ref Index].TextFrame.TextRange.Text.Contains("MENU") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau81(Application a, _Document d)
        {
            try
            {
                return d.Bookmarks.Count != 1 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau82(Application a, _Document d)
        {
            try
            {
                return !d.Paragraphs[d.Paragraphs.Count - 3].Range.Text.Contains("\tKhông khí trong") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau83(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[1].Range.Text != "\r" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau84(Application a, _Document d)
        {
            try
            {
                if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count != 1)
                    return "False";
                object Index = (object)1;
                return !d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[ref Index].Name.Contains("Group 158") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau85(Application a, _Document d)
        {
            try
            {
                if (d.Paragraphs[9].Range.InlineShapes.Count != 1)
                    return "False";
                return d.Paragraphs[9].Range.InlineShapes[1].Type != WdInlineShapeType.wdInlineShapeSmartArt ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau86(Application a, _Document d)
        {
            try
            {
                return !d.Paragraphs[9].Range.InlineShapes[1].SmartArt.Nodes[(object)3].TextFrame2.TextRange.Text.Contains("hone") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau87(Application a, _Document d)
        {
            try
            {
                if (d.Paragraphs[3].Range.InlineShapes[1].SmartArt.QuickStyle.Name != "Intense Effect")
                    return "False";
                return d.Paragraphs[3].Range.InlineShapes[1].SmartArt.Color.Name != "Colorful Range - Accent Colors 4 to 5" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau88(Application a, _Document d)
        {
            try
            {
                if (d.Paragraphs[3].Range.Text.Contains("matching"))
                    return "False";
                return !d.Paragraphs[3].Range.Text.Contains("nights") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau89(Application a, _Document d)
        {
            try
            {
                if (d.PageSetup.DifferentFirstPageHeaderFooter != -1)
                    return "False";
                return !d.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Contains("[") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau90(Application a, _Document d)
        {
            try
            {
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("ó ý định đi Đà Lạt"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                return !d.Paragraphs[Index + 1].Range.Text.Contains("Combo detail") ? "False)" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau91(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[10].Range.ListFormat.ListString != "1." ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau92(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[10].Range.ListFormat.ListString != "1001." ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau93(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[8].Range.Text.Contains("Đà Lạt") || d.Shapes.Count != 2 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau94(Application a, _Document d)
        {
            try
            {
                return d.Paragraphs[4].Range.Text.Contains("Vủng Tàu") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
    }
}
