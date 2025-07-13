// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Design
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Drawing;

namespace WindowsFormsApplication1
{
    internal class Design
    {
        public static string CheckCau(int causo, Application a, _Document d)
        {
            switch (causo)
            {
                case 1:
                    return Design.cau12(a, d);
                case 2:
                    return Design.cau14(a, d);
                case 3:
                    return Design.cau10(a, d);
                case 4:
                    return Design.cau8(a, d);
                case 5:
                    return Design.cau13(a, d);
                case 6:
                    return Design.cau9(a, d);
                case 7:
                    return Design.cau3(a, d);
                case 8:
                    return Design.cau11(a, d);
                case 9:
                    return Design.cau5(a, d);
                case 10:
                    return Design.cau10(a, d);
                case 11:
                    return Design.cau11(a, d);
                case 12:
                    return Design.cau12(a, d);
                case 13:
                    return Design.cau13(a, d);
                case 14:
                    return Design.cau14(a, d);
                case 15:
                    return Design.cau15(a, d);
                default:
                    return "False(Default sec 1)";
            }
        }

        private static string cau1(Application a, _Document d)
        {
            try
            {
                WdColor expectedColor = (WdColor)ColorTranslator.ToOle(Color.FromArgb(255, 128, 128));
                if (d.Sections[1].Borders[WdBorderType.wdBorderTop].Color != expectedColor)
                    return "False (LineColor)";
                if (d.Sections[1].Borders.Shadow)
                    return "False (shadow)";
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
                if (d.DocumentTheme.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent1).RGB != 1250736)
                    return "False(Chon them Ion)";
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
                if (d.Background.Fill.ForeColor.RGB != 14282722)
                    return "False (" + (object)d.Background.Fill.ForeColor.RGB + ")";
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
                WdColor expectedColor = (WdColor)ColorTranslator.ToOle(Color.FromArgb(255, 128, 128));
                if (d.Sections[1].Borders[WdBorderType.wdBorderTop].Color != expectedColor)
                    return "False (LineColor)";
                if (d.Sections[1].Borders.Shadow)
                    return "False (shadow)";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau5(Application a, _Document d)
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
                return "False (Insert Banded cover page)";
            }
            try
            {
                object Index = (object)"Rectangle 195";
                // ISSUE: variable of a compiler-generated type
                Microsoft.Office.Interop.Word.Shape groupItem = shape.GroupItems[ref Index];
                if (groupItem.TextFrame.TextRange.Text.Contains("Company address"))
                    return "False(delete Company address)";
            }
            catch (Exception ex)
            {
                return "False(something wrong)";
            }
            return "True";
        }

        private static string cau6(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Rectangle 16";
                if (!d.Shapes[ref Index].TextFrame.TextRange.Text.Contains("[DOCUMENT TITLE]"))
                    return "False(chen trang bia Gird)";
            }
            catch (Exception ex)
            {
                return "False (sai kieu trang bia)";
            }
            return "True";
        }

        private static string cau7(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Group 193";
                // ISSUE: variable of a compiler-generated type
                Microsoft.Office.Interop.Word.Shape shape = d.Shapes[ref Index];
            }
            catch (Exception ex)
            {
                return "False (sai kieu trang bia)";
            }
            return "True";
        }

        private static string cau8(Application a, _Document d)
        {
            try
            {
                if (d.DocumentTheme.ThemeFontScheme.MajorFont.Item(MsoFontLanguageIndex.msoThemeLatin).Name != "Candara")
                    return "False(sai theme font)";
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
                if (d.Sections[1].Borders[WdBorderType.wdBorderTop].LineWidth != WdLineWidth.wdLineWidth300pt)
                    return "False (do day duong vien)";
                if (d.Sections[1].Borders.Shadow)
                    return "False (shadow)";
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
                if (d.DocumentTheme.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent1).RGB != 1250736)
                    return "False(chon theme Ion )";
                object Index = (object)"Heading 1";
                if (d.Styles[ref Index].ParagraphFormat.Shading.BackgroundPatternColor.ToString() != "-738131969")
                    return "False(sai kieu Formating)";
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
                if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count != 1)
                    return "False(chen WaterMark)";
                object Index = (object)1;
                if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes[ref Index].TextEffect.Text != "DO NOT COPY")
                    return "False(sai kieu)";
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
                object Index = (object)"Heading 1";
                if (d.Styles[ref Index].ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                    return "False(Centered)";
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
                if (d.Sections[1].Borders[WdBorderType.wdBorderTop].LineWidth != WdLineWidth.wdLineWidth300pt)
                    return "False (do day duong vien)";
                if (d.Sections[1].Borders.Shadow)
                    return "False (không shadow)";
                if (d.Sections[1].Borders[WdBorderType.wdBorderTop].Color.ToString() != "-738131969")
                    return "False(màu)";
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
                object Index = (object)"Heading 1";
                if (d.Styles[ref Index].Borders[WdBorderType.wdBorderBottom].Color.ToString() != "-721354753")
                    return "False(sai kiểu)";
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
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }
    }
}
