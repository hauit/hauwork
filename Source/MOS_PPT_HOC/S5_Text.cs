// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S5_Text
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Chart = Microsoft.Office.Interop.PowerPoint.Chart;


namespace MOS_PPT_LEARN
{
    internal class S5_Text
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S5_Text.cau1(a, d);
                case 2:
                    return S5_Text.cau2(a, d);
                case 3:
                    return S5_Text.cau3(a, d);
                case 4:
                    return S5_Text.cau4(a, d);
                case 5:
                    return S5_Text.cau5(a, d);
                case 6:
                    return S5_Text.cau6(a, d);
                case 7:
                    return S5_Text.cau7(a, d);
                case 8:
                    return S5_Text.cau8(a, d);
                case 9:
                    return S5_Text.cau9(a, d);
                case 10:
                    return S5_Text.cau10(a, d);
                case 11:
                    return S5_Text.cau11(a, d);
                case 12:
                    return S5_Text.cau12(a, d);
                case 13:
                    return S5_Text.cau13(a, d);
                case 14:
                    return S5_Text.cau14(a, d);
                default:
                    return "out Indext";
            }
        }

        private static string ReadWorksheetCellText(Worksheet worksheet, string address)
        {
            try
            {
                if (worksheet == null)
                    return string.Empty;
                // follow the decompiled style using object parameters
                var rng = worksheet.get_Range((object)address, (object)address) as Range;
                if (rng == null)
                    return string.Empty;
                var txt = rng.Text;
                return txt?.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string cau1(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)5].Shapes[(object)3].TextFrame.TextRange.Text != "Company Confidential")
                    return "False";
                if (!a.ActivePresentation.Slides[(object)5].Shapes[(object)3].Name.Contains("Footer Placeholder"))
                    return "False";
                if (a.ActivePresentation.Slides[(object)6].Shapes.Count >= 3)
                {
                    if (a.ActivePresentation.Slides[(object)6].Shapes[(object)3].TextFrame.TextRange.Text == "Company Confidential")
                        return "False";
                    if (a.ActivePresentation.Slides[(object)6].Shapes[(object)3].Name.Contains("Footer Placeholder"))
                        return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private static string cau2(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)3].Shapes[(object)"Content Placeholder 3"].Table.Style.Name != "Medium Style 1 - Accent 5" ? "False(Medium Style 1 - Accent 5)" : "True";
            }
            catch (Exception ex)
            {
                return "False (loi khong xac dinh)";
            }
        }

        private static string cau3(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Hyperlinks.Count != 1)
                    return "False";
                if (!a.ActivePresentation.Slides[(object)3].Hyperlinks[1].Address.Contains("humongousinsurance.com"))
                    return "False";
                return a.ActivePresentation.Slides[(object)3].Hyperlinks[1].TextToDisplay != "Click here to view on website" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau4(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Comments.Count != 1)
                    return "False";
                if (a.ActivePresentation.Slides[(object)2].Comments[1].Text != "Update")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private static string cau5(Application a, Presentation d)
        {
            try
            {
                if (((object)a.ActivePresentation.Slides[(object)3].Shapes[(object)"Content Placeholder 4"].Type).ToString() != "msoPlaceholder")
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Content Placeholder 4"].Table.Rows.Count.ToString() != "6")
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Content Placeholder 4"].Table.Rows[3].Cells[1].Shape.TextFrame.TextRange.Text.Contains("Sinusitis"))
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Content Placeholder 4"].Table.Columns.Count.ToString() != "5")
                    return "False";
                return !a.ActivePresentation.Slides[(object)3].Shapes[(object)"Content Placeholder 4"].Table.Columns[5].Cells[1].Shape.TextFrame.TextRange.Text.Contains("Percentage Uninsured") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau6(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)7].Shapes[(object)"Table 1"].Table.Style.Name != "Medium Style 2 - Accent 1")
                    return "False";
                if (a.ActivePresentation.Slides[(object)7].Shapes[(object)"Table 1"].Table.HorizBanding.ToString() != "False")
                    return "False";
                return a.ActivePresentation.Slides[(object)7].Shapes[(object)"Table 1"].Table.VertBanding.ToString() != "True" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau7(Application a, Presentation d)
        {
            try
            {
                int count = a.ActivePresentation.Slides[(object)3].Shapes[(object)"Content Placeholder 6"].Table.Columns.Count;
                if (count.ToString() != "3")
                    return "False (xoa cot)";
                count = a.ActivePresentation.Slides[(object)3].Shapes[(object)"Content Placeholder 6"].Table.Rows.Count;
                return count.ToString() != "7" ? "Fales (them dong)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string cau8(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 2)
                    return "False";
                try
                {
                    if (a.ActivePresentation.Slides[(object)4].Shapes[(object)2].SmartArt.Layout.Name != "Vertical Picture Accent List")
                        return "False";
                }
                catch (Exception ex)
                {
                    return "False";
                }
                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau9(Application a, Presentation d)
        {
            try
            {
                if (((object)a.ActivePresentation.Slides[(object)2].Shapes[(object)"TextBox 3"].TextFrame.VerticalAnchor).ToString() != "msoAnchorTop")
                    return "False";
                return ((object)a.ActivePresentation.Slides[(object)2].Shapes[(object)"TextBox 3"].TextFrame2.TextRange.get_Characters().Font.Caps).ToString() != "msoSmallCaps" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau10(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)5].Shapes.Count != 6)
                    return "False";
                return !a.ActivePresentation.Slides[(object)5].Shapes[(object)6].Name.Contains("nk") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau11(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)7].Shapes[(object)2].TextFrame.TextRange.Font.Color.RGB.ToString() != "6968388" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau12(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)"Content Placeholder 3"].Table.Rows.Count.ToString() != "11")
                    return "False (number of rows)";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)"Content Placeholder 3"].Table.Rows[10].Cells[1].Shape.TextFrame.TextRange.Text != "Z1")
                    return "False (delete wrong row)";
                string name = a.ActivePresentation.Slides[(object)4].Shapes[(object)"Content Placeholder 3"].Table.Style.Name;
                return name != "Light Style 1 - Accent 1" ? $"False ({name})" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string cau13(Application a, Presentation d)
        {
            try
            {
                string name = a.ActivePresentation.Slides[(object)5].Shapes[(object)"Content Placeholder 3"].Table.Style.Name;
                return name != "Light Style 1 - Accent 1" ? $"False ({name})" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string cau14(Application a, Presentation d)
        {
            try
            {
                string result = "True";
                Slide slide = a.ActivePresentation.Slides[(object)3];
                Chart chart = slide.Shapes[(object)"Content Placeholder 3"].Chart;
                if (chart == null)
                    return "False (Something wrong)";

                // Access Excel workbook/worksheet used by chart data
                var rawWorkbook = chart.ChartData?.Workbook;
                Workbook workbook = rawWorkbook as Workbook;
                if (workbook == null)
                    return "False (chart workbook missing)";

                Worksheet worksheet = null;
                try
                {
                    worksheet = workbook.Worksheets[1] as Worksheet;
                }
                catch
                {
                    worksheet = null;
                }

                if (worksheet == null)
                    return "False (worksheet missing)";

                string f1 = ReadWorksheetCellText(worksheet, "F1");
                if (f1 != "Site 5")
                    return "False (Cell F1)";

                string f2 = ReadWorksheetCellText(worksheet, "F2");
                if (f2 != "46%")
                    return "False (Cell F2)";

                return chart.Legend != null && chart.Legend.Width >= 200.0 ? result : "False (Not include Site 5)";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string cau15(Application a, Presentation d)
        {
            try
            {
                string result = "True";
                Slide slide = a.ActivePresentation.Slides[(object)3];
                Chart chart = slide.Shapes[(object)"Content Placeholder 3"].Chart;
                if (chart == null)
                    return "False (Something wrong)";

                var rawWorkbook = chart.ChartData?.Workbook;
                Workbook workbook = rawWorkbook as Workbook;
                if (workbook == null)
                    return "False (chart workbook missing)";

                Worksheet worksheet = null;
                try
                {
                    worksheet = workbook.Worksheets[1] as Worksheet;
                }
                catch
                {
                    worksheet = null;
                }

                if (worksheet == null)
                    return "False (worksheet missing)";

                string f1 = ReadWorksheetCellText(worksheet, "F1");
                if (f1 != "Site 5")
                    return "False (Cell F1)";

                string f2 = ReadWorksheetCellText(worksheet, "F2");
                if (f2 != "46%")
                    return "False (Cell F2)";

                return chart.Legend != null && chart.Legend.Width >= 200.0 ? result : "False (Not include Site 5)";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string cau16(Application a, Presentation d)
        {
            try
            {
                string result = "True";
                Slide slide = a.ActivePresentation.Slides[(object)3];
                Chart chart = slide.Shapes[(object)"Content Placeholder 3"].Chart;
                if (chart == null)
                    return "False (Something wrong)";

                var rawWorkbook = chart.ChartData?.Workbook;
                Workbook workbook = rawWorkbook as Workbook;
                if (workbook == null)
                    return "False (chart workbook missing)";

                Worksheet worksheet = null;
                try
                {
                    worksheet = workbook.Worksheets[1] as Worksheet;
                }
                catch
                {
                    worksheet = null;
                }

                if (worksheet == null)
                    return "False (worksheet missing)";

                string f1 = ReadWorksheetCellText(worksheet, "F1");
                if (f1 != "Site 5")
                    return "False (Cell F1)";

                string f2 = ReadWorksheetCellText(worksheet, "F2");
                if (f2 != "46%")
                    return "False (Cell F2)";

                return chart.Legend != null && chart.Legend.Width >= 200.0 ? result : "False (Not include Site 5)";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }
    }
}