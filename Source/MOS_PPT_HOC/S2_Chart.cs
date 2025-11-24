// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S2_Chart
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using XlLegendPosition = Microsoft.Office.Interop.PowerPoint.XlLegendPosition;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Chart = Microsoft.Office.Interop.PowerPoint.Chart;
using XlChartType = Microsoft.Office.Core.XlChartType;


namespace MOS_PPT_LEARN
{
    internal class S2_Chart
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S2_Chart.cau1(a, d);
                case 2:
                    return S2_Chart.cau2(a, d);
                case 3:
                    return S2_Chart.cau3(a, d);
                case 4:
                    return S2_Chart.cau4(a, d);
                case 5:
                    return S2_Chart.cau5(a, d);
                case 6:
                    return S2_Chart.cau6(a, d);
                case 7:
                    return S2_Chart.cau7(a, d);
                case 8:
                    return S2_Chart.cau8(a, d);
                case 9:
                    return S2_Chart.cau9(a, d);
                case 10:
                    return S2_Chart.cau10(a, d);
                default:
                    return "case 11";
            }
        }

        private static string cau1(Application a, Presentation d)
        {
            try
            {
                if (a?.ActivePresentation == null)
                    return "False";

                if (a.ActivePresentation.Slides.Count < 7)
                    return "False";

                Slide slide = a.ActivePresentation.Slides[7];
                if (slide.Shapes.Count != 3)
                    return "False";

                Shape shape = slide.Shapes[3];
                if (((object)shape.Type).ToString() != "msoChart")
                    return "False";

                if (!((object)shape.Chart.ChartType).ToString().Contains("Line"))
                    return "False";

                Chart chart = shape.Chart;
                if (chart == null)
                    return "False";

                // ChartData.Workbook is an Excel Workbook COM object
                var rawWorkbook = chart.ChartData?.Workbook;
                Workbook workbook = rawWorkbook as Workbook;
                if (workbook == null)
                    return "False";

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
                    return "False";

                string ReadRangeText(Worksheet ws, string rangeAddress)
                {
                    try
                    {
                        var rng = ws.Range[rangeAddress] as Range;
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

                string a2 = ReadRangeText(worksheet, "A2");
                if (a2 != "2012")
                    return "False";

                string b1 = ReadRangeText(worksheet, "B1");
                if (b1 != "New Customers")
                    return "False";

                string b2 = ReadRangeText(worksheet, "B2");
                if (b2 != "1700000")
                    return "False";

                string b4 = ReadRangeText(worksheet, "B4");
                if (b4 != "3200000")
                    return "False";

                return "True";
            }
            catch (Exception)
            {
                return "False";
            }
        }

        private static string cau2(Application a, Presentation d)
        {
            try
            {
                if (a?.ActivePresentation == null)
                    return "False";

                if (a.ActivePresentation.Slides.Count < 4)
                    return "False";

                Slide slide = a.ActivePresentation.Slides[4];

                // find shape by name safely
                Shape target = null;
                foreach (Shape s in slide.Shapes)
                {
                    if (string.Equals(s.Name, "Content Placeholder 6", StringComparison.OrdinalIgnoreCase))
                    {
                        target = s;
                        break;
                    }
                }

                if (target == null)
                    return "False";

                if (((object)target.Type).ToString() != "msoPlaceholder")
                    return "False";

                Chart chart = null;
                try
                {
                    chart = target.Chart;
                }
                catch
                {
                    chart = null;
                }

                if (chart == null)
                    return "False";

                // If legend is absent, it's a failure for this test
                var legend = chart.Legend;
                if (legend == null)
                    return "False";

                // Compare to enum for robustness
                try
                {
                    // chart.Legend.Position is an XlLegendPosition enum
                    var posObj = legend.Position;
                    if (posObj is XlLegendPosition pos)
                    {
                        return pos == XlLegendPosition.xlLegendPositionTop ? "True" : "False";
                    }
                    // fallback to string compare
                    return posObj.ToString() == "xlLegendPositionTop" ? "True" : "False";
                }
                catch
                {
                    return "False";
                }
            }
            catch (Exception)
            {
                return "False";
            }
        }

        private static string cau3(Application a, Presentation d)
        {
            try
            {
                Slide slide = a.ActivePresentation.Slides[(object)6];
                Chart chart = slide.Shapes[(object)"Chart 4"].Chart;
                return ((object)chart.ChartType).ToString() != "xl3DColumnClustered" ? "False" : "True";
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
                Slide slide = a.ActivePresentation.Slides[(object)4];
                Chart chart = slide.Shapes[(object)4].Chart;
                return chart.ChartType != XlChartType.xl3DColumnClustered ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau5(Application a, Presentation d)
        {
            try
            {
                Slide slide = a.ActivePresentation.Slides[(object)4];
                Chart chart = slide.Shapes[(object)3].Chart;
                return chart.ChartType != XlChartType.xlBarClustered ? "False" : "True";
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
                Slide slide = a.ActivePresentation.Slides[(object)5];
                Chart chart = slide.Shapes[(object)2].Chart;
                return !chart.DataTable.ShowLegendKey ? "False" : "True";
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
                Slide slide = a.ActivePresentation.Slides[(object)6];
                Chart chart = slide.Shapes[(object)3].Chart;
                return chart.ChartType != XlChartType.xlLineMarkers ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau8(Application a, Presentation d) => "True";

        private static string cau9(Application a, Presentation d) => "True";

        private static string cau10(Application a, Presentation d) => "True";
    }
}