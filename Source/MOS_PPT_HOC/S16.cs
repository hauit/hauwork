// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S16
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;


namespace MOS_PPT_LEARN
{
    internal class S16
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S16.Cau1(a, d);
                case 2:
                    return S16.Cau2(a, d);
                case 3:
                    return S16.Cau3(a, d);
                case 4:
                    return S16.Cau4(a, d);
                case 5:
                    return S16.Cau5(a, d);
                case 6:
                    return S16.Cau6(a, d);
                case 7:
                    return S16.Cau7(a, d);
                default:
                    return "case out index";
            }
        }

        private static string Cau1(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 3)
                    return "False";
                if (((object)a.ActivePresentation.Slides[(object)3].Shapes[(object)3].Type).ToString() != "msoTable")
                    return "False";
                int count = a.ActivePresentation.Slides[(object)3].Shapes[(object)3].Table.Columns.Count;
                if (count.ToString() != "3")
                    return "False";
                count = a.ActivePresentation.Slides[(object)3].Shapes[(object)3].Table.Rows.Count;
                if (count.ToString() != "4")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private static string Cau2(Application a, Presentation d)
        {
            try
            {
                // Get slide 8
                var slide = a.ActivePresentation.Slides.Count >= 8
                    ? a.ActivePresentation.Slides[8]
                    : null;
                if (slide == null)
                    return "False";

                // Get shape "Chart 5"
                Shape chartShape = null;
                foreach (Shape s in slide.Shapes)
                {
                    if (s.Name == "Chart 5")
                    {
                        chartShape = s;
                        break;
                    }
                }
                if (chartShape == null)
                    return "False";

                // Check if shape has a Chart
                if (!chartShape.HasChart.Equals(Microsoft.Office.Core.MsoTriState.msoTrue))
                    return "False";

                var chart = chartShape.Chart;
                if (chart == null)
                    return "False";

                // Check ChartStyle
                if (chart.ChartStyle.ToString() != "261")
                    return "False";

                // Check ChartColor
                if (chart.ChartColor.ToString() != "13")
                    return "False";

                return "True";
            }
            catch (Exception)
            {
                return "False";
            }
        }

        private static string Cau3(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.PrintOptions.NumberOfCopies.ToString() != "4")
                    return "False";
                if (a.ActivePresentation.PrintOptions.OutputType.ToString() != "ppPrintOutputThreeSlideHandouts")
                    return "False";
                return ((object)a.ActivePresentation.PrintOptions.Collate).ToString() != "msoFalse" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string Cau4(Application a, Presentation d)
        {
            try
            {
                if (((object)a.ActivePresentation.Slides[(object)6].BackgroundStyle).ToString() != "msoBackgroundStyleNotAPreset")
                    return "False";
                if (a.ActivePresentation.Slides[(object)6].Background.Fill.GradientAngle.ToString() != "90")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private static string Cau5(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)7].Shapes[(object)"Content Placeholder 4"].Shadow.Blur.ToString() != "20")
                    return "False(ShapeStyle)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string Cau6(Application a, Presentation d) => "True";

        private static string Cau7(Application a, Presentation d) => "True";
    }
}