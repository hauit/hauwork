using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Hyperlink = Microsoft.Office.Interop.PowerPoint.Hyperlink;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.IO;
using System.Reflection;
using Binder = System.Reflection.Binder;
using System.Runtime.CompilerServices;
using mscore = Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapePPT = Microsoft.Office.Interop.PowerPoint.Shape;
using XlLegendPosition = Microsoft.Office.Interop.PowerPoint.XlLegendPosition;
using Chart = Microsoft.Office.Interop.PowerPoint.Chart;
using XlChartType = Microsoft.Office.Core.XlChartType;
using Presentation = Microsoft.Office.Interop.PowerPoint.Presentation;

namespace MOS_PPT_TEST
{
    public interface IQuestionCheck
    {
        string CheckAnswer(int questionNumber, Application a, Presentation d);
    }
    internal class QuestionCheck : IQuestionCheck
    {
        public string CheckAnswer(int questionNumber, Application a, Presentation d)
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
                case 126: return Cau126(a, d);
                case 127: return Cau127(a, d);
                default:
                    return "Invalid question number";
            }
        }

        private string Cau1(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count != 3)
                    return "False";
                string name = a.ActivePresentation.Slides[(object)2].Shapes[(object)2].SmartArt.Layout.Name;
                if (name != "Vertical Curved List")
                    return "False";
                if (a.ActivePresentation.Slides[(object)2].Shapes[(object)2].SmartArt.Nodes.Count != 2)
                    return "False";
                string text1 = a.ActivePresentation.Slides[(object)2].Shapes[(object)2].SmartArt.Nodes[(object)1].TextFrame2.TextRange.Text;
                if (text1 != "Structures")
                    return "False";
                string text2 = a.ActivePresentation.Slides[(object)2].Shapes[(object)2].SmartArt.Nodes[(object)2].TextFrame2.TextRange.Text;
                return text2 != "Perennials" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau2(Application a, Presentation d)
        {
            try
            {
                //var test = ((object)a.ActivePresentation.Slides[(object)3].Shapes[(object)2].Type).ToString() + " - " + a.ActivePresentation.Slides[(object)3].Shapes[(object)2].Height.ToString() + "\n";

                //return test;

                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 3)
                    return "False";
                if (((object)a.ActivePresentation.Slides[(object)3].Shapes[(object)2].Type).ToString() != "msoPlaceholder")
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)2].Height.ToString() != "288")
                    return "False";
                //return a.ActivePresentation.Slides[(object)3].Shapes[(object)2].Width.ToString() != "134.255" ? "False" : "True";

                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau3(Application a, Presentation d)
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
        private string Cau4(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].DisplayName != "Picture 3")
                    return "False";
                return a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectFadedSwivel" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau5(Application a, Presentation d)
        {
            try
            {
                foreach (Slide slide in a.ActivePresentation.Slides)
                {
                    if (slide.SlideShowTransition.EntryEffect.ToString() != "ppEffectWipeRight")
                        return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau6(Application a, Presentation d)
        {
            try
            {
                int sectionIndex1 = -1;
                for (int sectionIndex2 = 1; sectionIndex2 <= a.ActivePresentation.SectionProperties.Count; ++sectionIndex2)
                {
                    // ISSUE: reference to a compiler-generated method
                    if (a.ActivePresentation.SectionProperties.Name(sectionIndex2) == "Structures")
                        sectionIndex1 = sectionIndex2;
                }
                if (sectionIndex1 == -1)
                    return "False";
                // ISSUE: reference to a compiler-generated method
                int num = a.ActivePresentation.SectionProperties.FirstSlide(sectionIndex1);
                if (num.ToString() != "3")
                    return "False";
                // ISSUE: reference to a compiler-generated method
                num = a.ActivePresentation.SectionProperties.SlidesCount(sectionIndex1);
                return num.ToString() != "2" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau7(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count != 3)
                    return "False";
                return a.ActivePresentation.Slides[(object)2].Shapes[(object)"Content Placeholder 5"].AlternativeText != "ocean" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau8(Application a, Presentation d)
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
        private string Cau9(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 2)
                    return "False";
                return a.ActivePresentation.Slides[(object)4].Shapes[(object)"Text Placeholder 2"].TextFrame.TextRange.Text != "Admin\rFlowers\r…" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau10(Application a, Presentation d)
        {
            try
            {
                string categories = a.ActivePresentation.BuiltInDocumentProperties["Category"].Value.ToString();

                if (categories.Trim() != "Travel")
                    return "False";
            }
            catch
            {
                return "False";
            }
            return "True";
        }
        private string Cau11(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Table 5"].Table.Rows.Count.ToString() != "7")
                    return "False";
                return a.ActivePresentation.Slides[(object)3].Shapes[(object)"Table 5"].Table.Rows[7].Cells[1].Shape.TextFrame.TextRange.Text != "Sporting Event" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau12(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)1].Shapes.Count != 5)
                    return "False";
                return (double)a.ActivePresentation.Slides[(object)1].Shapes[(object)"Picture 5"].PictureFormat.CropRight < 100.0 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau13(Application a, Presentation d)
        {
            try
            {
                var slide = a.ActivePresentation.Slides[(object)5];  // slide 5

                var test = "";

                foreach (ShapePPT shape in slide.Shapes)
                {
                    if (shape.Name.Contains("Content Placeholder 9"))
                    {
                        //3D Model giờ là msoPlaceholder -> Dùng Name để xác định
                        test = test + shape.Name + "|" + shape.Type + "|" + "\n";

                        try
                        {
                            dynamic model = shape.GetType().InvokeMember("Model3D",
                                System.Reflection.BindingFlags.GetProperty, null, shape, null);

                            if (model != null)
                            {
                                float rotX = model.RotationX;
                                float rotY = model.RotationY;

                                test += rotX + "+" + rotY + "-" + (shape.Height - 324f);

                                //Above Front Right trên máy mới = RotationX ≈ -30, RotationY ≈ 45
                                if (Math.Abs(rotX) == 20 && Math.Abs(rotY) == 330 &&
                                    Math.Abs(shape.Height - 324f) == 0)   // 4.5 inch = 324 pt
                                    return "True";
                            }
                        }
                        catch { }

                        // Fallback máy cũ: dùng ThreeD
                        float rx = shape.ThreeD.RotationX;
                        float ry = shape.ThreeD.RotationY;

                        test += rx + "+" + ry + "-" + (shape.Height - 324f);

                        // Above Front Right trên máy cũ cũng là -30 / 45
                        if (Math.Abs(rx) == 20 && Math.Abs(ry) == 330 &&
                            Math.Abs(shape.Height - 324f) == 0)
                            return "True";
                    }
                }
                //return test + " 01";
                return "False";
            }
            catch (Exception ex)
            {
                return "False 3" + ex.Message;
            }
        }
        private string Cau14(Application a, Presentation d)
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
        private string Cau15(Application a, Presentation d)
        {
            try
            {
                foreach (Slide slide in a.ActivePresentation.Slides)
                {
                    if (slide.SlideShowTransition.EntryEffect.ToString() != "ppEffectRotateRight")
                        return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau16(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count.ToString() != "4")
                    return "Fales";
                if (a.ActivePresentation.Slides[(object)1].Shapes.Count.ToString() != "5")
                    return "Fales";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau17(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)7].Shapes.Count != 5)
                    return "False";
                if (a.ActivePresentation.Slides[(object)7].Shapes[(object)"Cloud 8"].Width.ToString() != "417.6")
                    return "False";
                return a.ActivePresentation.Slides[(object)7].Shapes[(object)"Cloud 8"].Height.ToString() != "144" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau18(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 2)
                    return "False";
                return a.ActivePresentation.Slides[(object)3].Shapes[(object)"Text Placeholder 2"].TextFrame.TextRange.Text != "Home stay\rHouse boat\r…" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau19(Application a, Presentation d)
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
        private string Cau20(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)7].TimeLine.MainSequence.Count != 1)
                    return "False";
                if (a.ActivePresentation.Slides[(object)7].TimeLine.MainSequence[1].DisplayName != "Picture 4")
                    return "False";
                if (a.ActivePresentation.Slides[(object)7].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectFly")
                    return "False";
                if (a.ActivePresentation.Slides[(object)7].TimeLine.MainSequence[1].EffectParameters.Direction.ToString() != "msoAnimDirectionUpLeft")
                    return "False";
                return a.ActivePresentation.Slides[(object)7].TimeLine.MainSequence[1].Timing.Duration.ToString() != "2" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau21(Application a, Presentation d)
        {
            try
            {
                object documentProperties = a.ActivePresentation.BuiltInDocumentProperties;
                object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (Binder)null, documentProperties, new object[1]
                {
(object) "Title"
                });
                string str = target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (Binder)null, target, new object[0]).ToString();
                if (str != "Preferred Customer Program")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau22(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes[(object)"Table 5"].Table.Columns.Count.ToString() != "3")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau23(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count != 5)
                    return "False";
                if (a.ActivePresentation.Slides[(object)2].Shapes[(object)1].Name != "Picture 3")
                    return "False";
                return a.ActivePresentation.Slides[(object)2].Shapes[(object)5].Name != "Picture 6" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau24(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 2)
                    return "False";
                return a.ActivePresentation.Slides[(object)3].Shapes[(object)2].SmartArt.Layout.Name != "Basic Block List" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau25(Application a, Presentation d)
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
        private string Cau26(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)5].SlideShowTransition.Hidden != MsoTriState.msoTrue ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau27(Application a, Presentation d)
        {
            try
            {
                foreach (Slide slide in a.ActivePresentation.Slides)
                {
                    if (slide.SlideShowTransition.Duration.ToString() != "3")
                        return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau28(Application a, Presentation d)
        {
            try
            {
                var slide1 = a.ActivePresentation.Slides[1];

                // PowerPoint lưu hyperlink trong slide.Hyperlinks (danh sách toàn bộ hyperlink trên slide)
                foreach (Hyperlink link in slide1.Hyperlinks)
                {
                    // link.TextToDisplay = text hiển thị (ví dụ: "Trey Research")
                    // link.Address = URL (ví dụ: "http://www.treyresearch.net")
                    string displayText = link.TextToDisplay.Trim();
                    string url = link.Address.Trim();

                    //return displayText.ToString() + " " + url.ToString();

                    if (displayText == "Trey Research" && url.Contains("http://www.treyresearch.net"))
                    {
                        return "True";
                    }
                }
                return "False";
            }
            catch
            {
                return "False";
            }
        }
        private string Cau29(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)6].Shapes.Count != 8)
                    return "False";
                string alternativeText = a.ActivePresentation.Slides[(object)6].Shapes[(object)"Soccer"].AlternativeText;
                return alternativeText != "Soccer players" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau30(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count != 2)
                    return "False";
                string name = a.ActivePresentation.Slides[(object)2].Shapes[(object)2].SmartArt.Layout.Name;
                return name != "Vertical Box List" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau31(Application a, Presentation d)
        {
            try
            {
                //return a.ActivePresentation.Slides[(object)5].Shapes.Count != 6 ? "False" : "True";

                //var test = ((object)a.ActivePresentation.Slides[(object)5].Shapes[(object)6].Type).ToString() + " - " + a.ActivePresentation.Slides[(object)5].Shapes[(object)6].Height.ToString() + " - " + a.ActivePresentation.Slides[(object)5].Shapes[(object)6].Width.ToString() + "\n";

                //return test;

                if (a.ActivePresentation.Slides[(object)5].Shapes.Count != 6)
                    return "False";
                // if (((object)a.ActivePresentation.Slides[(object)5].Shapes[(object)6].Type).ToString() != "msoPlaceholder")
                if (((object)a.ActivePresentation.Slides[(object)5].Shapes[(object)6].Type).ToString() != "30")
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].Shapes[(object)6].Height.ToString() != "144")
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].Shapes[(object)6].Width.ToString() != "128.16")
                    return "False";

                float top = a.ActivePresentation.Slides[(object)5].Shapes[(object)6].Top;
                float left = a.ActivePresentation.Slides[(object)5].Shapes[(object)6].Left;

                if (top > 199 && left > 103)
                    return "True";

                //return a.ActivePresentation.Slides[(object)5].Shapes[(object)6].Width.ToString() != "128.16" ? "False" : "True";
                return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau32(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence.Count != 4)
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].EffectParameters.Direction.ToString() != "msoAnimDirectionLeft")
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].Timing.Duration.ToString() != "1.5")
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[4].EffectParameters.Direction.ToString() != "msoAnimDirectionLeft")
                    return "False";
                return a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[4].Timing.Duration.ToString() != "1.5" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau33(Application a, Presentation d)
        {
            try
            {
                foreach (Slide slide in a.ActivePresentation.Slides)
                {
                    if (slide.SlideShowTransition.Duration.ToString() != "2")
                        return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau34(Application a, Presentation d)
        {
            try
            {
                int sectionIndex1 = -1;
                for (int sectionIndex2 = 1; sectionIndex2 <= a.ActivePresentation.SectionProperties.Count; ++sectionIndex2)
                {
                    // ISSUE: reference to a compiler-generated method
                    if (a.ActivePresentation.SectionProperties.Name(sectionIndex2) == "Student Clubs")
                        sectionIndex1 = sectionIndex2;
                }
                if (sectionIndex1 == -1)
                    return "False";
                // ISSUE: reference to a compiler-generated method
                int num = a.ActivePresentation.SectionProperties.FirstSlide(sectionIndex1);
                if (num.ToString() != "3")
                    return "False";
                // ISSUE: reference to a compiler-generated method
                num = a.ActivePresentation.SectionProperties.SlidesCount(sectionIndex1);
                return num.ToString() != "5" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau35(Application a, Presentation d)
        {
            try
            {
                string text1 = a.ActivePresentation.HandoutMaster.Shapes[(object)"Header Placeholder 1"].TextFrame.TextRange.Text;
                if (text1 != "First Up Consultants")
                    return "False";
                string text2 = a.ActivePresentation.HandoutMaster.Shapes[(object)"Footer Placeholder 3"].TextFrame.TextRange.Text;
                if (text2 != "www.firstupconsultants.com")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau36(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)5].Shapes.Count != 2)
                    return "False";
                return a.ActivePresentation.Slides[(object)5].Shapes[(object)2].SmartArt.Layout.Name != "Segmented Cycle" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau37(Application a, Presentation d)
        {
            try
            {
                foreach (Slide slide in a.ActivePresentation.Slides)
                {
                    if (slide.SlideShowTransition.EntryEffect.ToString() != "ppEffectWedge")
                        return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau38(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count != 2)
                    return "False";
                if (!a.ActivePresentation.Slides[(object)2].Shapes[(object)2].Name.Contains("ummary"))
                    return "Fales";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau39(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)6].Background.Fill.Transparency.ToString() != "0.75")
                    return "False";
                if (a.ActivePresentation.Slides[(object)6].Background.Fill.Type != MsoFillType.msoFillPicture)
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].Background.Fill.Transparency.ToString() == "0.75")
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].Background.Fill.Type == MsoFillType.msoFillPicture)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau40(Application a, Presentation d)
        {
            try
            {
                var slide1 = a.ActivePresentation.Slides[1];

                // PowerPoint lưu hyperlink trong slide.Hyperlinks (danh sách toàn bộ hyperlink trên slide)
                foreach (Hyperlink link in slide1.Hyperlinks)
                {
                    // link.TextToDisplay = text hiển thị (ví dụ: "Contact Us")
                    // link.Address = URL (ví dụ: "http://www.firstupconsultants.com")
                    string displayText = link.TextToDisplay.Trim();
                    string url = link.Address.Trim();

                    //return displayText.ToString() + " " + url.ToString();

                    if (displayText == "Contact Us" && url.Contains("http://www.firstupconsultants.com"))
                    {
                        return "True";
                    }
                }
                return "False";
            }
            catch
            {
                return "False";
            }
        }
        private string Cau41(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.SlideMaster.CustomLayouts.Count < 12)
                    return "False";
                for (int Index = 1; Index <= a.ActivePresentation.SlideMaster.CustomLayouts.Count; ++Index)
                {
                    if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Name == "Picture with Text")
                    {
                        if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Shapes.Count.ToString() != "5")
                            return "False";
                        string str = a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Shapes[(object)5].Name + a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Shapes[(object)4].Name;
                        if (str.Contains("Text") && str.Contains("Picture"))
                            return "True";
                    }
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "False";
        }
        private string Cau42(Application a, Presentation d)
        {
            try
            {
                if (((object)a.ActivePresentation.Slides[(object)2].Shapes[(object)"Picture 5"].Fill.PictureEffects[1].Type).ToString() != "msoEffectPaintBrush")
                    return "False";
                return a.ActivePresentation.Slides[(object)2].Shapes[(object)"Picture 5"].ThreeD.BevelTopDepth.ToString() != "4" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau43(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)6].Shapes.Count != 2)
                    return "False";
                if (a.ActivePresentation.Slides[(object)6].Shapes[(object)1].TextFrame.TextRange.Text != "Discover Your Campus")
                    return "False";
                if (a.ActivePresentation.Slides[(object)7].Shapes.Count != 2)
                    return "False";
                return a.ActivePresentation.Slides[(object)7].Shapes[(object)1].TextFrame.TextRange.Text != "Display Your Art" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau44(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)5].Shapes.Count != 3)
                    return "False";
                return a.ActivePresentation.Slides[(object)5].Shapes[(object)3].Name != "Sailing" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau45(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence.Count != 1)
                    return "False";
                return a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectPathDown" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau46(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.PrintOptions.NumberOfCopies != 3)
                    return "False";
                if (a.ActivePresentation.PrintOptions.OutputType != PpPrintOutputType.ppPrintOutputNotesPages)
                    return "False";
                if (a.ActivePresentation.PrintOptions.Collate != MsoTriState.msoFalse)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau47(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.SlideMaster.CustomLayouts.Count != 12)
                    return "False";
                for (int Index = 1; Index <= a.ActivePresentation.SlideMaster.CustomLayouts.Count; ++Index)
                {
                    if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Name == "Ingredients")
                    {
                        if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Shapes.Count.ToString() != "5")
                            return "False";
                        // ISSUE: reference to a compiler-generated method
                        return a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Shapes[(object)"Content Placeholder 6"].TextFrame.TextRange.Paragraphs(1, 1).ParagraphFormat.Bullet.Type.ToString() == "ppBulletPicture" ? "True" : "False";
                    }
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "False";
        }
        private string Cau48(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 8)
                    return "False";
                if (((object)a.ActivePresentation.Slides[(object)4].Shapes[(object)"Freeform 6"].Shadow.Style).ToString() != "msoShadowStyleInnerShadow")
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)"Freeform 6"].Shadow.OffsetX.ToString() != "-2.12132" || a.ActivePresentation.Slides[(object)4].Shapes[(object)"Freeform 6"].Shadow.OffsetY.ToString() != "-2.12132")
                    return "False";
                if (((object)a.ActivePresentation.Slides[(object)4].Shapes[(object)"Freeform 5"].Shadow.Style).ToString() != "msoShadowStyleInnerShadow")
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)"Freeform 5"].Shadow.OffsetX.ToString() != "-2.12132")
                    return "False";
                return a.ActivePresentation.Slides[(object)4].Shapes[(object)"Freeform 5"].Shadow.OffsetY.ToString() != "-2.12132" ? "False 6" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau49(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)5].Shapes.Count != 3)
                    return "False";
                return a.ActivePresentation.Slides[(object)5].Shapes[(object)2].Name != "Rectangle 5" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau50(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)1].Shapes[(object)"Sleep Away"].MediaFormat.FadeInDuration != 2000)
                    return "False";
                return a.ActivePresentation.Slides[(object)1].Shapes[(object)"Sleep Away"].AnimationSettings.PlaySettings.StopAfterSlides != 999 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau51(Application a, Presentation d)
        {
            try
            {
                object documentProperties = a.ActivePresentation.BuiltInDocumentProperties;
                object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (Binder)null, documentProperties, new object[1]
                {
(object) "Title"
                });
                if (target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (Binder)null, target, new object[0]).ToString() != "")
                    return "False";
                if (((object)a.ActivePresentation.RemovePersonalInformation).ToString() != "msoTrue")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau52(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.SlideShowSettings.AdvanceMode.ToString() != "ppSlideShowManualAdvance")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau53(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count != 5)
                    return "False";
                float top1 = a.ActivePresentation.Slides[(object)2].Shapes[(object)"Content Placeholder 5"].Top;
                float top2 = a.ActivePresentation.Slides[(object)2].Shapes[(object)"Content Placeholder 3"].Top;
                float left1 = a.ActivePresentation.Slides[(object)2].Shapes[(object)"Content Placeholder 5"].Left;
                float left2 = a.ActivePresentation.Slides[(object)2].Shapes[(object)"Content Placeholder 3"].Left;
                if ((double)top1 != (double)top2)
                    return "False";
                return left2.ToString() != "77.16708" || (double)left1 != 486.0 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau54(Application a, Presentation d)
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
        private string Cau55(Application a, Presentation d)
        {
            try
            {
                //var test = a.ActivePresentation.Slides[(object)6].Shapes[(object)3].Name + " - " + a.ActivePresentation.Slides[(object)6].Shapes[(object)3].Top.ToString() + " - " + a.ActivePresentation.Slides[(object)6].Shapes[(object)3].Left.ToString() + "\n";

                //return test;

                if (a.ActivePresentation.Slides[(object)6].Shapes.Count != 3)
                    return "False";

                if (a.ActivePresentation.Slides[(object)6].Shapes[(object)3].Name != "River")
                    return "False";

                float left = a.ActivePresentation.Slides[(object)6].Shapes[(object)3].Left;

                if (left > 426)
                    return "True";

                //return a.ActivePresentation.Slides[(object)6].Shapes[(object)3].Name != "River" ? "False" : "True";

                return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau56(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence.Count != 1)
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].DisplayName != "5-Point Star 5")
                    return "False";
                return a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectPathHeart" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau57(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.PrintOptions.NumberOfCopies != 5)
                    return "False";
                if (a.ActivePresentation.PrintOptions.OutputType != PpPrintOutputType.ppPrintOutputNotesPages)
                    return "False";
                if (a.ActivePresentation.PrintOptions.Collate != MsoTriState.msoFalse)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau58(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)1].Shapes.Count != 3)
                    return "False";
                return !a.ActivePresentation.Slides[(object)1].Shapes[(object)3].Name.Contains("oom") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau59(Application a, Presentation d)
        {
            try
            {
                var slide = a.ActivePresentation.Slides[(object)3];  // slide 3
                if (slide.Shapes.Count != 4) return "False";

                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[(object)i];

                    if ((int)shape.Type == 29 || shape.Name.Contains("3D Model") || shape.Tags["OBJECTTYPE"] == "3DMODEL")
                    {
                        // Tìm thấy 3D Model rồi đây

                        dynamic model = shape.GetType().InvokeMember("Model3D",
                                    System.Reflection.BindingFlags.GetProperty,
                                    null, shape, null);

                        // Nếu có Model3D (máy mới)
                        if (model != null)
                        {
                            float rotY = model.RotationY;
                            float rotX = model.RotationX;

                            // Máy mới 2023–2025: Left = 90
                            if (Math.Abs(rotY - 90f) < 10f && Math.Abs(rotX) < 10f)
                                return "True";
                        }
                        else
                        {
                            // Máy cũ: fallback về ThreeD
                            float rotY = shape.ThreeD.RotationY;
                            if (Math.Abs(rotY - 270f) < 10f && Math.Abs(shape.ThreeD.RotationX) < 10f)
                                return "True";
                        }
                    }
                }

                return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau60(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].DisplayName != "3D Model 3")
                    return "False";
                return a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectType.ToString() != "154" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau61(Application a, Presentation d)
        {
            try
            {
                foreach (Slide slide in a.ActivePresentation.Slides)
                {
                    if (slide.SlideShowTransition.EntryEffect.ToString() != "ppEffectWipeRight")
                        return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau62(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count <= 1)
                    return "False";
                return !a.ActivePresentation.Slides[(object)2].Shapes[(object)2].Name.Contains("Section Zoom") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau63(Application a, Presentation d)
        {
            try
            {
                var pres = a.ActivePresentation;
                var master = pres.SlideMaster;

                // 1. Kiểm tra Footer trên Slide Master
                var footer = master.HeadersFooters.Footer;
                bool footerTextCorrect = footer.Text.Trim().Equals("www.adventure-works.com",
                                                                  StringComparison.OrdinalIgnoreCase);

                bool slideNumberVisible = master.HeadersFooters.SlideNumber.Visible == MsoTriState.msoTrue;
                bool footerVisible = footer.Visible == MsoTriState.msoTrue;

                // 2. Kiểm tra Title Slide (slide 1) KHÔNG có footer và số slide
                var titleSlide = pres.Slides[1];
                bool titleSlideNoFooter = titleSlide.HeadersFooters.Footer.Visible == MsoTriState.msoFalse;
                bool titleSlideNoSlideNumber = titleSlide.HeadersFooters.SlideNumber.Visible == MsoTriState.msoFalse;

                // 3. Kiểm tra 1 slide bất kỳ (ví dụ slide 2) phải CÓ footer + số slide
                var slide2 = pres.Slides[2];
                bool slide2HasFooter = slide2.HeadersFooters.Footer.Visible == MsoTriState.msoTrue &&
                                            slide2.HeadersFooters.Footer.Text.Contains("adventure-works.com");
                bool slide2HasSlideNumber = slide2.HeadersFooters.SlideNumber.Visible == MsoTriState.msoTrue;

                return (footerTextCorrect &&
                        slideNumberVisible &&
                        footerVisible &&
                        titleSlideNoFooter &&
                        titleSlideNoSlideNumber &&
                        slide2HasFooter &&
                        slide2HasSlideNumber) ? "True" : "False";
            }
            catch
            {
                return "False";
            }
        }
        private string Cau64(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)8].Shapes[(object)2].TextFrame2.Column.Number != 2)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau65(Application a, Presentation d)
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
        private string Cau66(Application a, Presentation d)
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
        private string Cau67(Application a, Presentation d)
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
        private string Cau68(Application a, Presentation d)
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
        private string Cau69(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)8].Shapes[(object)1].TextFrame.TextRange.Text != "Title Layout")
                    return "False";
                if (a.ActivePresentation.Slides[(object)12].Shapes[(object)1].TextFrame.TextRange.Text != "Two Content Layout with SmartArt")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau70(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.TemplateName != "Ion Boardroom")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau71(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Diagram 2"].SmartArt.Color.Name != "Colorful Range -Accent Colors 4 to 5")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau72(Application a, Presentation d)
        {
            string str = "True";
            try
            {
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence.Count != 4)
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[1].DisplayName != "Restaurant ")
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectWipe")
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[1].EffectParameters.Direction != MsoAnimDirection.msoAnimDirectionUp)
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[1].Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[1].Timing.TriggerDelayTime.ToString() != "0")
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[2].EffectParameters.Direction != MsoAnimDirection.msoAnimDirectionUp)
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[2].Timing.TriggerDelayTime.ToString() != "1")
                    return "False";
                str = a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[2].Timing.TriggerType.ToString();
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[2].Timing.TriggerType != MsoAnimTriggerType.msoAnimTriggerAfterPrevious)
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[2].EffectType.ToString() != "msoAnimEffectWipe")
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[3].EffectParameters.Direction != MsoAnimDirection.msoAnimDirectionUp)
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[3].Timing.TriggerDelayTime.ToString() != "1")
                    return "False";
                str = a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[3].Timing.TriggerType.ToString();
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[3].Timing.TriggerType != MsoAnimTriggerType.msoAnimTriggerAfterPrevious)
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[3].EffectType.ToString() != "msoAnimEffectWipe")
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[3].EffectParameters.Direction != MsoAnimDirection.msoAnimDirectionUp)
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[3].Timing.TriggerDelayTime.ToString() != "1")
                    return "False";
                str = a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[3].Timing.TriggerType.ToString();
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[3].Timing.TriggerType != MsoAnimTriggerType.msoAnimTriggerAfterPrevious)
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[3].EffectType.ToString() != "msoAnimEffectWipe")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau73(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)7].Shapes.Count != 2)
                    return "False";
                if (((object)a.ActivePresentation.Slides[(object)7].Shapes[(object)2].Type).ToString() != "msoChart")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau74(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 2)
                    return "False";
                return ((object)a.ActivePresentation.Slides[(object)4].Shapes[(object)2].Type).ToString() != "msoEmbeddedOLEObject" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau75(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count != 2)
                    return "False";
                if (a.ActivePresentation.Slides[(object)2].Shapes[(object)"TextBox 4"].Fill.BackColor.RGB.ToString() != "14145397")
                    return "False";
                if (a.ActivePresentation.Slides[(object)2].Shapes[(object)"TextBox 4"].Line.Weight.ToString() != "3")
                    return "False";
                return a.ActivePresentation.Slides[(object)2].Shapes[(object)"TextBox 4"].ThreeD.BevelTopDepth.ToString() != "6" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau76(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides.Count != 10)
                    return "False";
                if (!a.ActivePresentation.Slides[(object)10].Shapes[(object)"Title 1"].TextFrame.TextRange.Text.Contains("New product"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau77(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.SlideMaster.CustomLayouts.Count < 12)
                    return "False";
                for (int Index = 1; Index <= a.ActivePresentation.SlideMaster.CustomLayouts.Count; ++Index)
                {
                    if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Name == "Trevorslayout")
                    {
                        if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Shapes.Count.ToString() != "6")
                            return "False";
                        if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Shapes[(object)5].PlaceholderFormat.Type.ToString() != "ppPlaceholderPicture")
                            return "False";
                        return a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Shapes[(object)6].PlaceholderFormat.Type.ToString() != "ppPlaceholderBody" ? "False" : "True";
                    }
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "False";
        }
        private string Cau78(Application a, Presentation d)
        {
            try
            {
                string str = "True";
                if (a.ActivePresentation.HandoutMaster.Shapes.Count != 4)
                    return "False";
                try
                {
                    // ISSUE: variable of a compiler-generated type
                    Microsoft.Office.Interop.PowerPoint.Shape shape = a.ActivePresentation.HandoutMaster.Shapes[(object)3];
                    if (shape.TextFrame.TextRange.Text != "First Copy")
                        return "False";
                }
                catch (Exception ex)
                {
                    return "False";
                }
                return str;
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau79(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)3].Shapes.Count.ToString() != "5")
                    return "False";
                if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)3].Shapes[(object)5].PlaceholderFormat.Type.ToString() != "ppPlaceholderMediaClip")
                    return "False";
                if ((double)a.ActivePresentation.SlideMaster.CustomLayouts[(object)3].Shapes[(object)5].Left != (double)a.ActivePresentation.SlideMaster.CustomLayouts[(object)3].Shapes[(object)1].Left)
                    return "False";
                if ((double)a.ActivePresentation.SlideMaster.CustomLayouts[(object)3].Shapes[(object)5].Width != (double)a.ActivePresentation.SlideMaster.CustomLayouts[(object)3].Shapes[(object)1].Width)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau80(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)5].Shapes[(object)"Picture 4"].Fill.PictureEffects.Count.ToString() != "1")
                    return "False";
                return ((object)a.ActivePresentation.Slides[(object)5].Shapes[(object)"Picture 4"].Fill.PictureEffects[1].Type).ToString() != "msoEffectPastelsSmooth" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau81(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 2)
                    return "False";
                return !a.ActivePresentation.Slides[(object)3].Shapes[(object)1].TextFrame.TextRange.Text.Contains("New product") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau82(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 2)
                    return "False";
                return !a.ActivePresentation.Slides[(object)4].Shapes[(object)1].TextFrame.TextRange.Text.Contains("New product") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau83(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)2].Shapes[(object)5].Top.ToString() != "267.12" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau84(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.PrintOptions.sectionIndex != 2)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau85(Application a, Presentation d)
        {
            try
            {
                var pres = a.ActivePresentation;
                var ps = pres.PageSetup;

                // 1. Kiểm tra kích thước chính xác đến 2 chữ số thập phân
                bool correctWidth = Math.Abs(ps.SlideWidth - 576f) < 0.5f;   // 8 inches  = 576 pt
                bool correctHeight = Math.Abs(ps.SlideHeight - 792f) < 0.5f;   // 11 inches = 792 pt

                // 2. Kiểm tra đã chọn "Ensure Fit" (rất quan trọng – MOS bắt buộc)
                // Khi chọn Ensure Fit → PowerPoint tự động set SlideOrientation = msoPortrait
                // và Scale để nội dung vừa khung → nhưng quan trọng nhất là Width & Height phải đúng
                bool ensureFit = correctWidth && correctHeight;

                if (ensureFit)
                    return "True";
                else
                    return "False";
            }
            catch
            {
                return "False";
            }
        }
        private string Cau86(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 11)
                    return "False";
                return a.ActivePresentation.Slides[(object)3].Shapes[(object)11].Name != "Oval 11" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau87(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)1].Shapes.Count != 1)
                    return "False";
                return a.ActivePresentation.Slides[(object)1].Shapes[(object)1].TextFrame.TextRange.Font.Color.RGB.ToString() != "16777215" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau88(Application a, Presentation d)
        {
            try
            {
                var slide2 = a.ActivePresentation.Slides[2];   // vị trí thứ 2

                // Duyệt TẤT CẢ shape trên slide 2
                foreach (Shape sh in slide2.Shapes)
                {
                    // Chỉ quan tâm shape có chữ
                    if (sh.HasTextFrame == MsoTriState.msoTrue &&
                        sh.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        string text = sh.TextFrame.TextRange.Text.Trim();

                        if (text.ToLower().Contains("try our two new flavours!"))
                            return "True";
                    }
                }
                return "False";
            }
            catch
            {
                return "False";
            }
        }
        private string Cau89(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.SlideMaster.CustomLayouts.Count < 12)
                    return "False";
                for (int Index = 1; Index <= a.ActivePresentation.SlideMaster.CustomLayouts.Count; ++Index)
                {
                    if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Name == "Custom1")
                    {
                        if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Shapes.Count.ToString() != "6")
                            return "False";
                        if (a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Shapes[(object)5].PlaceholderFormat.Type.ToString() != "ppPlaceholderPicture")
                            return "False";
                        return a.ActivePresentation.SlideMaster.CustomLayouts[(object)Index].Shapes[(object)6].PlaceholderFormat.Type.ToString() != "ppPlaceholderBody" ? "False" : "True";
                    }
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "False";
        }
        private string Cau90(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)2].Shapes[(object)"Picture 3"].Top.ToString() != "315.8544" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau91(Application a, Presentation d)
        {
            try
            {
                var slide = a.ActivePresentation.Slides[(object)3];
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[(object)i];

                    // Icon SVG thật sự có Type = msoAutoShape hoặc msoFreeform + có Fill + Line
                    if (shape.Fill.Type == MsoFillType.msoFillSolid &&
                        shape.Line.Visible == MsoTriState.msoTrue)
                    {
                        //return i.ToString();

                        int fill = shape.Fill.ForeColor.RGB & 0xFFFFFF;
                        int line = shape.Line.ForeColor.RGB & 0xFFFFFF;

                        if (fill == 12611584 && line == 65535)
                            return "True";
                    }
                }

                return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau92(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence.Count != 4)
                    return "False";
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[1].DisplayName != "Picture 3")
                    return "False";
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[2].DisplayName != "Picture 4")
                    return "False";
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[3].DisplayName != "Picture 5")
                    return "False";
                return a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[4].DisplayName != "Picture 6" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau93(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence.Count != 1)
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].DisplayName != "Picture 5")
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectType != MsoAnimEffect.msoAnimEffectFly)
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectParameters.Direction != MsoAnimDirection.msoAnimDirectionRight)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau94(Application a, Presentation d)
        {
            try
            {
                if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "Presentation.pdf")))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau95(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.PrintOptions.sectionIndex != 2)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau96(Application a, Presentation d)
        {
            try
            {
                if (((object)a.ActivePresentation.Slides[(object)5].Shapes[(object)2].Type).ToString() != "msoPlaceholder")
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].Shapes[(object)2].SmartArt.Layout.Name != "Pyramid List")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau97(Application a, Presentation d)
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
        private string Cau98(Application a, Presentation d)
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
        private string Cau99(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.PrintOptions.OutputType.ToString() != "ppPrintOutputNotesPages")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau100(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 3)
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)3].Name != "TextBox 4")
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)2].Name != "Picture 6")
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)1].Name != "Picture 3")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau101(Application a, Presentation d)
        {
            try
            {
                var design = a.ActivePresentation.Designs[1];        // Slide Master luôn là Designs[1]

                // 1. Kiểm tra Design Name chính xác là "Office Theme"
                bool correctTheme = design.Name.Equals("Office Theme",
                                         StringComparison.OrdinalIgnoreCase);

                // 2. Kiểm tra font của Title và Body trên Master phải là Arial
                var masterSlide = design.SlideMaster;

                // Title font
                bool titleIsArial = masterSlide.Shapes.Placeholders[1]
                                        .TextFrame.TextRange.Font.Name == "Arial";

                // Body font (thường là placeholder 2)
                bool bodyIsArial = masterSlide.Shapes.Placeholders[2]
                                       .TextFrame.TextRange.Font.Name == "Arial";

                // Nếu không có placeholder 2 (một số theme không có) → check toàn bộ layout đầu tiên
                if (masterSlide.Shapes.Placeholders.Count < 2)
                {
                    bodyIsArial = true; // coi như pass nếu theme không có body placeholder
                }

                if (correctTheme && titleIsArial && bodyIsArial)
                    return "True";
                else
                    return "False";
            }
            catch
            {
                return "False";
            }
        }
        private string Cau102(Application a, Presentation d)
        {
            try
            {
                float bevelTopDepth = a.ActivePresentation.Slides[(object)2].Shapes[(object)"Picture 9"].ThreeD.BevelTopDepth;
                if (bevelTopDepth.ToString() != "6")
                    return "False";
                bevelTopDepth = a.ActivePresentation.Slides[(object)2].Shapes[(object)"Picture 3"].ThreeD.BevelTopDepth;
                return bevelTopDepth.ToString() != "6" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau103(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Title 1"].TextFrame.TextRange.Text != "Extra")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau104(Application a, Presentation d)
        {
            try
            {
                // ISSUE: reference to a compiler-generated method
                if (a.ActivePresentation.SlideMaster.Shapes[(object)"Text Placeholder 2"].TextFrame.TextRange.Lines(1, 1).ParagraphFormat.Bullet.Type.ToString() != "ppBulletPicture")
                    return "False";
                if (a.ActivePresentation.SlideMaster.Shapes[(object)"Text Placeholder 2"].TextFrame.TextRange.ParagraphFormat.Bullet.Type.ToString() == "ppBulletPicture")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau105(Application a, Presentation d)
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
        private string Cau106(Application a, Presentation d)
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
        private string Cau107(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides.Count != 4)
                    return "False";
                if (!a.ActivePresentation.Slides[(object)4].Shapes[(object)1].TextFrame.TextRange.Text.Contains("Certificate"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau108(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes[(object)"New Advert"].MediaFormat.StartPoint.ToString() != "500")
                    return "False";
                if (a.ActivePresentation.Slides[(object)2].Shapes[(object)"New Advert"].MediaFormat.EndPoint.ToString() != "2500")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau109(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence.Count != 5)
                    return "False";
                if (!a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].DisplayName.Contains("Choose by 95 % "))
                    return "False";
                if (!a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[5].DisplayName.Contains("Top 10"))
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectWipe")
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectParameters.Direction.ToString() != "msoAnimDirectionLeft")
                    return "False";
                return a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[5].Timing.TriggerType.ToString() != "msoAnimTriggerOnPageClick" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau110(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)1].Shapes.Count == 1)
                    return "False";
                if (a.ActivePresentation.Slides[(object)1].Shapes.Count > 2)
                    return "False";
                switch (a.ActivePresentation.Slides[(object)1].Shapes[(object)1].Name)
                {
                    case "5 - Point Star 2":
                        return "False";
                    case "Ink 1":
                        return "False";
                    default:
                        switch (a.ActivePresentation.Slides[(object)1].Shapes[(object)2].Name)
                        {
                            case "5 - Point Star 2":
                                return "False";
                            case "Ink 1":
                                return "False";
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau111(Application a, Presentation d)
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
        private string Cau112(Application a, Presentation d)
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
        private string Cau113(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 2)
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)2].Name != "New Advert")
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)2].Top.ToString() != "143.75")
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)2].Left.ToString() != "144")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau114(Application a, Presentation d)
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
        private string Cau115(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence.Count.ToString() != "1")
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].Timing.TriggerType.ToString() != "msoAnimTriggerAfterPrevious")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau116(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Powerpoint mos vid"].PictureFormat.CropLeft.ToString() != "487.268")
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Powerpoint mos vid"].PictureFormat.CropRight.ToString() != "-7.277816")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau117(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence.Count.ToString() != "1")
                    return "False";
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[1].DisplayName != "No Way!")
                    return "False";
                return a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectPathCircle" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau118(Application a, Presentation d)
        {
            try
            {
                foreach (Slide slide in a.ActivePresentation.Slides)
                {
                    if (slide.SlideShowTransition.Duration.ToString() != "3")
                        return "False";
                    if (slide.SlideShowTransition.SoundEffect.Name != "breeze.wav")
                        return "False";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
        private string Cau119(Application a, Presentation d)
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
        private string Cau120(Application a, Presentation d)
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
        private string Cau121(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.SlideShowSettings.NamedSlideShows.Count != 1)
                    return "False";
                if (a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object)1].Name != "Important Findings")
                    return "False";
                return a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object)1].Count != 2 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau122(Application a, Presentation d)
        {
            try
            {
                return new FileInfo(a.ActivePresentation.FullName).Length < 62878L ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau123(Application a, Presentation d)
        {
            try
            {
                // Cách chuẩn nhất MOS chấm điểm 2025
                bool hidden = a.ActivePresentation.DisplayComments == MsoTriState.msoFalse;

                return hidden ? "True" : "False";
            }
            catch
            {
                return "False";
            }
        }
        private string Cau124(Application a, Presentation d)
        {
            try
            {
                var pres = a.ActivePresentation;
                var po = pres.PrintOptions;

                // 1. Copies = 3
                bool copiesOK = po.NumberOfCopies == 3;

                // 2. Collate = msoTrue (Uncollated → False, Collated → True)
                bool collateOK = po.Collate == MsoTriState.msoFalse;

                // 3. FrameSlides = msoTrue (có viền)
                bool frameOK = po.FrameSlides == MsoTriState.msoTrue;

                //// 4. QUAN TRỌNG NHẤT: PowerPoint lưu "3 Slides per page" trong PrintRanges
                //// Khi chọn "Handouts (3 slides per page)" → tự động tạo 1 PrintRange cho toàn bộ slide

                //bool hasFullRange = po.Ranges.Count == 1 &&
                //            po.Ranges[1].Start == 1 &&
                //            po.Ranges[1].End == pres.Slides.Count;

                if (copiesOK && collateOK && frameOK)
                    return "True";
                else
                    return "False";
            }
            catch (Exception ex)
            {
                return "False" + ex.Message;
            }
        }
        private string Cau125(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 2)
                    return "False";
                return !a.ActivePresentation.Slides[(object)3].Shapes[(object)2].Name.Contains("Group") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau126(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides.Count != 4)
                    return "False";
                return a.ActivePresentation.Slides[(object)4].Shapes.Count != 2 || a.ActivePresentation.Slides[(object)4].Shapes[(object)1].TextFrame.TextRange.Text != "Top Sellers: " ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau127(Application a, Presentation d)
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
    }
}
