// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S1_Illustrator
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System;
using Shape = Microsoft.Office.Core.Shape;


namespace MOS_PPT_LEARN
{

    internal class S1_Illustrator
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S1_Illustrator.cau1(a, d);
                case 2:
                    return S1_Illustrator.cau2(a, d);
                case 3:
                    return S1_Illustrator.cau3(a, d);
                case 4:
                    return S1_Illustrator.cau4(a, d);
                case 5:
                    return S1_Illustrator.cau5(a, d);
                case 6:
                    return S1_Illustrator.cau6(a, d);
                case 7:
                    return S1_Illustrator.cau7(a, d);
                case 8:
                    return S1_Illustrator.cau8(a, d);
                case 9:
                    return S1_Illustrator.cau9(a, d);
                case 10:
                    return S1_Illustrator.cau10(a, d);
                case 11:
                    return S1_Illustrator.cau11(a, d);
                case 12:
                    return S1_Illustrator.cau12(a, d);
                case 13:
                    return S1_Illustrator.cau13(a, d);
                case 14:
                    return S1_Illustrator.cau14(a, d);
                case 15:
                    return S1_Illustrator.cau15(a, d);
                case 16 /*0x10*/:
                    return S1_Illustrator.cau16(a, d);
                case 17:
                    return S1_Illustrator.cau17(a, d);
                case 18:
                    return S1_Illustrator.cau18(a, d);
                case 19:
                    return S1_Illustrator.cau19(a, d);
                case 20:
                    return S1_Illustrator.cau20(a, d);
                case 21:
                    return S1_Illustrator.cau21(a, d);
                case 22:
                    return S1_Illustrator.cau22(a, d);
                case 23:
                    return S1_Illustrator.cau23(a, d);
                case 24:
                    return S1_Illustrator.cau24(a, d);
                case 25:
                    return S1_Illustrator.cau25(a, d);
                case 26:
                    return S1_Illustrator.cau26(a, d);
                case 27:
                    return S1_Illustrator.cau27(a, d);
                case 28:
                    return S1_Illustrator.cau28(a, d);
                case 29:
                    return S1_Illustrator.cau29(a, d);
                case 30:
                    return S1_Illustrator.cau30(a, d);
                case 31 /*0x1F*/:
                    return S1_Illustrator.cau31(a, d);
                default:
                    return "case 31";
            }
        }

        private static string cau1(Application a, Presentation d)
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

        private static string cau2(Application a, Presentation d)
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

        private static string cau3(Application a, Presentation d)
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

        private static string cau4(Application a, Presentation d)
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

        private static string cau5(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 3)
                    return "False(khong them xoa shape)";
                return !a.ActivePresentation.Slides[(object)3].Shapes[(object)3].Name.Contains("Heart") ? "False(doi thanh hinh trai tim)" : "True";
            }
            catch (Exception ex)
            {
                return "False (add cho text khong phai cho shape)";
            }
        }

        private static string cau6(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count.ToString() != "6")
                    return "False(group)";
                if (!a.ActivePresentation.Slides[(object)2].Shapes[(object)6].Name.Contains("Group"))
                    return "False(Group)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string cau7(Application a, Presentation d)
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

        private static string cau8(Application a, Presentation d)
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

        private static string cau9(Application a, Presentation d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False (Convert to smartArt)";
            }
        }

        private static string cau10(Application a, Presentation d)
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

        private static string cau11(Application a, Presentation d)
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

        private static string cau12(Application a, Presentation d)
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

        private static string cau13(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)2].Shapes[(object)"Heart 7"].Fill.ForeColor.RGB.ToString() != "7826629" ? "False(shape style)" : "True";
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
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 2)
                    return "False";
                return !a.ActivePresentation.Slides[(object)3].Shapes[(object)2].Name.Contains("Group") ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau15(Application a, Presentation d)
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

        private static string cau16(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 3)
                    return "False";
                if (((object)a.ActivePresentation.Slides[(object)3].Shapes[(object)2].Type).ToString() != "msoPlaceholder")
                    return "False";
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)2].Height.ToString() != "288")
                    return "False";
                return a.ActivePresentation.Slides[(object)3].Shapes[(object)2].Width.ToString() != "134.255" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau17(Application a, Presentation d)
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

        private static string cau18(Application a, Presentation d)
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
                return a.ActivePresentation.Slides[(object)4].Shapes[(object)"Freeform 5"].Shadow.OffsetY.ToString() != "-2.12132" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau19(Application a, Presentation d)
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

        private static string cau20(Application a, Presentation d)
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

        private static string cau21(Application a, Presentation d)
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

        private static string cau22(Application a, Presentation d)
        {
            try
            {
                var slide = a.ActivePresentation.Slides[(object)3];  // slide 3
                if (slide.Shapes.Count != 4) return "False 1";

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

        private static string cau23(Application a, Presentation d)
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

        private static string cau24(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 2)
                    return "False";
                return a.ActivePresentation.Slides[(object)4].Shapes[(object)2].SmartArt.Layout.Name != "Segmented Cycle" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau25(Application a, Presentation d)
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

        private static string cau26(Application a, Presentation d)
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

        private static string cau27(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)5].Shapes.Count != 6 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau28(Application a, Presentation d)
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

        private static string cau29(Application a, Presentation d)
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

        private static string cau30(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)5].Shapes.Count != 3 ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau31(Application a, Presentation d)
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
    }
}