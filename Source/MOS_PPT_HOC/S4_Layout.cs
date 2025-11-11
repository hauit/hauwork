// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S4_Layout
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Interop.PowerPoint;
using System;


namespace MOS_PPT_LEARN
{
    internal class S4_Layout
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S4_Layout.cau1(a, d);
                case 2:
                    return S4_Layout.cau2(a, d);
                case 3:
                    return S4_Layout.cau3(a, d);
                case 4:
                    return S4_Layout.cau4(a, d);
                case 5:
                    return S4_Layout.cau5(a, d);
                case 6:
                    return S4_Layout.cau6(a, d);
                case 7:
                    return S4_Layout.cau7(a, d);
                case 8:
                    return S4_Layout.cau8(a, d);
                case 9:
                    return S4_Layout.cau9(a, d);
                case 10:
                    return S4_Layout.cau10(a, d);
                default:
                    return "case out indext";
            }
        }

        private static string cau1(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)2].Layout.ToString() != "ppLayoutCustom" ? "False(Highlights)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something Wrong)";
            }
        }

        private static string cau2(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides.Count == 3)
                {
                    if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Content Placeholder 2"].TextFrame2.Column.Number != 2)
                        return "False(2 cot)";
                }
                else
                {
                    if (a.ActivePresentation.Slides.Count != 4)
                        return "False(slide da bi them xoa qua nhieu)";
                    if (a.ActivePresentation.Slides[(object)4].Shapes[(object)"Content Placeholder 2"].TextFrame2.Column.Number != 2)
                        return "False(2 cot)";
                }
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string cau3(Application a, Presentation d)
        {
            try
            {
                if (!a.ActivePresentation.Slides[(object)2].Shapes[(object)"Title 1"].TextFrame.TextRange.Text.Contains("Strategy 2020"))
                    return "False (Not new slide or Order slides)";
                return a.ActivePresentation.Slides[(object)2].Layout != PpSlideLayout.ppLayoutContentWithCaption ? "False (Layout)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position Slide)";
            }
        }

        private static string cau4(Application a, Presentation d)
        {
            try
            {
                if (!a.ActivePresentation.Slides[(object)3].Shapes[(object)"Title 1"].TextFrame.TextRange.Text.Contains("Revenue by Product Lines"))
                    return "False (Not new slide or Order slides)";
                return a.ActivePresentation.Slides[(object)3].Layout != PpSlideLayout.ppLayoutObject ? "False (Layout)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position Slide)";
            }
        }

        private static string cau5(Application a, Presentation d)
        {
            try
            {
                if (!a.ActivePresentation.Slides[(object)3].Shapes[(object)"Title 1"].TextFrame.TextRange.Text.Contains("Project"))
                    return "False (Not new slide or Order slides)";
                return a.ActivePresentation.Slides[(object)3].Layout != PpSlideLayout.ppLayoutSectionHeader ? "False (Layout)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position Slide)";
            }
        }

        private static string cau6(Application a, Presentation d)
        {
            try
            {
                if (!a.ActivePresentation.Slides[(object)3].Shapes[(object)"Title 1"].TextFrame.TextRange.Text.Contains("Advantage"))
                    return "False (Not new slide or Order slides)";
                return a.ActivePresentation.Slides[(object)3].Layout != PpSlideLayout.ppLayoutTwoObjects ? "False (Layout)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position Slide)";
            }
        }

        private static string cau7(Application a, Presentation d) => "True";

        private static string cau8(Application a, Presentation d) => "True";

        private static string cau9(Application a, Presentation d) => "True";

        private static string cau10(Application a, Presentation d) => "True";
    }
}