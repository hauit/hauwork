// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S3_AddSlide
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;


namespace MOS_PPT_LEARN
{
    internal class S3_AddSlide
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S3_AddSlide.cau1(a, d);
                case 2:
                    return S3_AddSlide.cau2(a, d);
                case 3:
                    return S3_AddSlide.cau3(a, d);
                case 4:
                    return S3_AddSlide.cau4(a, d);
                case 5:
                    return S3_AddSlide.cau5(a, d);
                case 6:
                    return S3_AddSlide.cau6(a, d);
                case 7:
                    return S3_AddSlide.cau7(a, d);
                case 8:
                    return S3_AddSlide.cau8(a, d);
                case 9:
                    return S3_AddSlide.cau9(a, d);
                case 10:
                    return S3_AddSlide.cau10(a, d);
                case 11:
                    return S3_AddSlide.cau11(a, d);
                default:
                    return "case Out Indext";
            }
        }

        private static string cau1(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Title 1"].TextFrame.TextRange.Text != "Extra")
                    return "False(chen sai vi tri hoac sai outline)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string cau2(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides.Count != 4)
                    return "False(add slide from outline)";
                if (!a.ActivePresentation.Slides[(object)4].Shapes[(object)1].TextFrame.TextRange.Text.Contains("Certificate"))
                    return "False(Vanessa)";
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
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 2)
                    return "False(chen slide)";
                return !a.ActivePresentation.Slides[(object)3].Shapes[(object)1].TextFrame.TextRange.Text.Contains("New product") ? "False(sai outline)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position)";
            }
        }

        private static string cau4(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 2)
                    return "False(chen slide)";
                return !a.ActivePresentation.Slides[(object)4].Shapes[(object)1].TextFrame.TextRange.Text.Contains("New product") ? "False(sai outline)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position)";
            }
        }

        private static string cau5(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)6].Shapes.Count != 2 || !a.ActivePresentation.Slides[(object)6].Shapes[(object)1].TextFrame.TextRange.Text.Contains("Student") ? "False(Doplicate slide 5)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position)";
            }
        }

        private static string cau6(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides.Count != 4)
                    return "False(xoa slide)";
                return a.ActivePresentation.Slides[(object)4].Shapes.Count != 2 || a.ActivePresentation.Slides[(object)4].Shapes[(object)1].TextFrame.TextRange.Text != "Top Sellers: " ? "False(sai slide)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position)";
            }
        }

        private static string cau7(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 2)
                    return "False(add new slide from outline)";
                return a.ActivePresentation.Slides[(object)4].Shapes[(object)"Text Placeholder 2"].TextFrame.TextRange.Text != "Home stay\rHouse boat\r…" ? "False(Wrong file)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position)";
            }
        }

        private static string cau8(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)5].SlideShowTransition.Hidden != MsoTriState.msoTrue ? "False(ân slide 5)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position)";
            }
        }

        private static string cau9(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 2)
                    return "False(add new slide from outline)";
                return a.ActivePresentation.Slides[(object)3].Shapes[(object)"Text Placeholder 2"].TextFrame.TextRange.Text != "Home stay\rHouse boat\r…" ? "False(Wrong file)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position)";
            }
        }

        private static string cau10(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)6].Shapes.Count != 2)
                    return "False(number of shape)";
                if (a.ActivePresentation.Slides[(object)6].Shapes[(object)1].TextFrame.TextRange.Text != "Discover Your Campus")
                    return "False(Sai vi tri slide)";
                if (a.ActivePresentation.Slides[(object)7].Shapes.Count != 2)
                    return "False(number of shape)";
                return a.ActivePresentation.Slides[(object)7].Shapes[(object)1].TextFrame.TextRange.Text != "Display Your Art" ? "False(Sai vi tri slide)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position)";
            }
        }

        private static string cau11(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 2)
                    return "False(number of shape)";
                return a.ActivePresentation.Slides[(object)3].Shapes[(object)"Text Placeholder 2"].TextFrame.TextRange.Text != "Show where you need them\rTo change the way a picture fits in your document\rClick it and a button for layout options appears next to it" ? "False(Wrong file)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Wrong position)";
            }
        }
    }
}