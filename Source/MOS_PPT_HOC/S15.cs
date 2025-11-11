// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S15
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Interop.PowerPoint;
using System;
using System.IO;


namespace MOS_PPT_LEARN
{
    internal class S15
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S15.Cau1(a, d);
                case 2:
                    return S15.Cau2(a, d);
                case 3:
                    return S15.Cau3(a, d);
                case 4:
                    return S15.Cau4(a, d);
                case 5:
                    return S15.Cau5(a, d);
                case 6:
                    return S15.Cau6(a, d);
                case 7:
                    return S15.Cau7(a, d);
                default:
                    return "case out index";
            }
        }

        private static string Cau1(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.PageSetup.SlideWidth.ToString() != "576")
                    return "False(Width)";
                if (a.ActivePresentation.PageSetup.SlideHeight.ToString() != "792")
                    return "False(Height)";
                float num = a.ActivePresentation.Slides[(object)1].Shapes[(object)"Picture 3"].Width;
                if (num.ToString() != "444.662")
                    return "False(Fit)";
                num = a.ActivePresentation.Slides[(object)1].Shapes[(object)"Picture 3"].Height;
                if (num.ToString() != "229.982")
                    return "False(Fit)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string Cau2(Application a, Presentation d)
        {
            try
            {
                int count = a.ActivePresentation.SlideShowSettings.NamedSlideShows.Count;
                if (count.ToString() != "1")
                    return "False (add SlideShow)";
                if (a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object)1].Name != "Review")
                    return "False (SlideShow name)";
                count = a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object)1].Count;
                return count.ToString() != "8" ? "False (Number of slide in slide show)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something Wrong)";
            }
        }

        private static string Cau3(Application a, Presentation d)
        {
            try
            {
                return new FileInfo(a.ActivePresentation.FullName).Length < 7692100L ? "False (saved)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Somthing Wrong)";
            }
        }

        private static string Cau4(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)9].TimeLine.MainSequence.Count.ToString() != "1")
                    return "False(Auto)";
                if (a.ActivePresentation.Slides[(object)9].TimeLine.MainSequence[1].Timing.TriggerType.ToString() != "msoAnimTriggerAfterPrevious")
                    return "False(auto)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
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