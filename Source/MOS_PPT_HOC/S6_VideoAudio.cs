// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S6_VideoAudio
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Interop.PowerPoint;
using System;


namespace MOS_PPT_LEARN
{
    internal class S6_VideoAudio
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S6_VideoAudio.cau1(a, d);
                case 2:
                    return S6_VideoAudio.cau2(a, d);
                case 3:
                    return S6_VideoAudio.cau3(a, d);
                case 4:
                    return S6_VideoAudio.cau4(a, d);
                case 5:
                    return S6_VideoAudio.cau5(a, d);
                case 6:
                    return S6_VideoAudio.cau6(a, d);
                case 7:
                    return S6_VideoAudio.cau7(a, d);
                case 8:
                    return S6_VideoAudio.cau8(a, d);
                case 9:
                    return S6_VideoAudio.cau9(a, d);
                case 10:
                    return S6_VideoAudio.cau10(a, d);
                default:
                    return "case 11";
            }
        }

        private static string cau1(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 2)
                    return "False(khong them xoa cac doi tuong khac)";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)2].Name != "New Advert")
                    return "False(New Advert)";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)2].Top.ToString() != "143.75")
                    return "False(H)";
                if (a.ActivePresentation.Slides[(object)4].Shapes[(object)2].Left.ToString() != "144")
                    return "False(V)";
            }
            catch (Exception ex)
            {
                return "False(Gradient)";
            }
            return "True";
        }

        private static string cau2(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes[(object)"New Advert"].MediaFormat.StartPoint.ToString() != "500")
                    return "False(StartPoint)";
                if (a.ActivePresentation.Slides[(object)2].Shapes[(object)"New Advert"].MediaFormat.EndPoint.ToString() != "2500")
                    return "False(EndPoint)";
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
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence.Count.ToString() != "1")
                    return "False(Auto)";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].Timing.TriggerType.ToString() != "msoAnimTriggerAfterPrevious")
                    return "False(auto)";
            }
            catch (Exception ex)
            {
                return "False (khong xat dinh)";
            }
            return "True";
        }

        private static string cau4(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Powerpoint mos vid"].PictureFormat.CropLeft.ToString() != "487.268")
                    return "False(left)";
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Powerpoint mos vid"].PictureFormat.CropRight.ToString() != "-7.277816")
                    return "False(right)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string cau5(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)1].Shapes[(object)"Sleep Away"].MediaFormat.FadeInDuration != 2000)
                    return "False (FadeInDuration 2)";
                return a.ActivePresentation.Slides[(object)1].Shapes[(object)"Sleep Away"].AnimationSettings.PlaySettings.StopAfterSlides != 999 ? "False (Play cross slide)" : "True";
            }
            catch (Exception ex)
            {
                return "False (insert Train vedio into placeholder)";
            }
        }

        private static string cau6(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)5].Shapes.Count != 3)
                    return "False(chèn video)";
                return a.ActivePresentation.Slides[(object)5].Shapes[(object)3].Name != "Sailing" ? "False (sai vedio)" : "True";
            }
            catch (Exception ex)
            {
                return "False (insert Sailing video)";
            }
        }

        private static string cau7(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)6].Shapes.Count != 3)
                    return "False(chèn video)";
                return a.ActivePresentation.Slides[(object)6].Shapes[(object)3].Name != "River" ? "False (sai vedio)" : "True";
            }
            catch (Exception ex)
            {
                return "False (insert Train vedio into placeholder)";
            }
        }

        private static string cau8(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].Timing.Duration.ToString() != "4" ? "False (Duration 04)" : "True";
            }
            catch (Exception ex)
            {
                return "False (insert Train vedio into placeholder)";
            }
        }

        private static string cau9(Application a, Presentation d) => "True";

        private static string cau10(Application a, Presentation d) => "True";
    }
}