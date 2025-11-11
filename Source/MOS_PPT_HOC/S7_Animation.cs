// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S7_Animation
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Interop.PowerPoint;
using System;


namespace MOS_PPT_LEARN
{
    internal class S7_Animation
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S7_Animation.cau1(a, d);
                case 2:
                    return S7_Animation.cau2(a, d);
                case 3:
                    return S7_Animation.cau3(a, d);
                case 4:
                    return S7_Animation.cau4(a, d);
                case 5:
                    return S7_Animation.cau5(a, d);
                case 6:
                    return S7_Animation.cau6(a, d);
                case 7:
                    return S7_Animation.cau7(a, d);
                case 8:
                    return S7_Animation.cau8(a, d);
                case 9:
                    return S7_Animation.cau9(a, d);
                case 10:
                    return S7_Animation.cau10(a, d);
                case 11:
                    return S7_Animation.cau11(a, d);
                case 12:
                    return S7_Animation.cau12(a, d);
                default:
                    return "case 11";
            }
        }

        private static string cau1(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence.Count != 4)
                    return "False(khong them xoa hieu ung)";
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[1].DisplayName != "Picture 3")
                    return "False(picture 3)";
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[2].DisplayName != "Picture 4")
                    return "False(picture 4)";
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[3].DisplayName != "Picture 5")
                    return "False(picture 5)";
                return a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[4].DisplayName != "Picture 6" ? "False(picture 6)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Somthing Wrong)";
            }
        }

        private static string cau2(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence.Count != 1)
                    return "False(add hieu ung)";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].DisplayName != "Picture 5")
                    return "False(Hieu ung cho xe)";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectType != MsoAnimEffect.msoAnimEffectFly)
                    return "False(sai Hieu Ung)";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectParameters.Direction != MsoAnimDirection.msoAnimDirectionRight)
                    return "False(sai huong)";
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
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence.Count != 5)
                    return "False(add hieu ung)";
                if (!a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].DisplayName.Contains("Choose by 95%"))
                    return "False(Choose by 95%)";
                if (!a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[5].DisplayName.Contains("Top 10"))
                    return "False(Top 10)";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectWipe")
                    return "False(Wipe)";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectParameters.Direction.ToString() != "msoAnimDirectionLeft")
                    return "False(Left)";
                return a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[5].Timing.TriggerType.ToString() != "msoAnimTriggerOnPageClick" ? "False(OnClick)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something Wrong)";
            }
        }

        private static string cau4(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence.Count != 3)
                    return "False(khong them xoa hieu ung)";
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[1].DisplayName != "Picture 5")
                    return "False(picture 5)";
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[2].DisplayName != "Picture 2")
                    return "False(picture 2)";
                return a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[3].DisplayName != "Picture 8" ? "False(picture 8)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong )";
            }
        }

        private static string cau5(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence.Count != 1)
                    return "False(add hieu ung)";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].DisplayName != "Content Placeholder 4")
                    return "False(Hieu ung cho xe)";
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectType != MsoAnimEffect.msoAnimEffectFly)
                    return "False(sai Hieu Ung)";
                return a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectParameters.Direction != MsoAnimDirection.msoAnimDirectionDown ? "False(sai huong)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong )";
            }
        }

        private static string cau6(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence.Count.ToString() != "1")
                    return "False (Number of animation)";
                if (a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[1].DisplayName != "No Way!")
                    return "False (sai shape)";
                return a.ActivePresentation.Slides[(object)2].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectPathCircle" ? "False (sai hieu ung)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong )";
            }
        }

        private static string cau7(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].DisplayName != "Picture 3")
                    return "False (Picture 3)";
                return a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectFadedSwivel" ? "False (Swivel)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Add Animation)";
            }
        }

        private static string cau8(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].DisplayName != "3D Model 3")
                    return "False (3D Model)";
                return a.ActivePresentation.Slides[(object)3].TimeLine.MainSequence[1].EffectType.ToString() != "154" ? "False (sai hiệu ứng)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Add Animation)";
            }
        }

        private static string cau9(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence.Count != 1)
                    return "False(không thêm xóa hiệu Ứng)";
                return a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectPathDown" ? "False (hướng Down)" : "True";
            }
            catch (Exception ex)
            {
                return "False (không xác định)";
            }
        }

        private static string cau10(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence.Count != 4)
                    return "False(không thêm xóa hiệu Ứng)";
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].EffectParameters.Direction.ToString() != "msoAnimDirectionLeft")
                    return "False (hướng Left)";
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].Timing.Duration.ToString() != "1.5")
                    return "False (Duration 1.5)";
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[4].EffectParameters.Direction.ToString() != "msoAnimDirectionLeft")
                    return "False (hướng Left)";
                return a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[4].Timing.Duration.ToString() != "1.5" ? "False (Duration 1.5)" : "True";
            }
            catch (Exception ex)
            {
                return "False (không xác định)";
            }
        }

        private static string cau11(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence.Count != 1)
                    return "False(1 hiệu Ứng cho ngôi sao)";
                if (a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].DisplayName != "5-Point Star 5")
                    return "False (cho ngôi sao)";
                return a.ActivePresentation.Slides[(object)4].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectPathHeart" ? "False (sai kiểu)" : "True";
            }
            catch (Exception ex)
            {
                return "False (không xác định)";
            }
        }

        private static string cau12(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)6].TimeLine.MainSequence.Count != 1)
                    return "False(1 hiệu Ứng cho ảnh máy bay)";
                if (a.ActivePresentation.Slides[(object)6].TimeLine.MainSequence[1].DisplayName != "Picture 4")
                    return "False (hiệu Ứng cho ảnh máy bay)";
                if (a.ActivePresentation.Slides[(object)6].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectFly")
                    return "False (sai kiểu)";
                if (a.ActivePresentation.Slides[(object)6].TimeLine.MainSequence[1].EffectParameters.Direction.ToString() != "msoAnimDirectionUpLeft")
                    return "False (sai hướng)";
                return a.ActivePresentation.Slides[(object)6].TimeLine.MainSequence[1].Timing.Duration.ToString() != "2" ? "False (Duration 2)" : "True";
            }
            catch (Exception ex)
            {
                return "False (không xác định)";
            }
        }
    }
}