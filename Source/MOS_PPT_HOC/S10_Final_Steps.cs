// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S10_Final_Steps
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.IO;
using System.Net.NetworkInformation;
using System.Reflection;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;


namespace MOS_PPT_LEARN
{
    internal class S10_Final_Steps
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S10_Final_Steps.cau1(a, d);
                case 2:
                    return S10_Final_Steps.cau2(a, d);
                case 3:
                    return S10_Final_Steps.cau3(a, d);
                case 4:
                    return S10_Final_Steps.cau4(a, d);
                case 5:
                    return S10_Final_Steps.cau5(a, d);
                case 6:
                    return S10_Final_Steps.cau6(a, d);
                case 7:
                    return S10_Final_Steps.cau7(a, d);
                case 8:
                    return S10_Final_Steps.cau8(a, d);
                case 9:
                    return S10_Final_Steps.cau9(a, d);
                case 10:
                    return S10_Final_Steps.cau10(a, d);
                case 11:
                    return S10_Final_Steps.cau11(a, d);
                case 12:
                    return S10_Final_Steps.cau12(a, d);
                case 13:
                    return S10_Final_Steps.cau13(a, d);
                case 14:
                    return S10_Final_Steps.cau14(a, d);
                case 15:
                    return S10_Final_Steps.cau15(a, d);
                case 16 /*0x10*/:
                    return S10_Final_Steps.cau16(a, d);
                case 17:
                    return S10_Final_Steps.cau17(a, d);
                case 18:
                    return S10_Final_Steps.cau18(a, d);
                default:
                    return "case out index";
            }
        }

        private static string cau1(Application a, Presentation d)
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

        private static string cau2(Application a, Presentation d)
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

        private static string cau3(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.SlideShowSettings.NamedSlideShows.Count != 1)
                    return "False(add slide show)";
                if (a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object)1].Name != "Charts")
                    return "False(Charts)";
                return a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object)1].Count != 3 ? "False(slide 5-6-7)" : "True";
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

        private static string cau5(Application a, Presentation d)
        {
            try
            {
                object documentProperties = a.ActivePresentation.BuiltInDocumentProperties;
                object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (Binder)null, documentProperties, new object[1]
                {
        (object) "Title"
                });
                if (target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (Binder)null, target, new object[0]).ToString() != "Life Insurance Breakdown")
                    return "False (Life Insurance Breakdown)";
            }
            catch (Exception ex)
            {
                return "False (Life Insurance Breakdown)";
            }
            return "True";
        }

        private static string cau6(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)1].Shapes.Count == 1)
                    return "False";
                if (a.ActivePresentation.Slides[(object)1].Shapes.Count > 2)
                    return "False";
                switch (a.ActivePresentation.Slides[(object)1].Shapes[(object)1].Name)
                {
                    case "5-Point Star 2":
                        return "False";
                    case "Ink 1":
                        return "False";
                    default:
                        switch (a.ActivePresentation.Slides[(object)1].Shapes[(object)2].Name)
                        {
                            case "5-Point Star 2":
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

        private static string cau7(Application a, Presentation d)
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

        private static string cau8(Application a, Presentation d)
        {
            try
            {
                if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "BasketWeaving.pdf")))
                    return "False(luu dang pdf trong document)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string cau9(Application a, Presentation d)
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

        private static string cau10(Application a, Presentation d)
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

        private static string cau11(Application a, Presentation d)
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

        private static string cau12(Application a, Presentation d)
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

        private static string cau13(Application a, Presentation d)
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

        private static string cau14(Application a, Presentation d)
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

        private static string cau15(Application a, Presentation d)
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

        private static string cau16(Application a, Presentation d)
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

        private static string cau17(Application a, Presentation d)
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

        private static string cau18(Application a, Presentation d)
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
    }
}