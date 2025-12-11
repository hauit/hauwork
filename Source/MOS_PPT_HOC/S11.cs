// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S11
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.IO;
using System.Reflection;


namespace MOS_PPT_LEARN
{
    internal class S11
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S11.Cau1(a, d);
                case 2:
                    return S11.Cau2(a, d);
                case 3:
                    return S11.Cau3(a, d);
                case 4:
                    return S11.Cau4(a, d);
                case 5:
                    return S11.Cau5(a, d);
                case 6:
                    return S11.Cau6(a, d);
                case 7:
                    return S11.Cau7(a, d);
                case 8:
                    return S11.Cau8(a, d);
                case 9:
                    return S11.Cau9(a, d);
                case 10:
                    return S11.Cau10(a, d);
                case 11:
                    return S11.Cau11(a, d);
                case 12:
                    return S11.Cau12(a, d);
                case 13:
                    return S11.Cau13(a, d);
                case 14:
                    return S11.Cau14(a, d);
                case 15:
                    return S11.Cau15(a, d);
                case 16 /*0x10*/:
                    return S11.Cau16(a, d);
                case 17:
                    return S11.Cau17(a, d);
                case 18:
                    return S11.Cau18(a, d);
                default:
                    return "case out index";
            }
        }

        private static string Cau1(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence.Count != 1)
                    return "False (khong them xoa hieu ung)";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[1].DisplayName != "Luxembourg")
                    return "False (Luxembourg)";
                if (a.ActivePresentation.Slides[(object)5].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectPathCircle")
                    return "False (circle)";
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
                if (a.ActivePresentation.SnapToGrid != MsoTriState.msoTrue)
                    return "False(snap object)";
                if (a.DisplayGridLines != MsoTriState.msoTrue)
                    return "False(griline)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string Cau3(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)8].Shapes[(object)"Heart 3"].Width.ToString() != "569.6")
                    return "False(Width)";
                if (a.ActivePresentation.Slides[(object)8].Shapes[(object)"Heart 3"].Height.ToString() != "505.6")
                    return "False(Height)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string Cau4(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].SlideShowTransition.SoundEffect.Name != "applause.wav")
                    return "False(Transition sound applause)";
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
                if (!a.ActivePresentation.Slides[(object)9].Shapes[(object)"Title 1"].TextFrame.TextRange.Text.Contains("Extra"))
                    return "False(sai outline)";
            }
            catch (Exception ex)
            {
                return "False(add slide tu outline file)";
            }
            return "True";
        }

        private static string Cau6(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)3].Shapes[(object)"Table 5"].Table.Rows.Count.ToString() != "8")
                    return "False";
                return a.ActivePresentation.Slides[(object)3].Shapes[(object)"Table 5"].Table.Rows[8].Cells[1].Shape.TextFrame.TextRange.Text != "Sporting Event" ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string Cau7(Application a, Presentation d)
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

        private static string Cau8(Application a, Presentation d)
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

        private static string Cau9(Application a, Presentation d)
        {
            //try
            //{
            //    if (a.ActivePresentation.Slides[(object)1].Shapes.Count > 2)
            //        return "False";
            //    if (a.ActivePresentation.Slides[(object)2].Shapes.Count != 3)
            //        return "False";
            //    if (a.ActivePresentation.Slides[(object)3].Shapes.Count != 5)
            //        return "False";
            //}
            //catch (Exception ex)
            //{
            //    return "False";
            //}
            //return "True";

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

        private static string Cau10(Application a, Presentation d)
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

        private static string Cau11(Application a, Presentation d)
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

        private static string Cau12(Application a, Presentation d)
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

        private static string Cau13(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)5].Background.Fill.Transparency.ToString() != "0.75")
                    return "False";
                if (a.ActivePresentation.Slides[(object)5].Background.Fill.Type != MsoFillType.msoFillPicture)
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].Background.Fill.Transparency.ToString() == "0.75")
                    return "False";
                if (a.ActivePresentation.Slides[(object)4].Background.Fill.Type == MsoFillType.msoFillPicture)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private static string Cau14(Application a, Presentation d)
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

        private static string Cau15(Application a, Presentation d)
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

        private static string Cau16(Application a, Presentation d)
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

        private static string Cau17(Application a, Presentation d)
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

        private static string Cau18(Application a, Presentation d)
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

        private static string Cau19(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)4].Layout.ToString() != "ppLayoutTwoObjects")
                    return "False(Layout)";
            }
            catch (Exception ex)
            {
                return "False(reused slide)";
            }
            return "True";
        }

        private static string Cau20(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)6].Shapes[(object)"Heart 1"].Left.ToString() != "374.7")
                    return "False()";
                if (a.ActivePresentation.Slides[(object)6].Shapes[(object)"Smiley Face 2"].Left.ToString() != "435")
                    return "False()";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string Cau21(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.HandoutMaster.Shapes.Count.ToString() != "3")
                    return "False(bo date placehonder trong HandoutMaster)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string Cau22(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)8].Shapes[(object)"TextBox 4"].TextFrame2.TextRange.Font.Spacing.ToString() != "6")
                    return "False(6pt)";
                if (((object)a.ActivePresentation.Slides[(object)8].Shapes[(object)"TextBox 4"].TextFrame2.TextRange.Font.Shadow.Visible).ToString() != "msoTrue")
                    return "False(shadow)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string cau6(Application a, Presentation d)
        {
            try
            {
                return a.ActivePresentation.Slides[(object)3].Shapes[(object)"Table 5"].Table.Rows.Count.ToString() != "6" ? "False(thêm dòng vào table)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string cau7(Application a, Presentation d)
        {
            try
            {
                string str = "True";
                if (a.ActivePresentation.PrintOptions.OutputType != PpPrintOutputType.ppPrintOutputThreeSlideHandouts)
                    return "False (Handout three Slide)";
                return a.ActivePresentation.PrintOptions.FrameSlides != MsoTriState.msoTrue ? "False (FrameSlides)" : str;
            }
            catch (Exception ex)
            {
                return "False (Somthing Wrong)";
            }
        }

        private static string cau8(Application a, Presentation d)
        {
            try
            {
                return new FileInfo(a.ActivePresentation.FullName).Length < 1800000L ? "False (saved)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Somthing Wrong)";
            }
        }

        private static string cau9(Application a, Presentation d) => "True";

        private static string cau10(Application a, Presentation d)
        {
            try
            {
                object documentProperties = a.ActivePresentation.BuiltInDocumentProperties;
                object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (Binder)null, documentProperties, new object[1]
                {
        (object) "Title"
                });
                if (target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (Binder)null, target, new object[0]).ToString() == "")
                    return "False (not remove properties)";
            }
            catch (Exception ex)
            {
                return "False ()";
            }
            try
            {
                string str = "True";
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count == 4)
                    return "False (On Slide Content)";
                return a.ActivePresentation.Slides[(object)7].Shapes.Count == 4 ? "False (Off Slide Content)" : str;
            }
            catch (Exception ex)
            {
                return "Something Wrong";
            }
        }

        private static string cau11(Application a, Presentation d)
        {
            try
            {
                string str = "True";
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count == 4)
                    return "False (On Slide Content)";
                return a.ActivePresentation.Slides[(object)7].Shapes.Count == 4 ? "False (Off Slide Content)" : str;
            }
            catch (Exception ex)
            {
                return "False ()";
            }
        }

        private static string cau12(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count != 2)
                    return "False(chèn Summary Zoom link)";
                return !a.ActivePresentation.Slides[(object)2].Shapes[(object)2].Name.Contains("Summary Zoom") ? "False(Summary Zoom link)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private static string cau13(Application a, Presentation d)
        {
            try
            {
                return new FileInfo(a.ActivePresentation.FullName).Length < 1800000L ? "False (saved)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Somthing Wrong)";
            }
        }
    }
}