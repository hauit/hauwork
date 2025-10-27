// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S10_Final_Steps
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Interop.PowerPoint;
using System;
using System.IO;
using System.Reflection;


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
                    return "False(luu dang pdf trong document)";
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
                if (a.ActivePresentation.PrintOptions.sectionIndex != 2)
                    return "False(introduction)";
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

        private static string cau4(Application a, Presentation d) => "True";

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
                    return "False(khong xoa invisible conten)";
                if (a.ActivePresentation.Slides[(object)1].Shapes.Count > 2)
                    return "False(inpect document)";
                switch (a.ActivePresentation.Slides[(object)1].Shapes[(object)1].Name)
                {
                    case "5-Point Star 2":
                        return "False(xoa out off slide)";
                    case "Ink 1":
                        return "False(xoa ink)";
                    default:
                        switch (a.ActivePresentation.Slides[(object)1].Shapes[(object)2].Name)
                        {
                            case "5-Point Star 2":
                                return "False(xoa out off slide)";
                            case "Ink 1":
                                return "False(xoa ink)";
                        }
                        break;
                }
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
                if (a.ActivePresentation.PrintOptions.OutputType.ToString() != "ppPrintOutputNotesPages")
                    return "False(ppPrintOutputNotesPages)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
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
                    return "False(Course Introduction)";
            }
            catch (Exception ex)
            {
                return "False(loi khong xac dinh)";
            }
            return "True";
        }

        private static string cau10(Application a, Presentation d) => "True";

        private static string cau11(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.SlideShowSettings.NamedSlideShows.Count != 1)
                    return "False(add slide show)";
                if (a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object)1].Name != "Important Findings")
                    return "False(Important Findings)";
                return a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object)1].Count != 2 ? "False(chi slide 3 va 5)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Something Wrong)";
            }
        }

        private static string cau12(Application a, Presentation d)
        {
            try
            {
                return new FileInfo(a.ActivePresentation.FullName).Length < 62878L ? "False (saved)" : "True";
            }
            catch (Exception ex)
            {
                return "False (Somthing Wrong)";
            }
        }

        private static string cau13(Application a, Presentation d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False (Somthing Wrong)";
            }
        }

        private static string cau14(Application a, Presentation d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False (Somthing Wrong)";
            }
        }

        private static string cau15(Application a, Presentation d)
        {
            try
            {
                return "True";
            }
            catch (Exception ex)
            {
                return "False (Somthing Wrong)";
            }
        }

        private static string cau16(Application a, Presentation d)
        {
            try
            {
                object documentProperties = a.ActivePresentation.BuiltInDocumentProperties;
                object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (Binder)null, documentProperties, new object[1]
                {
        (object) "Category"
                });
                string str = target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (Binder)null, target, new object[0]).ToString();
                if (str != "Travel")
                    return $"False ({str})";
            }
            catch (Exception ex)
            {
                return "False (không xác định)";
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
                    return "False (xóa các thuộc tính ẩn)";
                if (((object)a.ActivePresentation.RemovePersonalInformation).ToString() != "msoTrue")
                    return "False(dung chức năng RemovePersonalInformation)";
            }
            catch (Exception ex)
            {
                return "False (không xác định)";
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
                    return $"False ({str})";
            }
            catch (Exception ex)
            {
                return "False (không xác định)";
            }
            return "True";
        }
    }
}