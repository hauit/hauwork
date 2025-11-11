// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S17
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Interop.PowerPoint;
using System;


namespace MOS_PPT_LEARN
{
    internal class S17
    {
        public static string CheckCau(int causo, Application a, Presentation d)
        {
            switch (causo)
            {
                case 1:
                    return S17.Cau1(a, d);
                case 2:
                    return S17.Cau2(a, d);
                case 3:
                    return S17.Cau3(a, d);
                case 4:
                    return S17.Cau4(a, d);
                case 5:
                    return S17.Cau5(a, d);
                case 6:
                    return S17.Cau6(a, d);
                case 7:
                    return S17.Cau7(a, d);
                default:
                    return "case out index";
            }
        }

        private static string Cau1(Application a, Presentation d)
        {
            try
            {
                if (a.ActivePresentation.Slides[(object)7].Shapes.Count != 2)
                    return "False(insert chart)";
                if (((object)a.ActivePresentation.Slides[(object)7].Shapes[(object)2].Type).ToString() != "msoChart")
                    return "False(chart)";
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
                if (a.ActivePresentation.Slides[(object)4].Shapes.Count != 2)
                    return "False(insert object)";
                return ((object)a.ActivePresentation.Slides[(object)4].Shapes[(object)2].Type).ToString() != "msoEmbeddedOLEObject" ? "False(Object)" : "True";
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
                if (a.ActivePresentation.Slides[(object)2].Shapes.Count != 2)
                    return "False(khong them xoa shape)";
                if (a.ActivePresentation.Slides[(object)2].Shapes[(object)"TextBox 4"].Fill.BackColor.RGB.ToString() != "14145397")
                    return "False(shape style)";
                if (a.ActivePresentation.Slides[(object)2].Shapes[(object)"TextBox 4"].Line.Weight.ToString() != "3")
                    return "False(duong vien)";
                return a.ActivePresentation.Slides[(object)2].Shapes[(object)"TextBox 4"].ThreeD.BevelTopDepth.ToString() != "6" ? "False(bevel)" : "True";
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
                if (a.ActivePresentation.Slides.Count != 10)
                    return "False(add slide from outline)";
                if (!a.ActivePresentation.Slides[(object)10].Shapes[(object)"Title 1"].TextFrame.TextRange.Text.Contains("New product"))
                    return "False(sai outline or sai vi tri)";
            }
            catch (Exception ex)
            {
                return "False(Somthing Wrong)";
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