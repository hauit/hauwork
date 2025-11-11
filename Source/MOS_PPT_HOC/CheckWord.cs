// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.CheckWord
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.IO;

namespace MOS_PPT_LEARN
{
    internal class CheckWord
    {
        public static bool CheckCau(int causo, Application a, Presentation d)
        {
            try
            {
                switch (causo)
                {
                    case 1:
                        return CheckWord.cau1(a, d);
                    case 2:
                        return CheckWord.cau2(a, d);
                    case 3:
                        return CheckWord.cau3(a, d);
                    case 4:
                        return CheckWord.cau4(a, d);
                    case 5:
                        return CheckWord.cau5(a, d);
                    case 6:
                        return CheckWord.cau6(a, d);
                    case 7:
                        return CheckWord.cau7(a, d);
                    case 8:
                        return CheckWord.cau8(a, d);
                    case 9:
                        return CheckWord.cau9(a, d);
                    case 10:
                        return CheckWord.cau10(a, d);
                    case 11:
                        return CheckWord.cau11(a, d);
                    case 12:
                        return CheckWord.cau12(a, d);
                    case 13:
                        return CheckWord.cau13(a, d);
                    case 14:
                        return CheckWord.cau14(a, d);
                    case 15:
                        return CheckWord.cau15(a, d);
                    case 16 /*0x10*/:
                        return CheckWord.cau16(a, d);
                    case 17:
                        return CheckWord.cau17(a, d);
                    case 18:
                        return CheckWord.cau18(a, d);
                    case 19:
                        return CheckWord.cau19(a, d);
                    case 20:
                        return CheckWord.cau20(a, d);
                    case 21:
                        return CheckWord.cau21(a, d);
                    case 22:
                        return CheckWord.cau22(a, d);
                    case 23:
                        return CheckWord.cau23(a, d);
                    case 24:
                        return CheckWord.cau24(a, d);
                    case 25:
                        return CheckWord.cau25(a, d);
                    case 26:
                        return CheckWord.cau26(a, d);
                    case 27:
                        return CheckWord.cau27(a, d);
                    case 28:
                        return CheckWord.cau28(a, d);
                    case 29:
                        return CheckWord.cau29(a, d);
                    case 30:
                        return CheckWord.cau30(a, d);
                    case 31 /*0x1F*/:
                        return CheckWord.cau31(a, d);
                    case 32 /*0x20*/:
                        return CheckWord.cau32(a, d);
                    case 33:
                        return CheckWord.cau33(a, d);
                    case 34:
                        return CheckWord.cau34(a, d);
                    case 35:
                        return CheckWord.cau35(a, d);
                    case 36:
                        return CheckWord.cau36(a, d);
                    case 37:
                        return CheckWord.cau37(a, d);
                    case 38:
                        return CheckWord.cau38(a, d);
                    case 39:
                        return CheckWord.cau39(a, d);
                    case 40:
                        return CheckWord.cau40(a, d);
                    case 41:
                        return CheckWord.cau41(a, d);
                    case 42:
                        return CheckWord.cau42(a, d);
                    case 43:
                        return CheckWord.cau43(a, d);
                    case 44:
                        return CheckWord.cau44(a, d);
                    case 45:
                        return CheckWord.cau45(a, d);
                    case 46:
                        return CheckWord.cau46(a, d);
                    case 47:
                        return CheckWord.cau47(a, d);
                    case 48 /*0x30*/:
                        return CheckWord.cau48(a, d);
                    case 49:
                        return CheckWord.cau49(a, d);
                    case 50:
                        return CheckWord.cau50(a, d);
                    case 51:
                        return CheckWord.cau51(a, d);
                    case 52:
                        return CheckWord.cau52(a, d);
                    case 53:
                        return CheckWord.cau53(a, d);
                    case 54:
                        return CheckWord.cau54(a, d);
                    case 55:
                        return CheckWord.cau55(a, d);
                    case 56:
                        return CheckWord.cau56(a, d);
                    case 57:
                        return CheckWord.cau57(a, d);
                    case 58:
                        return CheckWord.cau58(a, d);
                    case 59:
                        return CheckWord.cau59(a, d);
                    case 60:
                        return CheckWord.cau60(a, d);
                    case 61:
                        return CheckWord.cau61(a, d);
                    case 62:
                        return CheckWord.cau62(a, d);
                    case 63 /*0x3F*/:
                        return CheckWord.cau63(a, d);
                    case 64 /*0x40*/:
                        return CheckWord.cau64(a, d);
                    case 65:
                        return CheckWord.cau65(a, d);
                    case 66:
                        return CheckWord.cau66(a, d);
                    case 67:
                        return CheckWord.cau67(a, d);
                    case 68:
                        return CheckWord.cau68(a, d);
                    case 69:
                        return CheckWord.cau69(a, d);
                    case 70:
                        return CheckWord.cau70(a, d);
                    case 71:
                        return CheckWord.cau71(a, d);
                    case 72:
                        return CheckWord.cau72(a, d);
                    case 73:
                        return CheckWord.cau73(a, d);
                    case 74:
                        return CheckWord.cau74(a, d);
                    case 75:
                        return CheckWord.cau75(a, d);
                    case 76:
                        return CheckWord.cau76(a, d);
                    case 77:
                        return CheckWord.cau77(a, d);
                    case 78:
                        return CheckWord.cau78(a, d);
                    case 79:
                        return CheckWord.cau79(a, d);
                    case 80 /*0x50*/:
                        return CheckWord.cau80(a, d);
                    case 81:
                        return CheckWord.cau81(a, d);
                    case 82:
                        return CheckWord.cau82(a, d);
                    case 83:
                        return CheckWord.cau83(a, d);
                    case 84:
                        return CheckWord.cau84(a, d);
                    case 85:
                        return CheckWord.cau85(a, d);
                    case 86:
                        return CheckWord.cau86(a, d);
                    case 87:
                        return CheckWord.cau87(a, d);
                    case 88:
                        return CheckWord.cau88(a, d);
                    case 89:
                        return CheckWord.cau89(a, d);
                    case 90:
                        return CheckWord.cau90(a, d);
                    case 91:
                        return CheckWord.cau91(a, d);
                    case 92:
                        return CheckWord.cau92(a, d);
                    case 93:
                        return CheckWord.cau93(a, d);
                    case 94:
                        return CheckWord.cau94(a, d);
                    case 95:
                        return CheckWord.cau95(a, d);
                    case 96 /*0x60*/:
                        return CheckWord.cau96(a, d);
                    case 97:
                        return CheckWord.cau97(a, d);
                    case 98:
                        return CheckWord.cau98(a, d);
                    case 99:
                        return CheckWord.cau99(a, d);
                    case 100:
                        return CheckWord.cau100(a, d);
                    case 101:
                        return CheckWord.cau101(a, d);
                    case 102:
                        return CheckWord.cau102(a, d);
                    case 103:
                        return CheckWord.cau103(a, d);
                    case 104:
                        return CheckWord.cau104(a, d);
                    case 105:
                        return CheckWord.cau105(a, d);
                    case 106:
                        return CheckWord.cau106(a, d);
                    case 107:
                        return CheckWord.cau107(a, d);
                    case 108:
                        return CheckWord.cau108(a, d);
                    case 109:
                        return CheckWord.cau109(a, d);
                    case 110:
                        return CheckWord.cau110(a, d);
                    case 111:
                        return CheckWord.cau111(a, d);
                    case 112 /*0x70*/:
                        return CheckWord.cau112(a, d);
                    case 113:
                        return CheckWord.cau113(a, d);
                    case 114:
                        return CheckWord.cau114(a, d);
                    case 115:
                        return CheckWord.cau115(a, d);
                    case 116:
                        return CheckWord.cau116(a, d);
                    case 117:
                        return CheckWord.cau117(a, d);
                    case 118:
                        return CheckWord.cau118(a, d);
                    case 119:
                        return CheckWord.cau119(a, d);
                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static bool cau1(Application a, Presentation d) => true;

        private static bool cau2(Application a, Presentation d) => true;

        private static bool cau3(Application a, Presentation d)
        {
            return a.ActiveWindow.View.Zoom == 92 && a.ActiveWindow.View.Type == PpViewType.ppViewNormal && a.ActiveWindow.Panes[2].Active == MsoTriState.msoTrue;
        }

        private static bool cau4(Application a, Presentation d)
        {
            return a.ActiveWindow.View.Zoom == 66 && a.ActiveWindow.View.Type == PpViewType.ppViewNotesPage;
        }

        private static bool cau5(Application a, Presentation d)
        {
            return a.ActiveWindow.View.Zoom == 66 && a.ActiveWindow.View.Type == PpViewType.ppViewSlideSorter;
        }

        private static bool cau6(Application a, Presentation d)
        {
            return a.Windows.Count == 2 && (double)a.Windows[1].Top == (double)a.Windows[2].Top;
        }

        private static bool cau7(Application a, Presentation d)
        {
            return a.ActiveWindow.BlackAndWhite == MsoTriState.msoTrue;
        }

        private static bool cau8(Application a, Presentation d)
        {
            return a.ActiveWindow.BlackAndWhite == MsoTriState.msoTrue;
        }

        private static bool cau9(Application a, Presentation d)
        {
            return d.Slides.Count == 3 && d.Slides[(object)2].Shapes.Count == 3 && d.Slides[(object)2].Shapes[(object)1].Type == MsoShapeType.msoPlaceholder && d.Slides[(object)2].Shapes[(object)2].Type == MsoShapeType.msoPicture && d.Slides[(object)2].Shapes[(object)3].Type == MsoShapeType.msoAutoShape && d.Slides[(object)2].Shapes[(object)2].Shadow.Visible == MsoTriState.msoTrue && d.Slides[(object)2].Shapes[(object)2].PictureFormat.ColorType == MsoPictureColorType.msoPictureAutomatic;
        }

        private static bool cau10(Application a, Presentation d)
        {
            return d.Slides.Count == 3 && d.Slides[(object)2].Shapes.Count == 2 && d.Slides[(object)2].Shapes[(object)1].Type == MsoShapeType.msoAutoShape && d.Slides[(object)2].Shapes[(object)2].Type == MsoShapeType.msoPicture && d.Slides[(object)2].ColorScheme[PpColorSchemeIndex.ppAccent1].RGB == 5924977;
        }

        private static bool cau11(Application a, Presentation d)
        {
            return d.Slides.Count == 2 && d.Slides[(object)2].Shapes.Count == 4 && d.Slides[(object)2].Shapes[(object)1].Type == MsoShapeType.msoPicture && d.Slides[(object)2].Shapes[(object)2].Type == MsoShapeType.msoAutoShape && d.Slides[(object)2].Shapes[(object)3].Type == MsoShapeType.msoAutoShape && d.Slides[(object)2].Shapes[(object)4].Type == MsoShapeType.msoAutoShape && d.Slides[(object)2].ColorScheme[PpColorSchemeIndex.ppAccent1].RGB == 6786989;
        }

        private static bool cau12(Application a, Presentation d)
        {
            return d.Slides.Count == 3 && d.Slides[(object)2].Shapes.Count == 3 && d.Slides[(object)2].Shapes[(object)1].Type == MsoShapeType.msoPlaceholder && d.Slides[(object)2].Shapes[(object)2].Type == MsoShapeType.msoPicture && d.Slides[(object)2].Shapes[(object)3].Type == MsoShapeType.msoAutoShape && d.Slides[(object)2].Shapes[(object)2].Shadow.Visible == MsoTriState.msoTrue && d.Slides[(object)2].Shapes[(object)2].PictureFormat.ColorType == MsoPictureColorType.msoPictureAutomatic;
        }

        private static bool cau13(Application a, Presentation d)
        {
            return a.ActivePresentation.Slides.Count == 2 && a.ActivePresentation.Slides[(object)2].Shapes.Count == 3 && a.ActivePresentation.Slides[(object)2].Shapes[(object)1].Type == MsoShapeType.msoPlaceholder && a.ActivePresentation.Slides[(object)2].Shapes[(object)2].Type == MsoShapeType.msoPicture && a.ActivePresentation.Slides[(object)2].Shapes[(object)3].Type == MsoShapeType.msoPicture && a.ActivePresentation.Slides[(object)2].Shapes[(object)2].PictureFormat.ColorType == MsoPictureColorType.msoPictureGrayscale;
        }

        private static bool cau14(Application a, Presentation d)
        {
            return (int)d.Slides[(object)3].Shapes[(object)2].TextFrame.TextRange.BoundLeft == 94 && (int)d.Slides[(object)3].Shapes[(object)2].TextFrame.TextRange.BoundHeight == 339;
        }

        private static bool cau15(Application a, Presentation d)
        {
            return (int)d.Slides[(object)3].Shapes[(object)2].TextFrame.TextRange.BoundLeft == 128 /*0x80*/ && (int)d.Slides[(object)3].Shapes[(object)2].TextFrame.TextRange.BoundHeight == 307;
        }

        private static bool cau16(Application a, Presentation d)
        {
            return (int)d.Slides[(object)5].Shapes[(object)3].TextFrame.TextRange.BoundTop == 278;
        }

        private static bool cau17(Application a, Presentation d)
        {
            return d.Slides[(object)3].Shapes[(object)2].TextFrame2.Column.Number == 2;
        }

        private static bool cau18(Application a, Presentation d)
        {
            return d.Slides[(object)3].Shapes[(object)2].TextFrame2.Column.Number == 1;
        }

        private static bool cau19(Application a, Presentation d)
        {
            if (d.Slides.Count != 9)
                return false;
            for (int index = 1; index <= d.Slides.Count; ++index)
            {
                if (d.Slides[(object)1].Layout == PpSlideLayout.ppLayoutSectionHeader)
                    return false;
            }
            return true;
        }

        private static bool cau20(Application a, Presentation d)
        {
            return d.Slides[(object)1].ColorScheme[PpColorSchemeIndex.ppAccent1].RGB == 5924977 && !(d.Slides[(object)1].Shapes[(object)1].TextFrame.TextRange.Font.Name != "Lucida Sans");
        }

        private static bool cau21(Application a, Presentation d)
        {
            return d.Slides[(object)1].ColorScheme[PpColorSchemeIndex.ppAccent1].RGB == 5660786 && !(d.Slides[(object)1].Shapes[(object)1].TextFrame.TextRange.Font.Name != "Arial");
        }

        private static bool cau22(Application a, Presentation d)
        {
            return (double)d.PageSetup.SlideHeight == 619.25 && (double)d.PageSetup.SlideWidth == 792.0;
        }

        private static bool cau23(Application a, Presentation d)
        {
            return (double)d.PageSetup.SlideHeight == 648.0 && (double)d.PageSetup.SlideWidth == 900.0;
        }

        private static bool cau24(Application a, Presentation d)
        {
            return d.Slides.Count == 5 && d.Slides[(object)1].Shapes.Count == 2 && d.Slides[(object)2].Shapes.Count == 3 && !(d.Slides[(object)2].Shapes[(object)3].TextFrame.TextRange.Text != "whales sizes and Ages");
        }

        private static bool cau25(Application a, Presentation d)
        {
            return d.HandoutMaster.Shapes[(object)2].TextFrame.TextRange.Text.Length >= 11 && d.HandoutMaster.Shapes[(object)2].TextFrame.TextRange.Text.Length <= 18 && !(d.HandoutMaster.Shapes[(object)3].TextFrame.TextRange.Text != "Whales");
        }

        private static bool cau26(Application a, Presentation d)
        {
            return d.HandoutMaster.Shapes[(object)2].TextFrame.TextRange.Text.Length >= 20 && !(d.HandoutMaster.Shapes[(object)3].TextFrame.TextRange.Text != "Contoso");
        }

        private static bool cau27(Application a, Presentation d)
        {
            return d.Slides[(object)4].Shapes[(object)1].Line.ForeColor.RGB == 13056;
        }

        private static bool cau28(Application a, Presentation d)
        {
            return d.Slides[(object)4].Shapes[(object)1].Line.ForeColor.RGB == 13004559;
        }

        private static bool cau29(Application a, Presentation d)
        {
            return d.Slides[(object)4].Shapes[(object)1].Line.ForeColor.RGB == 14261504;
        }

        private static bool cau30(Application a, Presentation d)
        {
            return d.Slides[(object)4].Shapes[(object)1].Line.ForeColor.RGB == 14274571;
        }

        private static bool cau31(Application a, Presentation d)
        {
            return d.Slides[(object)3].Shapes.Count == 3 && (double)d.Slides[(object)3].Shapes[(object)3].Shadow.Blur == 6.0 && (double)d.Slides[(object)3].Shapes[(object)3].Shadow.Size == 100.0 && (double)d.Slides[(object)3].Shapes[(object)3].Shadow.Transparency == 0.60000002384185791;
        }

        private static bool cau32(Application a, Presentation d)
        {
            return d.Slides[(object)3].Shapes.Count == 3 && (double)d.Slides[(object)3].Shapes[(object)3].Shadow.Blur == 5.118110179901123 && (double)d.Slides[(object)3].Shapes[(object)3].Shadow.Size == 100.0 && (double)d.Slides[(object)3].Shapes[(object)3].Shadow.Transparency == 0.699999988079071;
        }

        private static bool cau33(Application a, Presentation d)
        {
            return d.Slides[(object)3].Shapes.Count == 3 && (double)d.Slides[(object)3].Shapes[(object)3].Shadow.Blur != 5.118110179901123 && (double)d.Slides[(object)3].Shapes[(object)3].Shadow.Transparency != 0.699999988079071 && (double)d.Slides[(object)3].Shapes[(object)3].PictureFormat.Contrast == 0.5 && (double)d.Slides[(object)3].Shapes[(object)3].PictureFormat.Brightness == 0.5;
        }

        private static bool cau34(Application a, Presentation d)
        {
            return d.Slides[(object)3].Shapes.Count == 3 && (double)d.Slides[(object)3].Shapes[(object)3].Shadow.Blur != 5.118110179901123 && (double)d.Slides[(object)3].Shapes[(object)3].Shadow.Transparency != 0.699999988079071 && (double)d.Slides[(object)3].Shapes[(object)3].PictureFormat.Contrast == 0.5 && (double)d.Slides[(object)3].Shapes[(object)3].PictureFormat.Brightness == 0.5;
        }

        private static bool cau35(Application a, Presentation d)
        {
            return d.Slides[(object)4].Shapes.Count == 3 && !(d.Slides[(object)4].Shapes[(object)1].Name == "Title 1");
        }

        private static bool cau36(Application a, Presentation d)
        {
            return d.Slides[(object)4].Shapes.Count == 3 && !(d.Slides[(object)4].Shapes[(object)1].Name == "Title 1");
        }

        private static bool cau37(Application a, Presentation d)
        {
            return d.Slides[(object)4].Shapes.Count == 3 && !(d.Slides[(object)4].Shapes[(object)1].Name == "Title 1");
        }

        private static bool cau38(Application a, Presentation d)
        {
            return d.Slides.Count == 6 && d.Slides[(object)6].Shapes.Count == 5 && d.Slides[(object)6].Shapes[(object)5].GroupItems.Count == 2 && !(d.Slides[(object)6].Shapes[(object)5].GroupItems[(object)2].TextFrame.TextRange.Text != "Vocal chords");
        }

        private static bool cau39(Application a, Presentation d)
        {
            return d.Slides.Count == 6 && d.Slides[(object)6].Shapes.Count == 5 && d.Slides[(object)6].Shapes[(object)4].GroupItems.Count == 1 && !(d.Slides[(object)6].Shapes[(object)4].GroupItems[(object)1].TextFrame.TextRange.Text != "All year");
        }

        private static bool cau40(Application a, Presentation d)
        {
            return d.Slides.Count == 6 && d.Slides[(object)6].Shapes.Count == 5 && d.Slides[(object)6].Shapes[(object)5].GroupItems.Count == 4 && (double)d.Slides[(object)6].Shapes[(object)5].GroupItems[(object)1].Top == 192.0 && (double)d.Slides[(object)6].Shapes[(object)5].GroupItems[(object)1].Left == 113.34370422363281 && (double)d.Slides[(object)6].Shapes[(object)5].GroupItems[(object)1].Width == 151.43748474121094 && (double)d.Slides[(object)6].Shapes[(object)5].GroupItems[(object)1].Height == 151.43748474121094;
        }

        private static bool cau41(Application a, Presentation d)
        {
            return d.Slides.Count == 6 && d.Slides[(object)6].Shapes.Count == 5 && d.Slides[(object)6].Shapes[(object)5].GroupItems.Count == 6 && (double)d.Slides[(object)6].Shapes[(object)5].GroupItems[(object)1].Top == 316.77197265625 && (double)d.Slides[(object)6].Shapes[(object)5].GroupItems[(object)1].Left == 30.01960563659668 && (double)d.Slides[(object)6].Shapes[(object)5].GroupItems[(object)1].Width == 86.935356140136719 && (double)d.Slides[(object)6].Shapes[(object)5].GroupItems[(object)1].Height == 86.935356140136719;
        }

        private static bool cau42(Application a, Presentation d)
        {
            return d.Slides.Count == 6 && d.Slides[(object)4].Shapes.Count == 3 && d.Slides[(object)4].Shapes[(object)3].MediaType == PpMediaType.ppMediaTypeSound && d.Slides[(object)4].Shapes[(object)3].AnimationSettings.PlaySettings.StopAfterSlides == 999;
        }

        private static bool cau43(Application a, Presentation d)
        {
            return d.Slides.Count == 6 && d.Slides[(object)6].Shapes.Count == 2 && d.Slides[(object)6].Shapes[(object)2].MediaType == PpMediaType.ppMediaTypeSound && d.Slides[(object)6].Shapes[(object)2].AnimationSettings.AdvanceMode == PpAdvanceMode.ppAdvanceOnTime;
        }

        private static bool cau44(Application a, Presentation d)
        {
            return a.DisplayGridLines == MsoTriState.msoTrue;
        }

        private static bool cau45(Application a, Presentation d)
        {
            return d.Slides.Count == 7 && d.Slides[(object)3].Shapes.Count == 2;
        }

        private static bool cau46(Application a, Presentation d)
        {
            return d.Slides.Count == 7 && d.Slides[(object)3].Shapes.Count == 2;
        }

        private static bool cau47(Application a, Presentation d)
        {
            return d.Slides.Count == 7 && d.Slides[(object)3].Shapes.Count == 2;
        }

        private static bool cau48(Application a, Presentation d)
        {
            return d.Slides.Count == 7 && d.Slides[(object)3].Shapes.Count == 2;
        }

        private static bool cau49(Application a, Presentation d)
        {
            return d.Slides.Count == 7 && d.Slides[(object)3].Shapes.Count == 2;
        }

        private static bool cau50(Application a, Presentation d)
        {
            return d.Slides.Count == 7 && d.Slides[(object)3].Shapes.Count == 2;
        }

        private static bool cau51(Application a, Presentation d)
        {
            return d.Slides.Count == 7 && d.Slides[(object)3].Shapes.Count == 2;
        }

        private static bool cau52(Application a, Presentation d)
        {
            return d.Slides.Count == 7 && d.Slides[(object)3].Shapes.Count == 2;
        }

        private static bool cau53(Application a, Presentation d)
        {
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            return d.Slides.Count == 7 && d.Slides[(object)3].Shapes.Count == 2 && d.Slides[(object)3].Shapes[(object)2].Table.Columns.Count == 3 && d.Slides[(object)3].Shapes[(object)2].Table.Rows.Count == 3 && !(d.Slides[(object)3].Shapes[(object)2].Table.Cell(1, 1).Shape.TextFrame.TextRange.Text != "Whales") && !(d.Slides[(object)3].Shapes[(object)2].Table.Cell(1, 2).Shape.TextFrame.TextRange.Text != "1990") && !(d.Slides[(object)3].Shapes[(object)2].Table.Cell(1, 3).Shape.TextFrame.TextRange.Text != "2000");
        }

        private static bool cau54(Application a, Presentation d)
        {
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            return d.Slides.Count == 7 && d.Slides[(object)3].Shapes.Count == 2 && d.Slides[(object)3].Shapes[(object)2].Table.Columns.Count == 2 && d.Slides[(object)3].Shapes[(object)2].Table.Rows.Count == 3 && !(d.Slides[(object)3].Shapes[(object)2].Table.Cell(1, 1).Shape.TextFrame.TextRange.Text != "Whales") && !(d.Slides[(object)3].Shapes[(object)2].Table.Cell(1, 2).Shape.TextFrame.TextRange.Text != "2000");
        }

        private static bool cau55(Application a, Presentation d)
        {
            return d.Slides.Count == 5 && !(d.Slides[(object)3].SlideShowTransition.SoundEffect.Name != "breeze.wav") && !(d.Slides[(object)5].SlideShowTransition.SoundEffect.Name != "breeze.wav") && !(d.Slides[(object)1].SlideShowTransition.SoundEffect.Name != "arrow.wav") && !(d.Slides[(object)2].SlideShowTransition.SoundEffect.Name != "arrow.wav") && !(d.Slides[(object)4].SlideShowTransition.SoundEffect.Name != "arrow.wav");
        }

        private static bool cau56(Application a, Presentation d)
        {
            return d.Slides.Count == 5 && !(d.Slides[(object)3].SlideShowTransition.SoundEffect.Name != "chimes.wav") && !(d.Slides[(object)5].SlideShowTransition.SoundEffect.Name != "chimes.wav") && !(d.Slides[(object)1].SlideShowTransition.SoundEffect.Name != "arrow.wav") && !(d.Slides[(object)2].SlideShowTransition.SoundEffect.Name != "arrow.wav") && !(d.Slides[(object)4].SlideShowTransition.SoundEffect.Name != "arrow.wav");
        }

        private static bool cau57(Application a, Presentation d)
        {
            if (d.Slides.Count != 5)
                return false;
            for (int Index = 1; Index <= 5; ++Index)
            {
                if (d.Slides[(object)Index].SlideShowTransition.AdvanceOnTime != MsoTriState.msoTrue || d.Slides[(object)Index].SlideShowTransition.AdvanceOnClick != MsoTriState.msoFalse || (double)d.Slides[(object)Index].SlideShowTransition.AdvanceTime != 15.0)
                    return false;
            }
            return true;
        }

        private static bool cau58(Application a, Presentation d)
        {
            if (d.Slides.Count != 5)
                return false;
            for (int Index = 1; Index <= 5; ++Index)
            {
                if (d.Slides[(object)Index].SlideShowTransition.AdvanceOnTime != MsoTriState.msoTrue || d.Slides[(object)Index].SlideShowTransition.AdvanceOnClick != MsoTriState.msoFalse || (double)d.Slides[(object)Index].SlideShowTransition.AdvanceTime != 20.0)
                    return false;
            }
            return true;
        }

        private static bool cau59(Application a, Presentation d)
        {
            return d.Slides.Count == 5 && d.Slides[(object)4].Shapes.Count == 2;
        }

        private static bool cau60(Application a, Presentation d) => true;

        private static bool cau61(Application a, Presentation d) => true;

        private static bool cau62(Application a, Presentation d) => true;

        private static bool cau63(Application a, Presentation d) => true;

        private static bool cau64(Application a, Presentation d)
        {
            return d.Slides[(object)1].Comments.Count == 1 && !(d.Slides[(object)1].Comments[1].Text != "Good picture");
        }

        private static bool cau65(Application a, Presentation d)
        {
            return d.Slides.Count == 5 && d.Slides[(object)4].Comments.Count == 1 && !(d.Slides[(object)4].Comments[1].Text != "Review");
        }

        private static bool cau66(Application a, Presentation d)
        {
            return d.Slides.Count == 5 && d.Slides[(object)3].Comments.Count == 0 && d.Slides[(object)5].Comments.Count == 1;
        }

        private static bool cau67(Application a, Presentation d)
        {
            return d.Slides.Count == 5 && d.Slides[(object)3].Comments.Count == 0 && d.Slides[(object)5].Comments.Count == 1;
        }

        private static bool cau68(Application a, Presentation d) => true;

        private static bool cau69(Application a, Presentation d) => true;

        private static bool cau70(Application a, Presentation d) => true;

        private static bool cau71(Application a, Presentation d) => true;

        private static bool cau72(Application a, Presentation d)
        {
            return File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "TreePicture.ppt"));
        }

        private static bool cau73(Application a, Presentation d)
        {
            return File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "shawes.ppsx"));
        }

        private static bool cau74(Application a, Presentation d) => true;

        private static bool cau75(Application a, Presentation d) => true;

        private static bool cau76(Application a, Presentation d) => true;

        private static bool cau77(Application a, Presentation d)
        {
            return d.SlideShowSettings.ShowType != PpSlideShowType.ppShowTypeKiosk && d.SlideShowSettings.ShowType != PpSlideShowType.ppShowTypeWindow && d.SlideShowSettings.ShowType != PpSlideShowType.ppShowTypeSpeaker;
        }

        private static bool cau78(Application a, Presentation d)
        {
            return d.SlideShowSettings.ShowType == PpSlideShowType.ppShowTypeKiosk;
        }

        private static bool cau79(Application a, Presentation d)
        {
            return d.SlideShowSettings.NamedSlideShows.Count == 1 && !(d.SlideShowSettings.NamedSlideShows[(object)1].Name != "Whale Graphs") && d.SlideShowSettings.NamedSlideShows[(object)1].Count == 4;
        }

        private static bool cau80(Application a, Presentation d)
        {
            return d.SlideShowSettings.NamedSlideShows.Count == 1 && !(d.SlideShowSettings.NamedSlideShows[(object)1].Name != "Whale Graphs") && d.SlideShowSettings.NamedSlideShows[(object)1].Count == 3;
        }

        private static bool cau81(Application a, Presentation d) => true;

        private static bool cau82(Application a, Presentation d) => true;

        private static bool cau83(Application a, Presentation d) => true;

        private static bool cau84(Application a, Presentation d) => true;

        private static bool cau85(Application a, Presentation d) => true;

        private static bool cau86(Application a, Presentation d) => true;

        private static bool cau87(Application a, Presentation d) => true;

        private static bool cau88(Application a, Presentation d) => true;

        private static bool cau89(Application a, Presentation d) => true;

        private static bool cau90(Application a, Presentation d) => true;

        private static bool cau91(Application a, Presentation d) => true;

        private static bool cau92(Application a, Presentation d) => true;

        private static bool cau93(Application a, Presentation d) => true;

        private static bool cau94(Application a, Presentation d) => true;

        private static bool cau95(Application a, Presentation d) => true;

        private static bool cau96(Application a, Presentation d) => true;

        private static bool cau97(Application a, Presentation d) => true;

        private static bool cau98(Application a, Presentation d) => true;

        private static bool cau99(Application a, Presentation d) => true;

        private static bool cau100(Application a, Presentation d) => true;

        private static bool cau101(Application a, Presentation d) => true;

        private static bool cau102(Application a, Presentation d) => true;

        private static bool cau103(Application a, Presentation d) => true;

        private static bool cau104(Application a, Presentation d) => true;

        private static bool cau105(Application a, Presentation d) => true;

        private static bool cau106(Application a, Presentation d) => true;

        private static bool cau107(Application a, Presentation d) => true;

        private static bool cau108(Application a, Presentation d) => true;

        private static bool cau109(Application a, Presentation d) => true;

        private static bool cau110(Application a, Presentation d) => true;

        private static bool cau111(Application a, Presentation d) => true;

        private static bool cau112(Application a, Presentation d) => true;

        private static bool cau113(Application a, Presentation d) => true;

        private static bool cau114(Application a, Presentation d) => true;

        private static bool cau115(Application a, Presentation d) => true;

        private static bool cau116(Application a, Presentation d) => true;

        private static bool cau117(Application a, Presentation d) => true;

        private static bool cau118(Application a, Presentation d) => true;

        private static bool cau119(Application a, Presentation d) => true;

        private static bool cau120(Application a, Presentation d) => true;

        private static bool cau121(Application a, Presentation d) => true;

        private static bool cau122(Application a, Presentation d) => true;

        private static bool cau123(Application a, Presentation d) => true;

        private static bool cau124(Application a, Presentation d) => true;

        private static bool cau125(Application a, Presentation d) => true;

        private static bool cau126(Application a, Presentation d) => true;

        private static bool cau127(Application a, Presentation d) => true;

        private static bool cau128(Application a, Presentation d) => true;

        private static bool cau129(Application a, Presentation d) => true;

        private static bool cau130(Application a, Presentation d) => true;

        private static bool cau131(Application a, Presentation d) => true;

        private static bool cau132(Application a, Presentation d) => true;

        private static bool cau133(Application a, Presentation d) => true;

        private static bool cau134(Application a, Presentation d) => true;

        private static bool cau135(Application a, Presentation d) => true;

        private static bool cau136(Application a, Presentation d) => true;

        private static bool cau137(Application a, Presentation d) => true;

        private static bool cau138(Application a, Presentation d) => true;

        private static bool cau139(Application a, Presentation d) => true;

        private static bool cau140(Application a, Presentation d) => true;

        private static bool cau141(Application a, Presentation d) => true;

        private static bool cau142(Application a, Presentation d) => true;

        private static bool cau143(Application a, Presentation d) => true;

        private static bool cau144(Application a, Presentation d) => true;

        private static bool cau145(Application a, Presentation d) => true;

        private static bool cau146(Application a, Presentation d) => true;

        private static bool cau147(Application a, Presentation d) => true;

        private static bool cau148(Application a, Presentation d) => true;

        private static bool cau149(Application a, Presentation d) => true;

        private static bool cau150(Application a, Presentation d) => true;
    }
}