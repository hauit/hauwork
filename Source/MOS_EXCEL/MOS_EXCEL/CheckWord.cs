// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.CheckWord
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.CompilerServices;

namespace MOS_EXCEL_LEARN
{
    internal class CheckWord
    {
        public static bool CheckCau(int causo, _Application a, _Workbook d)
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
                case 16:
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
                case 31:
                    return CheckWord.cau31(a, d);
                case 32:
                    return CheckWord.cau32(a, d);
                case 33:
                    return CheckWord.cau33(a, d);
                case 34:
                    return CheckWord.cau34(a, d);
                case 35:
                    return CheckWord.cau35(a, d);
                case 36:
                    return CheckWord.cau36(a, d);
                default:
                    return false;
            }
        }

        private static bool cau1(_Application a, _Workbook d)
        {
            Worksheet worksheet = null;

            // Tìm sheet có tên là "Sales"
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Sales")
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet == null)
                return false;

            try
            {
                // Đọc giá trị từ các ô cụ thể và kiểm tra nội dung
                string c4 = (string)worksheet.get_Range("C4", "C4").Text;
                if (c4 != "50121") return false;

                string c9 = (string)worksheet.get_Range("C9", "C9").Text;
                if (c9 != "50123") return false;

                string c14 = (string)worksheet.get_Range("C14", "C14").Text;
                if (c14 != "50168") return false;

                string c27 = (string)worksheet.get_Range("C27", "C27").Text;
                if (c27 != "50335") return false;

                string e13 = (string)worksheet.get_Range("E13", "E13").Text;
                if (e13 != "3210") return false;

                string e20 = (string)worksheet.get_Range("E20", "E20").Text;
                if (e20 != "6464") return false;
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private static bool cau2(_Application a, _Workbook d)
        {
            try
            {
                // Kiểm tra: Zoom của cửa sổ hiện tại phải là 50%
                if ((double)a.ActiveWindow.Zoom != 50.0)
                    return false;

                // Kiểm tra: Chế độ xem phải là Normal View
                if (a.ActiveWindow.View != XlWindowView.xlNormalView)
                    return false;

                // Kiểm tra: Trang tính đang active phải là "Sales"
                Worksheet worksheet = (Worksheet)d.ActiveSheet;
                if (worksheet.Name != "Sales")
                    return false;
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private static bool cau3(_Application a, _Workbook d)
        {
            try
            {
                // Kiểm tra số lượng Custom View phải là 1
                if (a.ActiveWorkbook.CustomViews.Count != 1)
                    return false;

                // Lấy Custom View đầu tiên (vị trí 1)
                CustomView customView = a.ActiveWorkbook.CustomViews.Item(1);

                // Tên phải là "Preview"
                if (customView.Name != "Preview")
                    return false;

                // Áp dụng custom view đó
                customView.Show();

                // Kiểm tra các thiết lập phải được bật
                if (!customView.PrintSettings || !customView.RowColSettings)
                    return false;

                // Zoom phải là 50%
                if ((double)a.ActiveWindow.Zoom != 50.0)
                    return false;

                // View phải là Normal
                if (a.ActiveWindow.View != XlWindowView.xlNormalView)
                    return false;
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private static bool cau4(_Application a, _Workbook d)
        {
            Worksheet worksheet = null;

            // Tìm sheet có tên "Sales"
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Sales")
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet == null)
                return false;

            // Kiểm tra số lượng shape trên sheet phải là 2
            if (worksheet.Shapes.Count != 2)
                return false;

            Microsoft.Office.Interop.Excel.Shape shape1;

            try
            {
                // Lấy shape có tên "Diagram 3"
                shape1 = worksheet.Shapes.Item("Diagram 3");
            }
            catch (Exception)
            {
                return false;
            }

            // Kiểm tra shape1 có đúng 6 item trong nhóm không
            if (shape1.GroupItems.Count != 6)
                return false;

            // Lấy shape con thứ 3 trong nhóm
            Microsoft.Office.Interop.Excel.Shape shape2 = shape1.GroupItems.Item(3);

            // Kiểm tra style, màu, góc xoay, ánh sáng 3D, độ sâu và góc nhìn 3D
            return
                shape2.ShapeStyle == MsoShapeStyleIndex.msoShapeStyleMixed &&
                shape2.Fill.ForeColor.RGB == 9291697 &&
                (double)shape2.Rotation == 0.0 &&
                (double)shape2.ThreeD.LightAngle == 20.0 &&
                (double)shape2.ThreeD.Depth == 0.0 &&
                (double)shape2.ThreeD.FieldOfView == 45.0;
        }


        private static bool cau5(_Application a, _Workbook d)
        {
            foreach (Worksheet worksheet in d.Worksheets)
            {
                try
                {
                    // Thử truy cập màu tab của sheet
                    int color = (int)worksheet.Tab.Color;
                }
                catch (Exception)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool cau6(_Application a, _Workbook d)
        {
            // ISSUE: variable of a compiler-generated type
            Worksheet worksheet = (Worksheet)null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Summary")
                    worksheet = sheet;
            }
            if (worksheet == null)
                return false;
            try
            {
                // ISSUE: reference to a compiler-generated method
                // ISSUE: reference to a compiler-generated method
                string str = worksheet.get_Range((object)"c7", (object)"c7").Comment.Text((object)"", (object)1, (object)false);
                if (str.Contains("Review"))
                    return false;
                if (!str.Contains("Done"))
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau7(_Application a, _Workbook d)
        {
            Worksheet worksheet = null;

            // Tìm sheet có tên "Partner Orders"
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Partner Orders")
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet == null)
                return false;

            try
            {
                // Kiểm tra giá trị tại các ô cụ thể
                if ((string)worksheet.get_Range("D4", "D4").Text != "2762") return false;
                if ((string)worksheet.get_Range("D9", "D9").Text != "2762") return false;
                if ((string)worksheet.get_Range("D16", "D16").Text != "2762") return false;
                if ((string)worksheet.get_Range("D32", "D32").Text != "2762") return false;

                if ((string)worksheet.get_Range("F31", "F31").Text != "2765") return false;
                if ((string)worksheet.get_Range("F33", "F33").Text != "2765") return false;
                if ((string)worksheet.get_Range("F38", "F38").Text != "2765") return false;
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private static bool cau8(_Application a, _Workbook d)
        {
            // ISSUE: variable of a compiler-generated type
            Worksheet worksheet = (Worksheet)null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Partner Orders")
                    worksheet = sheet;
            }
            if (worksheet == null)
                return false;
            try
            {
                if (worksheet.PageSetup.PrintTitleRows != "$3:$3")
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau9(_Application a, _Workbook d)
        {
            try
            {
                if (!d.RemovePersonalInformation)
                    return false;
                foreach (_Worksheet sheet in d.Sheets)
                {
                    if (sheet.Comments.Count != 0)
                        return false;
                }
                if (d.Title != "" || d.Author != "")
                    return false;
                if (d.Keywords != "")
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau10(_Application a, _Workbook d)
        {
            try
            {
                // ISSUE: variable of a compiler-generated type
                Worksheet worksheet = (Worksheet)null;
                foreach (Worksheet sheet in d.Sheets)
                {
                    if (sheet.Name == "Summary")
                        worksheet = sheet;
                }
                // ISSUE: reference to a compiler-generated method
                // ISSUE: reference to a compiler-generated method
                // ISSUE: reference to a compiler-generated method
                if (worksheet == null || (int)worksheet.get_Range((object)"b17", (object)"b21").Cells.HorizontalAlignment != -4131 || (int)worksheet.get_Range((object)"b17", (object)"b21").Cells.VerticalAlignment != -4108 || (string)worksheet.get_Range((object)"b17", (object)"b17").Cells.Text != "Northwind Traders")
                    return false;
                // ISSUE: reference to a compiler-generated method
                if ((string)worksheet.get_Range((object)"b21", (object)"b21").Cells.Text != "Fabrikam, Inc")
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau11(_Application a, _Workbook d)
        {
            Worksheet worksheet = null;

            // Tìm worksheet có tên "Partner Orders"
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Partner Orders")
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet == null)
                return false;

            try
            {
                // Kiểm tra công thức ở ô G4
                if ((string)worksheet.get_Range("G4", "G4").Formula != "=E4*F4*$G$2")
                    return false;

                // Kiểm tra công thức ở ô G43
                if ((string)worksheet.get_Range("G43", "G43").Formula != "=E43*F43*$G$2")
                    return false;
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private static bool cau12(_Application a, _Workbook d)
        {
            Worksheet worksheet = null;

            // Tìm sheet có tên "Sales by Region"
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Sales by Region")
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet == null)
                return false;

            try
            {
                // Lấy range từ B2 đến I2
                Range range = worksheet.get_Range("B2", "I2");

                // Kiểm tra xem các ô trong range đã được merge chưa
                if (!(bool)range.MergeCells)
                    return false;

                // Kiểm tra căn lề ngang (HorizontalAlignment)
                // -4131 tương ứng với xlCenter
                if ((int)range.HorizontalAlignment != -4131) // -4131 = XlHAlign.xlHAlignCenter
                    return false;
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private static bool cau13(_Application a, _Workbook d)
        {
            try
            {
                // ISSUE: variable of a compiler-generated type
                Worksheet worksheet = (Worksheet)null;
                foreach (Worksheet sheet in d.Sheets)
                {
                    if (sheet.Name == "Sales by Region")
                        worksheet = sheet;
                }
                if (worksheet == null || worksheet.Cells.SparklineGroups.Count != 1)
                    return false;
                if (worksheet.Cells.SparklineGroups[(object)1].Type != XlSparkType.xlSparkColumn)
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau14(_Application a, _Workbook d)
        {
            try
            {
                Worksheet worksheet = null;

                // Tìm sheet có tên "Summary"
                foreach (Worksheet sheet in d.Sheets)
                {
                    if (sheet.Name == "Summary")
                    {
                        worksheet = sheet;
                        break;
                    }
                }

                if (worksheet == null)
                    return false;

                // Lấy nội dung ô C6
                string textC6 = (string)worksheet.get_Range("C6", "C6").Text;
                if (textC6 != "         808,501,235 ")
                    return false;

                // Lấy công thức trong ô C6
                string formulaC6 = (string)worksheet.get_Range("C6", "C6").Formula;
                if (formulaC6 != "=SUM('2006'!F9,'2007'!F9,'2008'!F9,'2009'!F9)")
                    return false;
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private static bool cau15(_Application a, _Workbook d)
        {
            bool hasFontCondition = false;
            bool hasInteriorCondition = false;

            Worksheet worksheet = null;

            // Tìm sheet có tên "2006"
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "2006")
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet == null)
                return false;

            // Duyệt từng FormatCondition trong vùng E4:E8
            foreach (FormatCondition formatCondition in worksheet.get_Range("E4", "E8").FormatConditions)
            {
                // Kiểm tra công thức và kiểu điều kiện
                if (formatCondition.Formula1 != "=100" || formatCondition.Type != 1)
                    return false;

                if (formatCondition.Operator != 5 && formatCondition.Operator != 6)
                    return false;

                try
                {
                    if (formatCondition.Operator == 6) // Format cho giá trị lớn hơn
                    {
                        double interiorColor = Convert.ToDouble(formatCondition.Interior.Color);
                        if (interiorColor != 13551615.0)
                            return false;

                        hasInteriorCondition = true;
                    }

                    if (formatCondition.Operator == 5) // Format cho giá trị nhỏ hơn
                    {
                        double fontColor = Convert.ToDouble(formatCondition.Font.Color);
                        if (fontColor != 26012.0)
                            return false;

                        double fontTintAndShade = Convert.ToDouble(formatCondition.Font.TintAndShade);
                        if (fontTintAndShade != 0.0)
                            return false;

                        double interiorColor = Convert.ToDouble(formatCondition.Interior.Color);
                        if (interiorColor != 10284031.0)
                            return false;

                        double interiorTintAndShade = Convert.ToDouble(formatCondition.Interior.TintAndShade);
                        if (interiorTintAndShade != 0.0)
                            return false;

                        hasFontCondition = true;
                    }
                }
                catch (Exception)
                {
                    return false;
                }
            }

            return hasFontCondition && hasInteriorCondition;
        }

        private static bool cau16(_Application a, _Workbook d)
        {
            try
            {
                if (!a.ActiveWindow.Split || a.ActiveWindow.SplitColumn == 0)
                    return false;
                if (a.ActiveWindow.SplitRow != 0)
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau17(_Application a, _Workbook d)
        {
            Worksheet worksheet = null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Team_ADatum")
                    worksheet = sheet;
            }
            if (worksheet == null)
                return false;

            // Kiểm tra 3 hàng: C9:H9, C15:H15, C21:H21
            string[] ranges = { "C9:H9", "C15:H15", "C21:H21" };

            foreach (string range in ranges)
            {
                FormatConditions formatConditions = worksheet.get_Range(range).FormatConditions;
                if (formatConditions.Count != 1)
                    return false;

                for (int i = 1; i <= formatConditions.Count; i++)
                {
                    dynamic databar;
                    try
                    {
                        databar = formatConditions.Item(i);
                    }
                    catch
                    {
                        return false;
                    }

                    if (databar.Type != 4) // 4 = xlDataBar
                        return false;

                    try
                    {
                        int color = (int)databar.BarColor.Color;
                        int percentMin = (int)databar.PercentMin;
                        int percentMax = (int)databar.PercentMax;

                        if (color != 2668287 || percentMin != 0 || percentMax != 100)
                            return false;
                    }
                    catch
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool cau18(_Application a, _Workbook d)
        {
            try
            {
                Worksheet worksheet = null;
                foreach (Worksheet sheet in d.Sheets)
                {
                    if (sheet.Name == "Sales by Region")
                        worksheet = sheet;
                }
                if (worksheet == null)
                    return false;

                Range range1 = worksheet.get_Range("D6", "I9");
                if ((int)range1.HorizontalAlignment != -4108) // xlCenter
                    return false;
                if ((int)range1.VerticalAlignment != -4107) // xlTop
                    return false;

                Borders borders1 = range1.Borders;
                if (borders1.Count != 6)
                    return false;
                if ((double)(borders1.Color) != 5880731.0)
                    return false;

                if ((double)worksheet.get_Range("D6", "I6").Interior.Color != 14610923.0)
                    return false;
                if ((double)worksheet.get_Range("D9", "I9").Interior.Color != 16777215.0)
                    return false;

                Range range2 = worksheet.get_Range("D12", "I15");
                if ((int)range2.HorizontalAlignment != -4108)
                    return false;
                if ((int)range2.VerticalAlignment != -4107)
                    return false;

                Borders borders2 = range2.Borders;
                if (borders2.Count != 6)
                    return false;
                if ((double)(borders2.Color) != 5880731.0)
                    return false;

                if ((double)worksheet.get_Range("D12", "I12").Interior.Color != 14610923.0)
                    return false;
                if ((double)worksheet.get_Range("D15", "I15").Interior.Color != 16777215.0)
                    return false;

                Range range3 = worksheet.get_Range("D18", "I21");
                if ((int)range3.HorizontalAlignment != -4108)
                    return false;
                if ((int)range3.VerticalAlignment != -4107)
                    return false;

                Borders borders3 = range3.Borders;
                if (borders3.Count != 6)
                    return false;
                if ((double)(borders3.Color) != 5880731.0)
                    return false;

                if ((double)worksheet.get_Range("D18", "I18").Interior.Color != 14610923.0)
                    return false;
                if ((double)worksheet.get_Range("D21", "I21").Interior.Color != 16777215.0)
                    return false;

                if ((double)worksheet.get_Range("D10", "I11").Interior.Color != 16777215.0)
                    return false;
                if ((double)worksheet.get_Range("D16", "I17").Interior.Color != 16777215.0)
                    return false;

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool cau19(_Application a, _Workbook d)
        {
            try
            {
                Worksheet worksheet = null;
                foreach (Worksheet sheet in d.Sheets)
                {
                    if (sheet.Name == "Sales by Region")
                    {
                        worksheet = sheet;
                        break;
                    }
                }

                if (worksheet == null)
                    return false;

                bool isCHidden = (bool)((Range)worksheet.get_Range("C1", "C1")).EntireColumn.Hidden;
                if (!isCHidden)
                    return false;

                bool isLHidden = (bool)((Range)worksheet.get_Range("L1", "L1")).EntireColumn.Hidden;
                if (!isLHidden)
                    return false;

                bool isMHidden = (bool)((Range)worksheet.get_Range("M1", "M1")).EntireColumn.Hidden;
                if (!isMHidden)
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool cau20(_Application a, _Workbook d)
        {
            // ISSUE: variable of a compiler-generated type
            Worksheet worksheet = (Worksheet)null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Partner Orders")
                    worksheet = sheet;
            }
            if (worksheet == null)
                return false;
            try
            {
                // ISSUE: reference to a compiler-generated method
                int count = worksheet.get_Range((object)"F1", (object)"H1").Hyperlinks.Count;
                // ISSUE: reference to a compiler-generated method
                // ISSUE: reference to a compiler-generated method
                // ISSUE: variable of a compiler-generated type
                Hyperlink hyperlink = worksheet.get_Range((object)"F1", (object)"H1").Hyperlinks.get_Item((object)2);
                if (count != 2 || hyperlink.Name != "Order Analysis Worksheet")
                    return false;
                if (hyperlink.SubAddress != "'Order Analysis'!B2")
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau21(_Application a, _Workbook d)
        {
            // ISSUE: variable of a compiler-generated type
            Worksheet worksheet = (Worksheet)null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Order Analysis")
                    worksheet = sheet;
            }
            if (worksheet == null)
                return false;
            try
            {
                // ISSUE: reference to a compiler-generated method
                // ISSUE: variable of a compiler-generated type
                Microsoft.Office.Interop.Excel.Shape shape = worksheet.Shapes.Item((object)"Chart 1");
                if (shape.LockAspectRatio != MsoTriState.msoTrue || (int)shape.Height != 203)
                    return false;
                if ((int)shape.Width != 302)
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau22(_Application a, _Workbook d)
        {
            try
            {
                Name targetName = null;

                foreach (Name name in d.Names)
                {
                    if (name.Name == "goals_Blue_Yonder_Airlines")
                    {
                        targetName = name;
                        break;
                    }
                }

                if (targetName == null)
                    return false;

                string refersToR1C1 = targetName.RefersToR1C1Local as string;
                if (refersToR1C1 != "=Goals!R4C5:R7C5")
                    return false;

                string refersToA1 = targetName.RefersToLocal as string;
                if (refersToA1 != "=Goals!$E$4:$E$7")
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool cau23(_Application a, _Workbook d)
        {
            try
            {
                Worksheet worksheet = null;
                foreach (Worksheet sheet in d.Sheets)
                {
                    if (sheet.Name == "Sales")
                    {
                        worksheet = sheet;
                        break;
                    }
                }

                if (worksheet == null)
                    return false;

                Range cell = worksheet.get_Range("J12", "J12");

                string formulaA1 = cell.Formula as string;
                if (formulaA1 != "=MIN(C8:I8)")
                    return false;

                string formulaR1C1 = cell.FormulaR1C1 as string;
                if (formulaR1C1 != "=MIN(R[-4]C[-7]:R[-4]C[-1])")
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool cau24(_Application a, _Workbook d)
        {
            try
            {
                // ISSUE: variable of a compiler-generated type
                Worksheet worksheet = (Worksheet)null;
                foreach (Worksheet sheet in d.Sheets)
                {
                    if (sheet.Name == "Sales")
                        worksheet = sheet;
                }
                if (worksheet == null)
                    return false;
                if (worksheet.Comments.Count != 3)
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau25(_Application a, _Workbook d)
        {
            Worksheet worksheet = null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Team_ADatum")
                {
                    worksheet = sheet;
                    break;
                }
            }
            if (worksheet == null)
                return false;

            try
            {
                Range cell = worksheet.get_Range("M6", "M6");
                string formula = cell.Formula?.ToString().Replace(" ", "").ToUpper();

                // So sánh công thức sau khi chuẩn hóa bỏ khoảng trắng và viết hoa
                if (formula != "=IF(J6>=L6,\"MET\",\"MISSED\")")
                    return false;
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static bool cau26(_Application a, _Workbook d)
        {
            try
            {
                // ISSUE: variable of a compiler-generated type
                Worksheet worksheet = (Worksheet)null;
                foreach (Worksheet sheet in d.Sheets)
                {
                    if (sheet.Name == "Sales")
                        worksheet = sheet;
                }
                if (worksheet == null || (double)(int)worksheet.PageSetup.LeftMargin != 90.0 || (double)(int)worksheet.PageSetup.RightMargin != 90.0 || (double)(int)worksheet.PageSetup.TopMargin != 108.0 || (double)(int)worksheet.PageSetup.BottomMargin != 108.0 || worksheet.PageSetup.HeaderMargin != 54.0)
                    return false;
                if (worksheet.PageSetup.FooterMargin != 54.0)
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau27(_Application a, _Workbook d)
        {
            try
            {
                if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "27.csv")))
                    return false;
                if (d.FileFormat != XlFileFormat.xlCSV)
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau28(_Application a, _Workbook d)
        {
            // ISSUE: variable of a compiler-generated type
            Worksheet worksheet = (Worksheet)null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Sales")
                    worksheet = sheet;
            }
            if (worksheet == null)
                return false;
            try
            {
                if (worksheet.PageSetup.LeftHeader != "Owner" || worksheet.PageSetup.CenterHeader != "Page &P" || worksheet.PageSetup.RightHeader != "&D")
                    return false;
                if (worksheet.PageSetup.LeftFooter != "&A")
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau29(_Application a, _Workbook d)
        {
            Worksheet worksheet = null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Sales")
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet == null)
                return false;

            try
            {
                ListObject table = worksheet.ListObjects["Table2"];
                AutoFilter autoFilter = table.AutoFilter;

                // Kiểm tra cột thứ 5 và 6 đã bật filter
                if (!autoFilter.Filters[5].On || !autoFilter.Filters[6].On)
                    return false;

                // Lưu ý: Filters[] là 1-based index
                var filter5 = autoFilter.Filters[5];
                var filter6 = autoFilter.Filters[6];

                string criteria5 = filter5.Criteria1?.ToString().Replace(" ", "").ToLower();
                string criteria6 = filter6.Criteria1?.ToString().Replace(" ", "").ToLower();

                if (criteria5 != "<200" || criteria6 != ">=10000")
                    return false;
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static bool cau30(_Application a, _Workbook d)
        {
            // ISSUE: variable of a compiler-generated type
            Worksheet worksheet = (Worksheet)null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Budget")
                    worksheet = sheet;
            }
            if (worksheet == null)
                return false;
            try
            {
                // ISSUE: reference to a compiler-generated method
                // ISSUE: reference to a compiler-generated method
                // ISSUE: reference to a compiler-generated method
                if (worksheet.Shapes.Count != 5 || worksheet.Shapes.Item((object)"Picture 3").Fill.PictureEffects.Count != 3 || worksheet.Shapes.Item((object)"Picture 3").Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectBackgroundRemoval || worksheet.Shapes.Item((object)"Picture 3").Fill.PictureEffects[2].Type != MsoPictureEffectType.msoEffectGlowDiffused)
                    return false;
                // ISSUE: reference to a compiler-generated method
                if (worksheet.Shapes.Item((object)"Picture 3").Fill.PictureEffects[3].Type != MsoPictureEffectType.msoEffectSharpenSoften)
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau31(_Application a, _Workbook d)
        {
            try
            {
                Workbook resultWorkbook = a.Workbooks.get_Item("Result");

                // Kiểm tra cả 2 workbook đều có đúng 3 sheet
                if (d.Worksheets.Count != 3 || resultWorkbook.Worksheets.Count != 3)
                    return false;

                // Lấy worksheet thứ 2 trong workbook "Result"
                Worksheet ws = (Worksheet)resultWorkbook.Worksheets[2];

                // Kiểm tra tên của worksheet này
                if (ws.Name != "Partner Orders")
                    return false;
            }
            catch
            {
                return false;
            }

            return true;
        }

        //private static bool cau31_Cu(_Application a, _Workbook d)
        //{
        //  try
        //  {
        //    // ISSUE: reference to a compiler-generated method
        //    // ISSUE: variable of a compiler-generated type
        //    Workbook workbook = a.Workbooks.get_Item((object) "Result");
        //    if (d.Worksheets.Count != 3 || workbook.Worksheets.Count != 3)
        //      return false;
        //    // ISSUE: reference to a compiler-generated field
        //    if (CheckWord.\u003Ccau31\u003Eo__SiteContainer50.\u003C\u003Ep__Site51 == null)
        //    {
        //      // ISSUE: reference to a compiler-generated field
        //      CheckWord.\u003Ccau31\u003Eo__SiteContainer50.\u003C\u003Ep__Site51 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (CheckWord)));
        //    }
        //    // ISSUE: reference to a compiler-generated field
        //    // ISSUE: reference to a compiler-generated field
        //    // ISSUE: variable of a compiler-generated type
        //    Worksheet worksheet = CheckWord.\u003Ccau31\u003Eo__SiteContainer50.\u003C\u003Ep__Site51.Target((CallSite) CheckWord.\u003Ccau31\u003Eo__SiteContainer50.\u003C\u003Ep__Site51, workbook.Worksheets[(object) 2]);
        //    if (worksheet.Name != "Partner Orders")
        //      return false;
        //  }
        //  catch (Exception ex)
        //  {
        //    return false;
        //  }
        //  return true;
        //}

        private static bool cau32(_Application a, _Workbook d)
        {
            // ISSUE: variable of a compiler-generated type
            Worksheet worksheet = (Worksheet)null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Order Analysis")
                    worksheet = sheet;
            }
            if (worksheet == null)
                return false;
            try
            {
                // ISSUE: reference to a compiler-generated method
                // ISSUE: reference to a compiler-generated method
                if (worksheet.Shapes.Item((object)"Picture 4").Fill.PictureEffects.Count != 2 || worksheet.Shapes.Item((object)"Picture 4").Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectBackgroundRemoval)
                    return false;
                // ISSUE: reference to a compiler-generated method
                if (worksheet.Shapes.Item((object)"Picture 4").Fill.PictureEffects[2].Type != MsoPictureEffectType.msoEffectFilmGrain)
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private static bool cau33(_Application a, _Workbook d)
        {
            Worksheet worksheet = null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Partner Orders")
                {
                    worksheet = sheet;
                    break;
                }
            }
            if (worksheet == null)
                return false;

            try
            {
                Range cellA50 = worksheet.get_Range("A50", "A50");
                string textA50 = cellA50.Text?.ToString().Trim();
                if (textA50 != "09/06/10")
                    return false;

                Range rangeA4ToA50 = worksheet.get_Range("A4", "A50");
                var font = rangeA4ToA50.Font;

                // Kiểm tra in đậm
                if (!(font.Bold is bool isBold) || !isBold)
                    return false;

                // Kiểm tra màu chữ
                double colorValue = Convert.ToDouble(font.Color);
                if (colorValue != 8210719.0)
                    return false;
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static bool cau34(_Application a, _Workbook d)
        {
            Worksheet worksheet = null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Sales")
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet == null)
                return false;

            try
            {
                Range cellJ6 = worksheet.get_Range("J6", "J6");

                // Kiểm tra giá trị hiển thị trong ô J6
                string textValue = cellJ6.Text?.ToString().Trim();
                if (textValue != "85619626.00")
                    return false;

                // Kiểm tra công thức trong ô J6
                string formulaValue = cellJ6.Formula?.ToString().Trim();
                if (formulaValue != "=SUM(Total)")
                    return false;
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static bool cau35(_Application a, _Workbook d)
        {
            Worksheet worksheet = null;
            foreach (Worksheet sheet in d.Sheets)
            {
                if (sheet.Name == "Summary")
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet == null)
                return false;

            try
            {
                Range cellF12 = worksheet.get_Range("F12", "F12");
                string formula = cellF12.Formula?.ToString().Trim();

                if (formula != "=('Quarter 2'!F12-'Quarter 1'!F12)/'Quarter 1'!F12")
                    return false;
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static bool cau36(_Application a, _Workbook d)
        {
            try
            {
                Style titleStyle = d.Styles["Title"];
                object fontSizeObj = titleStyle.Font.Size;

                if (fontSizeObj == null)
                    return false;

                double fontSize;
                if (!double.TryParse(fontSizeObj.ToString(), out fontSize))
                    return false;

                if (fontSize != 20.0)
                    return false;
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }
    }
}
