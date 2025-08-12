// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Sec0
// Assembly: MOS_EXCEL_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace MOS_EXCEL_LEARN
{
    internal class Sec0
    {
        public static string CheckCau(int causo, Application a, Workbook d)
        {
            switch (causo)
            {
                case 1:
                    return Sec0.cau3(a, d);
                case 2:
                    return Sec0.cau8(a, d);
                case 3:
                    return Sec0.cau9(a, d);
                case 4:
                    return Sec0.cau5(a, d);
                case 5:
                    return Sec0.cau17(a, d);
                case 6:
                    return Sec0.cau18(a, d);
                case 7:
                    return Sec0.cau15(a, d);
                case 8:
                    return Sec0.cau24(a, d);
                case 9:
                    return Sec0.cau7(a, d);
                case 10:
                    return Sec0.cau6(a, d);
                case 11:
                    return Sec0.cau2(a, d);
                case 12:
                    return Sec0.cau1(a, d);
                case 13:
                    return Sec0.cau1_1(a, d);
                case 14:
                    return Sec0.cau27(a, d);
                case 15:
                    return Sec0.cau25(a, d);
                case 16:
                    return Sec0.cau26(a, d);
                case 17:
                    return Sec0.cau12(a, d);
                case 18:
                    return Sec0.dorongcot(a, d);
                case 19:
                    return Sec0.cau19(a, d);
                case 20:
                    return Sec0.cau20(a, d);
                case 21:
                    return Sec0.cau21(a, d);
                case 22:
                    return Sec0.cau22(a, d);
                case 23:
                    return Sec0.cau23(a, d);
                case 24:
                    return Sec0.cau24(a, d);
                case 25:
                    return Sec0.cau25(a, d);
                case 26:
                    return Sec0.cau26(a, d);
                case 27:
                    return Sec0.cau27(a, d);
                case 28:
                    return Sec0.cau27(a, d);
                case 29:
                    return Sec0.cau27(a, d);
                default:
                    return "default sec0";
            }
        }
        private static string cau1(Application excelApp, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["roster"] as Worksheet;
                if (worksheet == null)
                    return "False (Không tìm thấy sheet 'roster')";

                // Duyệt dòng 1 đến 6: phải ẩn
                for (int row = 1; row <= 6; row++)
                {
                    Range range = worksheet.Range[$"A{row}"];
                    if (range.EntireRow.Hidden == false)
                        return $"False(ẩn dòng {row})";
                }

                // Dòng 7: không được ẩn
                Range row7 = worksheet.Range["A7"];
                if (row7.EntireRow.Hidden == false)
                {
                    // Ý muốn kiểm tra xem dòng 7 PHẢI bị ẩn,
                    // nhưng hiện tại KHÔNG bị ẩn => sai
                    return "False(không ẩn dòng 7)";
                }

                return "True";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }
        }

        private static string cau1_1(Application excelApp, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["roster"] as Worksheet;
                if (worksheet == null)
                    return "False (Không tìm thấy sheet 'roster')";

                // Duyệt dòng 1 đến 6, kiểm tra xem có bị ẩn không
                for (int row = 1; row <= 6; row++)
                {
                    Range range = worksheet.Range[$"A{row}"];
                    if (!range.EntireRow.Hidden)
                        return $"False(mở ẩn dòng {row})";
                }

                return "True";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }
        }


        private static string cau2(Application excelApp, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["students"] as Worksheet;
                if (worksheet == null)
                    return "False (Không tìm thấy sheet 'students')";

                string d1Text = worksheet.Range["D1"].Text.ToString();
                if (string.IsNullOrWhiteSpace(d1Text))
                    return "False";

                string e1Text = worksheet.Range["E1"].Text.ToString();
                if (e1Text != "GPA")
                    return "False";

                return "True";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }
        }

        private static string cau3(Application excelApp, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Microsoft Word"] as Worksheet;
                if (worksheet == null)
                    return "False(Microsoft Word)";
            }
            catch (Exception)
            {
                return "False(Microsoft Word)";
            }
            return "True";
        }

        private static string cau4(Application app, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Q3 Sales"] as Worksheet;
                if (worksheet == null)
                    return "False (Q3 Sales)";

                if (worksheet.Index <= 2)
                    return "False(sai vị trí)";
            }
            catch (Exception)
            {
                return "False (Q3 Sales)";
            }
            return "True";
        }

        private static string cau5(Application app, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["October (2)"] as Worksheet;
                if (worksheet == null)
                    return "False (tạo bảng copy)";

                if (worksheet.Index <= 1)
                    return "False(move to end)";
            }
            catch (Exception)
            {
                return "False (tạo bảng copy)";
            }
            return "True";
        }

        private static string cau6(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Donor Contact Info"] as Worksheet;
                if (ws == null)
                    return "False (Something not finish!)";

                string cellI1 = ws.get_Range("I1", "I1").Text?.ToString();
                if (cellI1 != "Cell Phone")
                    return "False (Cell Phone)";

                string cellH1 = ws.get_Range("H1", "H1").Text?.ToString();
                if (cellH1 != "Home Phone")
                    return "False (Home Phone)";

                string cellJ1 = ws.get_Range("J1", "J1").Text?.ToString();
                if (cellJ1 != "Email")
                    return "False (Email)";
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau7(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Shareholders Info"] as Worksheet;
                if (ws == null)
                    return "False (Sheet not found)";

                double row2Height = ws.get_Range("A2", "A2").EntireRow.Height;

                if (row2Height != 30)
                    return "False (row 2)"; // Hàng 2 không có chiều cao 30
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau8(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Outdoor sports"] as Worksheet;
                if (ws == null)
                    return "False (Sheet not found)";

                if (ws.Index != 3)
                    return "False(Sửa tên Outdoor toys)";
            }
            catch
            {
                return "False (Rename)";
            }

            return "True";
        }

        private static string cau9(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Donor List"] as Worksheet;
                if (ws == null)
                    return "False (Sheet not found)";

                if (ws.Index <= 3)
                    return "False(Move to end)";
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau10(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Top Donors"] as Worksheet;
                if (ws == null)
                    return "False (Sheet not found)";

                if (ws.Index <= 3)
                    return "False(Move to end)";
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau11(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Top Donors"] as Worksheet;
                if (ws == null)
                    return "False (Sheet not found)";

                if (ws.Index != 1)
                    return "False(Move to first)";
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau12(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Products"] as Worksheet;
                if (ws == null)
                    return "False (Sheet not found)";

                Range rng = ws.Range["A1"];
                int align = Convert.ToInt32(rng.HorizontalAlignment);

                if (align != -4131) // -4131 là xlLeft
                    return "False(Left)";
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau13(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Demographics"] as Worksheet;
                if (ws == null)
                    return "False (Sheet not found)";

                Range g9 = ws.Range["G9"];
                Range h9 = ws.Range["H9"];
                Range h8 = ws.Range["H8"];

                string g9Text = g9.Text?.ToString() ?? "";
                string h9Text = h9.Text?.ToString() ?? "";
                string h8Text = h8.Text?.ToString() ?? "";

                if (string.IsNullOrWhiteSpace(g9Text) || string.IsNullOrWhiteSpace(h9Text))
                    return "False(xoa dong cuoi)";

                if (string.IsNullOrWhiteSpace(h8Text))
                    return "False (chỉ xoa dong cuoi)";
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau14(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["indoor toys"] as Worksheet;
                if (ws == null)
                    return "False (đổi tên indoor toys)";

                if (ws.Index != 4)
                    return "False (sai trang tinh)";
            }
            catch
            {
                return "False (đổi tên indoor toys)";
            }

            return "True";
        }

        private static string cau15(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Outdoor toys"] as Worksheet;
                if (ws == null)
                    return "False (Tên sheet sai)";

                // Lấy mã màu của tab (Tab.Color có kiểu int)
                int tabColor = (int)(ws.Tab.Color);

                // So sánh với mã màu mong muốn (12419407)
                if (tabColor != 12419407)
                    return "False";
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau16(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Outdoor toys"] as Worksheet;
                if (ws == null)
                    return "False (Không tìm thấy sheet)";

                if (ws.Visible != XlSheetVisibility.xlSheetHidden)
                    return "False";
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau17(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Outdoor toys"] as Worksheet;
                if (ws == null)
                    return "False (Không tìm thấy sheet)";

                if (ws.Visible != XlSheetVisibility.xlSheetHidden)
                    return "False";
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau18(Application app, Workbook wb)
        {
            try
            {
                foreach (Worksheet sheet in wb.Worksheets)
                {
                    if (sheet.Visible == XlSheetVisibility.xlSheetHidden)
                        return $"False({sheet.Name})";
                }
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau19(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["Shareholders Info"];
                if (sheet.Visible != XlSheetVisibility.xlSheetVisible)
                    return "False";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau20(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet1 = wb.Worksheets["Sheet1"];
                if (sheet1.Visible != XlSheetVisibility.xlSheetVisible)
                    return "False(bỏ ẩn Sheet1)";

                Worksheet shareholders = wb.Worksheets["Shareholders Info"];
                if (shareholders.Visible == XlSheetVisibility.xlSheetVisible)
                    return "False (vẫn ẩn Shareholders Info)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau21(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["Materials"];
                Range cell = worksheet.get_Range("A1", "A1");
                if (!cell.WrapText)
                    return "False (A1)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau22(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["Materials"];

                // Kiểm tra Freeze Panes đã bật chưa
                if (!a.ActiveWindow.FreezePanes)
                    return "False(FreezePanes)";

                // Kiểm tra địa chỉ của ô hiện tại (active cell) có phải là A6 không
                string address = a.ActiveCell.get_Address(
                    Type.Missing, Type.Missing,
                    XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                if (address != "$A$6")
                    return "False(để con trỏ ô A6 trước khi FreezePanes)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau23(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["Materials"];

                if (worksheet.PageSetup.CenterFooter != "Page &P of &N")
                    return "False";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau24(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["Instructional Hours"];
                double widthB5 = ((Range)worksheet.Range["B5", "B5"]).Width;

                string[] columns = { "C5", "D5", "E5", "F5", "G5" };

                foreach (var cell in columns)
                {
                    double otherWidth = ((Range)worksheet.Range[cell, cell]).Width;
                    if (otherWidth != widthB5)
                        return "False";
                }
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau25(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["Last Semester"];
                Range range = worksheet.Range["A3", "F3"];

                // Nếu WrapText == true => không đạt yêu cầu
                if ((bool)range.WrapText == true)
                    return "False (A3:F3)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau26(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["New Policies"];
                Range range = worksheet.Range["A5", "A13"];

                // Kiểm tra canh trái (Left = -4131)
                if ((int)range.HorizontalAlignment != -4131)
                    return "False(Left(Indent))";

                // Kiểm tra thụt đầu dòng là 1
                if (range.IndentLevel != 1)
                    return "False(Indent 1)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau27(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["Products"];

                // Kiểm tra xem đã FreezePanes chưa
                if (!a.ActiveWindow.FreezePanes)
                    return "False(FreezePanes)";

                // Kiểm tra con trỏ hiện tại có ở đúng ô A3 không
                string activeAddress = a.ActiveCell.get_Address(
                                        Missing.Value,
                                        Missing.Value,
                                        XlReferenceStyle.xlA1,
                                        Missing.Value,
                                        Missing.Value);

                if (activeAddress != "$A$3")
                    return "False(để con trỏ ô A3 trước khi FreezePanes)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string dorongcot(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["Materials"];

                double widthA4 = ((Range)worksheet.get_Range("A4", "A4")).Width;
                if (widthA4 < 159.0 || widthA4 > 161.0)
                    return "False";

                double widthB4 = ((Range)worksheet.get_Range("B4", "B4")).Width;
                if (widthB4 < 68.0 || widthB4 > 70.0)
                    return "False";

                double widthN4 = ((Range)worksheet.get_Range("N4", "N4")).Width;
                if (widthN4 < 35.0 || widthN4 > 37.0)
                    return "False";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }
    }
}
