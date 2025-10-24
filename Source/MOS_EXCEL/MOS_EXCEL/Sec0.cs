// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Sec0
// Assembly: MOS_EXCEL_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
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
                    //return Sec0.cau25(a, d);
                    return Sec0.cau15New(a, d);
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
                    return "False";

                // Duyệt dòng 1 đến 6: phải ẩn
                for (int row = 1; row <= 6; row++)
                {
                    Range range = worksheet.Range[$"A{row}"];
                    if (range.EntireRow.Hidden == false)
                        return "False";
                }

                // Dòng 7: không được ẩn
                Range row7 = worksheet.Range["A7"];
                if (row7.EntireRow.Hidden == true)
                {
                    // Ý muốn kiểm tra xem dòng 7 PHẢI bị ẩn,
                    // nhưng hiện tại KHÔNG bị ẩn => sai
                    return "False";
                }

                return "True";
            }
            catch (Exception)
            {
                return "False";
            }
        }

        private static string cau1_1(Application excelApp, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["roster"] as Worksheet;
                if (worksheet == null)
                    return "False";

                // Duyệt dòng 1 đến 6, kiểm tra xem có bị ẩn không
                for (int row = 1; row <= 6; row++)
                {
                    Range range = worksheet.Range[$"A{row}"];
                    if (range.EntireRow.Hidden == true)
                        return "False";
                }

                return "True";
            }
            catch (Exception)
            {
                return "False";
            }
        }


        private static string cau2(Application excelApp, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["students"] as Worksheet;
                if (worksheet == null)
                    return "False";

                string d1Text = worksheet.Range["D1"].Text.ToString();
                if (!string.IsNullOrEmpty(d1Text))
                    return "False";

                //string e1Text = worksheet.Range["E1"].Text.ToString();
                if (d1Text.Trim() == "GPA")
                    return "False";

                string d30Text = worksheet.Range["D30"].Text.ToString();
                if (!string.IsNullOrEmpty(d30Text))
                    return "False";
                if (d30Text.Trim() == "2")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau3(Application excelApp, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Microsoft Word"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }
            return "True";
        }

        private static string cau4(Application app, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Q3 Sales"] as Worksheet;
                if (worksheet == null)
                    return "False";

                if (worksheet.Index <= 2)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }
            return "True";
        }

        private static string cau5(Application app, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["October (2)"] as Worksheet;
                if (worksheet == null)
                    return "False";

                if (worksheet.Index <= 1)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }
            return "True";
        }

        private static string cau6(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Donor Contact Info"] as Worksheet;
                if (ws == null)
                    return "False";

                string cellI1 = ws.get_Range("I1", "I1").Text?.ToString();
                if (cellI1 != "Cell Phone")
                    return "False";

                string cellH1 = ws.get_Range("H1", "H1").Text?.ToString();
                if (cellH1 != "Home Phone")
                    return "False";

                string cellJ1 = ws.get_Range("J1", "J1").Text?.ToString();
                if (cellJ1 != "Email")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau7(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Shareholders Info"] as Worksheet;
                if (ws == null)
                    return "False";

                double row2Height = ws.get_Range("A2", "A2").EntireRow.Height;

                if (row2Height != 30)
                    return "False"; // Hàng 2 không có chiều cao 30
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau8(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Outdoor sports"] as Worksheet;
                if (ws == null)
                    return "False";

                if (ws.Index != 3)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau9(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Donor List"] as Worksheet;
                if (ws == null)
                    return "False";

                if (ws.Index <= 3)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau10(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Top Donors"] as Worksheet;
                if (ws == null)
                    return "False";

                if (ws.Index <= 3)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau11(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Top Donors"] as Worksheet;
                if (ws == null)
                    return "False";

                if (ws.Index != 1)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau12(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Products"] as Worksheet;
                if (ws == null)
                    return "False";

                Range rng = ws.Range["A1"];
                int align = Convert.ToInt32(rng.HorizontalAlignment);

                if (align != -4131) // -4131 là xlLeft
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau13(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Demographics"] as Worksheet;
                if (ws == null)
                    return "False";

                Range g9 = ws.Range["G9"];
                Range h9 = ws.Range["H9"];
                Range h8 = ws.Range["H8"];

                string g9Text = g9.Text?.ToString() ?? "";
                string h9Text = h9.Text?.ToString() ?? "";
                string h8Text = h8.Text?.ToString() ?? "";

                if (string.IsNullOrWhiteSpace(g9Text) || string.IsNullOrWhiteSpace(h9Text))
                    return "False";

                if (string.IsNullOrWhiteSpace(h8Text))
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau14(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["indoor toys"] as Worksheet;
                if (ws == null)
                    return "False";

                if (ws.Index != 4)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau15(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Outdoor toys"] as Worksheet;
                if (ws == null)
                    return "False";

                // Lấy mã màu của tab (Tab.Color có kiểu int)
                int tabColor = (int)(ws.Tab.Color);

                // So sánh với mã màu mong muốn (12419407)
                if (tabColor != 12419407)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau16(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Outdoor toys"] as Worksheet;
                if (ws == null)
                    return "False";

                if (ws.Visible != XlSheetVisibility.xlSheetHidden)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau17(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Outdoor toys"] as Worksheet;
                if (ws == null)
                    return "False";

                if (ws.Visible != XlSheetVisibility.xlSheetHidden)
                    return "False";
            }
            catch
            {
                return "False";
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
                        return "False";
                }
            }
            catch (Exception)
            {
                return "False";
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
                return "False";
            }

            return "True";
        }


        private static string cau20(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet1 = wb.Worksheets["Sheet1"];
                if (sheet1.Visible != XlSheetVisibility.xlSheetVisible)
                    return "False";

                Worksheet shareholders = wb.Worksheets["Shareholders Info"];
                if (shareholders.Visible == XlSheetVisibility.xlSheetVisible)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
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
                    return "False";
            }
            catch (Exception)
            {
                return "False";
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
                    return "False";

                // Kiểm tra địa chỉ của ô hiện tại (active cell) có phải là A6 không
                string address = a.ActiveCell.get_Address(
                    Type.Missing, Type.Missing,
                    XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                if (address != "$A$6")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
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
                return "False";
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
                return "False";
            }

            return "True";
        }

        private static string cau15New(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["Last Semester"];
                Range range = worksheet.Range["A3", "F3"];

                // Nếu WrapText == true => đạt yêu cầu
                if ((bool)range.WrapText == true)
                    return "True";
            }
            catch (Exception)
            {
                return "False";
            }

            return "False";
        }

        private static string cau25(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["Last Semester"];
                Range range = worksheet.Range["A3", "F3"];

                // Nếu WrapText == true => không đạt yêu cầu
                if ((bool)range.WrapText == true)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
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
                    return "False";

                // Kiểm tra thụt đầu dòng là 1
                if (range.IndentLevel != 1)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
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
                    return "False";

                //// Kiểm tra con trỏ hiện tại có ở đúng ô A3 không
                //string activeAddress = a.ActiveCell.get_Address(
                //                        Missing.Value,
                //                        Missing.Value,
                //                        XlReferenceStyle.xlA1,
                //                        Missing.Value,
                //                        Missing.Value);

                //if (activeAddress != "$A$3")
                //    return "False(để con trỏ ô A3 trước khi FreezePanes)";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string dorongcot(Application a, Workbook d)
        {
            try
            {
                //Worksheet ws = (Worksheet)d.Worksheets["Materials"];
                //if (ws == null)
                //    return "False";

                //double wA = ((Range)ws.get_Range("A4")).ColumnWidth;
                //double wB = ((Range)ws.get_Range("B4")).ColumnWidth;
                //double wN = ((Range)ws.get_Range("N4")).ColumnWidth;

                //if (wA < 34.0 || wA > 35.0) return "False";
                //if (wB < 14.0 || wB > 15.0) return "False";
                //if (wN < 6.0 || wN > 8.0) return "False";

                Worksheet ws = (Worksheet)d.Worksheets["Materials"];
                if (ws == null)
                    return "False (Không tìm thấy sheet 'Materials')";

                Range rng = ws.Range["A:E"];
                int n = rng.Columns.Count;

                // lấy độ rộng hiện tại
                double[] current = new double[n];
                for (int i = 1; i <= n; i++)
                    current[i - 1] = ((Range)rng.Columns[i]).ColumnWidth;

                // tạo bản sao tạm
                Worksheet temp = d.Worksheets.Add();
                Range tempRange = temp.Range["A1:E100"];

                // copy nội dung và định dạng để tính chính xác độ rộng
                ws.Range["A1:E100"].Copy();
                tempRange.PasteSpecial(XlPasteType.xlPasteAll);

                // gọi AutoFit trên sheet tạm
                tempRange.Columns.AutoFit();

                // lấy độ rộng sau AutoFit
                double[] autofit = new double[n];
                for (int i = 1; i <= n; i++)
                    autofit[i - 1] = ((Range)tempRange.Columns[i]).ColumnWidth;

                // tắt cảnh báo trước khi xóa
                Application app = d.Application;
                bool oldAlerts = app.DisplayAlerts;
                app.DisplayAlerts = false;

                // xóa sheet tạm mà không hiển thị thông báo
                temp.Delete();

                // bật lại cảnh báo
                app.DisplayAlerts = oldAlerts;

                // so sánh độ rộng hiện tại với độ rộng AutoFit
                bool ok = true;
                for (int i = 0; i < n; i++)
                {
                    if (Math.Abs(current[i] - autofit[i]) > 0.2)
                    {
                        ok = false;
                        //return ($"Cột {i + 1}: hiện tại = {current[i]}, AutoFit = {autofit[i]}");
                        break;
                    }
                }

                if (ok)
                    return "True";
                else
                    return "False";

            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
    }
}
