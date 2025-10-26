using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;
using mscore = Microsoft.Office.Core;

namespace MOS_WORD_TEST
{
    public interface IQuestionCheck
    {
        string CheckAnswer(int questionNumber, Application a, Workbook d);
    }
    internal class QuestionCheck : IQuestionCheck
    {
        public string CheckAnswer(int questionNumber, Application a, Workbook d)
        {
            switch (questionNumber)
            {
                case 1: return Cau1(a, d);
                case 2: return Cau2(a, d);
                case 3: return Cau3(a, d);
                case 4: return Cau4(a, d);
                case 5: return Cau5(a, d);
                case 6: return Cau6(a, d);
                case 7: return Cau7(a, d);
                case 8: return Cau8(a, d);
                case 9: return Cau9(a, d);
                case 10: return Cau10(a, d);
                case 11: return Cau11(a, d);
                case 12: return Cau12(a, d);
                case 13: return Cau13(a, d);
                case 14: return Cau14(a, d);
                case 15: return Cau15(a, d);
                case 16: return Cau16(a, d);
                case 17: return Cau17(a, d);
                case 18: return Cau18(a, d);
                case 19: return Cau19(a, d);
                case 20: return Cau20(a, d);
                case 21: return Cau21(a, d);
                case 22: return Cau22(a, d);
                case 23: return Cau23(a, d);
                case 24: return Cau24(a, d);
                case 25: return Cau25(a, d);
                case 26: return Cau26(a, d);
                case 27: return Cau27(a, d);
                case 28: return Cau28(a, d);
                case 29: return Cau29(a, d);
                case 30: return Cau30(a, d);
                case 31: return Cau31(a, d);
                case 32: return Cau32(a, d);
                case 33: return Cau33(a, d);
                case 34: return Cau34(a, d);
                case 35: return Cau35(a, d);
                case 36: return Cau36(a, d);
                case 37: return Cau37(a, d);
                case 38: return Cau38(a, d);
                case 39: return Cau39(a, d);
                case 40: return Cau40(a, d);
                case 41: return Cau41(a, d);
                case 42: return Cau42(a, d);
                case 43: return Cau43(a, d);
                case 44: return Cau44(a, d);
                case 45: return Cau45(a, d);
                case 46: return Cau46(a, d);
                case 47: return Cau47(a, d);
                case 48: return Cau48(a, d);
                case 49: return Cau49(a, d);
                case 50: return Cau50(a, d);
                case 51: return Cau51(a, d);
                case 52: return Cau52(a, d);
                case 53: return Cau53(a, d);
                case 54: return Cau54(a, d);
                case 55: return Cau55(a, d);
                case 56: return Cau56(a, d);
                case 57: return Cau57(a, d);
                case 58: return Cau58(a, d);
                case 59: return Cau59(a, d);
                case 60: return Cau60(a, d);
                case 61: return Cau61(a, d);
                case 62: return Cau62(a, d);
                case 63: return Cau63(a, d);
                case 64: return Cau64(a, d);
                case 65: return Cau65(a, d);
                case 66: return Cau66(a, d);
                case 67: return Cau67(a, d);
                case 68: return Cau68(a, d);
                case 69: return Cau69(a, d);
                case 70: return Cau70(a, d);
                case 71: return Cau71(a, d);
                case 72: return Cau72(a, d);
                case 73: return Cau73(a, d);
                case 74: return Cau74(a, d);
                case 75: return Cau75(a, d);
                case 76: return Cau76(a, d);
                case 77: return Cau77(a, d);
                case 78: return Cau78(a, d);
                case 79: return Cau79(a, d);
                case 80: return Cau80(a, d);
                case 81: return Cau81(a, d);
                case 82: return Cau82(a, d);
                case 83: return Cau83(a, d);
                case 84: return Cau84(a, d);
                case 85: return Cau85(a, d);
                case 86: return Cau86(a, d);
                case 87: return Cau87(a, d);
                case 88: return Cau88(a, d);
                case 89: return Cau89(a, d);
                case 90: return Cau90(a, d);
                case 91: return Cau91(a, d);
                case 92: return Cau92(a, d);
                case 93: return Cau93(a, d);
                case 94: return Cau94(a, d);
                case 95: return Cau95(a, d);
                case 96: return Cau96(a, d);
                case 97: return Cau97(a, d);
                case 98: return Cau98(a, d);
                case 99: return Cau99(a, d);
                case 100: return Cau100(a, d);
                case 101: return Cau101(a, d);
                case 102: return Cau102(a, d);
                case 103: return Cau103(a, d);
                case 104: return Cau104(a, d);
                case 105: return Cau105(a, d);
                case 106: return Cau106(a, d);
                case 107: return Cau107(a, d);
                case 108: return Cau108(a, d);
                case 109: return Cau109(a, d);
                case 110: return Cau110(a, d);
                case 111: return Cau111(a, d);
                default:
                    return "Invalid question number";
            }
        }

        private string Cau1(Application a, Workbook d)
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
        private string Cau9(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Microsoft Word"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }
            return "True";
        }
        private string Cau10(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Donor Contact Info"] as Worksheet;
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
        private string Cau17(Application a, Workbook d)
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
        private string Cau28(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Outdoor toys"] as Worksheet;
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
        private string Cau29(Application a, Workbook d)
        {
            try
            {
                foreach (Worksheet sheet in d.Worksheets)
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
        private string Cau30(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Outdoor toys"] as Worksheet;
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
        private string Cau31(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Outdoor sports"] as Worksheet;
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
        private string Cau32(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Shareholders Info"] as Worksheet;
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
        private string Cau48(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Microsoft Word"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }
            return "True";
        }
        private string Cau55(Application a, Workbook d)
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
        private string Cau61(Application a, Workbook d)
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
        private string Cau67(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["October (2)"] as Worksheet;
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
        private string Cau71(Application a, Workbook d)
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
        private string Cau72(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Products"] as Worksheet;
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
        private string Cau85(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["students"] as Worksheet;
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
        private string Cau86(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["roster"] as Worksheet;
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
        private string Cau87(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["roster"] as Worksheet;
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
        private string Cau2(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Substitutes"];
                object cellValue = worksheet.Range["A1", "A1"].Text;

                string value = cellValue != null ? cellValue.ToString() : "";

                if (value != "Rank")
                {
                    return "False";
                }
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau7(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Customers by Order"] as Worksheet;
                if (worksheet == null)
                    return "False";

                // Lấy giá trị ô C502
                Range cellC502 = worksheet.Range["C502"];
                string valueC502 = cellC502?.Text?.ToString() ?? "";

                if (!string.IsNullOrEmpty(valueC502))
                    return "False";

                // Lấy giá trị ô B502
                Range cellB502 = worksheet.Range["B502"];
                string valueB502 = cellB502?.Text?.ToString() ?? "";

                if (!string.IsNullOrEmpty(valueB502))
                    return "False";

                return "True";
            }
            catch (Exception)
            {
                return "False";
            }
        }
        private string Cau11(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Top Donors"] as Worksheet;
                if (worksheet == null)
                    return "False";

                // Kiểm tra ô A2 có phải là "Mark Bebbington" hay không
                string cellA2 = worksheet.Range["A2"].Text?.ToString()?.Trim();
                if (!string.Equals(cellA2, "Mark Bebbington", StringComparison.OrdinalIgnoreCase))
                    return "False";

                // Kiểm tra ô C6 có phải là "Platinum" hay không
                string cellC6 = worksheet.Range["C6"].Text?.ToString()?.Trim();
                if (!string.Equals(cellC6, "Platinum", StringComparison.OrdinalIgnoreCase))
                    return "False";

                return "True";
            }
            catch
            {
                return "False";
            }
        }
        private string Cau12(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Top Donors"] as Worksheet;

                string a2 = ws.Range["A2"].Text.ToString();
                if (a2 != "Daniel P. Taylor")
                    return "False";

                string c2 = ws.Range["C2"].Text.ToString();
                if (c2 != "Charles Fitzgerald")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau13(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Donor Contact Info"] as Worksheet;

                string f2 = ws.Range["F2"].Text.ToString();
                string f3 = ws.Range["F3"].Text.ToString();
                string h46 = ws.Range["H46"].Text.ToString();

                if (f2 != "64043" || f3 != "64043")
                    return "False";

                if (h46 != "314-555-64044")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau14(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Donor List"];

                string c16 = worksheet.Range["C16"].Text?.ToString();
                if (c16 != "Bronze Count")
                    return "False";

                string c27 = worksheet.Range["C27"].Text?.ToString();
                if (c27 != "Silver Count")
                    return "False";

                string c67 = worksheet.Range["C67"].Text?.ToString();
                if (c67 != "Platinum Count")
                    return "False";

                string a67 = worksheet.Range["A67"].Text?.ToString();
                if (a67 != "28")
                    return "False";

                if (worksheet.HPageBreaks.Count != 4)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau15(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Donor List"];

                string cellValue = worksheet.Range["G8", "G8"].Text.ToString();

                // Nếu ô G8 khác rỗng thì trả về "False"
                if (!string.IsNullOrEmpty(cellValue))
                {
                    return "False";
                }
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau16(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Demographics"];
            }
            catch (Exception)
            {
                return "False";
            }

            try
            {
                var value = worksheet.Range["C5", "C5"].Value;

                string cellValue = value != null ? value.ToString() : "";

                if (cellValue != "300")
                {
                    return "False";
                }
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau42(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Exchange Rates"];
                if (worksheet == null)
                    return "False";

                Range cellRange = worksheet.Range["A7", "D8"];
                if (cellRange == null)
                    return "False";

                // Duyệt qua từng ô trong vùng
                foreach (Range cell in cellRange.Cells)
                {
                    object val = cell.Value2;

                    if (val != null && val != DBNull.Value && val.ToString().Trim() != "")
                    {
                        // Nếu có ít nhất 1 ô có dữ liệu → sai
                        return "False";
                    }
                }
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau70(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["Top 25"];

                string cellA2 = worksheet.get_Range("A2", "A2").Text.ToString();
                if (cellA2 != "Rank")
                    return "False";

                string cellG27 = worksheet.get_Range("G27", "G27").Text.ToString();
                if (cellG27 != "London")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau78(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Quote"];

                string e12 = ws.Range["E12", "E12"].Text.ToString();
                if (e12.Trim() != "Quantity")
                    return "False";

                string f12 = ws.Range["F12", "F12"].Text.ToString();
                if (f12.Trim() != "Unit Price")
                    return "False";

                string f30 = ws.Range["F30", "F30"].Text.ToString();
                if (f30.Trim() != "Total")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau79(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet = d.Worksheets["Quote"];

                string e12 = sheet.Range["E12"].Text?.ToString().Trim();
                if (e12 != "Quantity")
                    return "False";

                string f12 = sheet.Range["F12"].Text?.ToString();
                if (f12 != "Unit Price ") // có dấu cách ở cuối
                    return "False";

                string f30 = sheet.Range["F30"].Text?.ToString().Trim();
                if (f30 != "Total")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau92(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Social Media Ads"] as Worksheet;
                if (worksheet == null)
                    return "False";

                // Lấy giá trị ô A17
                Range cellA17 = worksheet.Range["A17"];
                string valueA17 = cellA17?.Text?.ToString()?.Trim() ?? "";

                if (!string.Equals(valueA17, "RANCH", StringComparison.OrdinalIgnoreCase))
                    return "False";

                return "True";
            }
            catch (Exception)
            {
                return "False";
            }
        }
        private string Cau33(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Games"];
                if (worksheet == null)
                    return "False";

                //// Nếu cả vùng A12:B18 bị merge lại thì sai
                //Range mergedRange = worksheet.Range["A12", "B18"];
                //if ((bool)mergedRange.MergeCells == true)
                //    return "False(chọn merge across) 1";

                //// Nếu A12:B12 không được merge thì sai
                //Range topRow = worksheet.Range["A12", "B12"];
                //if ((bool)topRow.MergeCells != true)
                //    return "False(chọn merge across)";

                //// Nếu A18:B18 không được merge thì sai
                //Range bottomRow = worksheet.Range["A18", "B18"];
                //if ((bool)bottomRow.MergeCells != true)
                //    return "False(chọn merge across)";

                // Vùng kiểm tra tổng
                Range fullRange = worksheet.Range["A12", "B18"];

                // Nếu toàn vùng bị merge 1 khối → sai
                object fullRangeVal = fullRange.MergeCells;
                if (fullRange.MergeCells != null && fullRangeVal != DBNull.Value && (bool)fullRange.MergeCells)
                    return "False";

                // Duyệt từng dòng để đảm bảo mỗi hàng được merge across
                for (int row = 12; row <= 18; row++)
                {
                    Range rowRange = worksheet.Range[$"A{row}", $"B{row}"];
                    object mergeValue = rowRange.MergeCells;

                    if (mergeValue == null || mergeValue == DBNull.Value || (bool)mergeValue == false)
                        return "False";
                }
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau43(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Exchange Rates"];
                if (worksheet == null)
                    return "False";

                Range range = worksheet.Range["B4", "D8"];
                string numberFormat = range.NumberFormat as string;

                if (numberFormat != "0.00")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau56(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Materials"] as Worksheet;
                if (worksheet == null)
                    return "False";

                Range range = worksheet.get_Range("A1", "N1");
                if (range == null)
                    return "False";

                // Kiểm tra MergeCells
                if (!(range.MergeCells is bool merged) || !merged)
                    return "False";

                // Kiểm tra canh lề ngang (1 = xlLeft)
                if (!(range.HorizontalAlignment is int alignment) || alignment != 1)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau80(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Prices"];
                Range range = worksheet.Range["A1", "A1"];
                string styleName = range.Style.Name;

                if (styleName != "Title")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau105(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Projects"];

                Range a1 = worksheet.Range["A1"];
                Range a2 = worksheet.Range["A2"];

                // Kiểm tra A1 đã merge chưa
                if (!(bool)a1.MergeCells)
                    return "False";

                // Kiểm tra A1 căn giữa (1 == xlCenter)
                if ((int)a1.HorizontalAlignment != 1)
                    return "False";

                // Kiểm tra A1 có màu nền là 14408667 không
                if ((int)(a1.Interior.Color) != 14408667)
                    return "False";

                // Kiểm tra A2 đã merge chưa
                if (!(bool)a2.MergeCells)
                    return "False";

                // Kiểm tra A2 căn giữa
                if ((int)a2.HorizontalAlignment != 1)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau18(Application a, Workbook d)
        {
            try
            {
                if (d.Names.Count != 1)
                    return "False";

                Name rangeName = d.Names.Item(1);
                if (rangeName.Name != "Enrollment")
                    return "False";

                string refersTo = rangeName.RefersToLocal?.ToString() ?? "";
                if (refersTo != "='Enrollment Summary'!$A$3:$B$7")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau23(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Non_Fiction"];

                // Kiểm tra công thức ở ô H5
                string formulaH5 = worksheet.Range["H5"].Formula.ToString();
                if (formulaH5 != "=F5-G5")
                    return "False";

                // Kiểm tra công thức ở ô H35
                string formulaH35 = worksheet.Range["H35"].Formula.ToString();
                if (formulaH35 != "=F35-G35")
                    return "False";

                // Kiểm tra định dạng số ở ô H5
                string numberFormatH5 = worksheet.Range["H5"].NumberFormat.ToString();
                if (numberFormatH5 != "General")
                    return "False";

                // Kiểm tra màu nền ô H6
                string colorH6 = worksheet.Range["H6"].Interior.Color.ToString();
                if (colorH6 != "16777215") // 16777215 là màu trắng (trong Excel)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau24(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Non_Fiction"];
                string formula = worksheet.Range["F37"].Formula.ToString();

                if (!formula.Contains("=AVERAGEIF(D5:D35,\"Lucerne Publishing\",F5:F35)") &&
                    !formula.Contains("=AVERAGEIF($D$5:$D$35,\"Lucerne Publishing\",$F$5:$F$35)"))
                {
                    return "False";
                }
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau37(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["February"];

                string formulaF5 = worksheet.Range["F5"].Formula?.ToString() ?? "";
                if (!formulaF5.Contains("=LEFT([@[Policy Number ]],2)"))
                    return "False";

                string formulaF18 = worksheet.Range["F18"].Formula?.ToString() ?? "";
                if (!formulaF18.Contains("=LEFT([@[Policy Number ]],2)"))
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau38(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["February"];

                string formulaG5 = worksheet.Range["G5"].Formula?.ToString() ?? "";
                if (!formulaG5.Contains("=IF([@[Years as Member]]>3,\"Yes\",\"No\")"))
                    return "False";

                string formulaG18 = worksheet.Range["G18"].Formula?.ToString() ?? "";
                if (!formulaG18.Contains("=IF([@[Years as Member]]>3,\"Yes\",\"No\")"))
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau44(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet = d.Worksheets["Key Accounts"];

                string formulaC4 = sheet.Range["C4"].Formula.ToString();
                if (!formulaC4.Contains("=AVERAGE(Table1[@[January]:[April]])"))
                    return "False";

                string formulaC12 = sheet.Range["C12"].Formula.ToString();
                if (!formulaC12.Contains("=AVERAGE(Table1[@[January]:[April]])"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau45(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet = d.Worksheets["Contact"];

                string formulaC5 = sheet.Range["C5"].Formula.ToString();
                if (!formulaC5.Contains("=CONCAT([@[First Name]],\"@woodgrovebank.com\")"))
                    return "False";

                string formulaC19 = sheet.Range["C19"].Formula.ToString();
                if (!formulaC19.Contains("=CONCAT([@[First Name]],\"@woodgrovebank.com\")"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau49(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet = d.Worksheets["Prices"];

                string formulaJ5 = sheet.Range["J5"].Formula.ToString();
                if (!formulaJ5.Contains("=[@[Unit Price]]*$L$2"))
                    return "False";

                string formulaJ25 = sheet.Range["J25"].Formula.ToString();
                if (!formulaJ25.Contains("=[@[Unit Price]]*$L$2"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau50(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet = d.Worksheets["London"];

                string formulaE21 = sheet.Range["E21"].Formula.ToString();
                if (formulaE21 != "=[@[Air Miles]]*0.08")
                    return "False";

                string formatE21 = sheet.Range["E21"].NumberFormat.ToString();
                if (formatE21 != "General")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau51(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet = d.Worksheets["New York City"];
                string formula = sheet.Range["D23"].Formula.ToString();

                if (!formula.Contains("=MAX(Table1[Air Miles]"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau62(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["New Policies"];

                string formulaI5 = worksheet.Range["I5"].Formula?.ToString() ?? "";
                if (!formulaI5.Contains("=COUNTBLANK(Table1[@[January]:[June]])"))
                    return "False";

                string formulaI13 = worksheet.Range["I13"].Formula?.ToString() ?? "";
                if (!formulaI13.Contains("=COUNTBLANK(Table1[@[January]:[June]])"))
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau63(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Contact"];

                string formulaC5 = worksheet.Range["C5"].Formula?.ToString() ?? "";
                if (!formulaC5.Contains("=CONCAT([@[First Name]],\"@humongousinsurance.com\")"))
                    return "False";

                string formulaC13 = worksheet.Range["C13"].Formula?.ToString() ?? "";
                if (!formulaC13.Contains("=CONCAT([@[First Name]],\"@humongousinsurance.com\")"))
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau68(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["October"];
                string formula = worksheet.Range["E37"].Formula.ToString();

                if (!formula.Contains("=AVERAGEIF(E11:E35,\">300\",E11:E35)") &&
                    !formula.Contains("=AVERAGEIF($E$11:$E$35,\">300\",$E$11:$E$35)"))
                {
                    return "False";
                }
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau73(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Products"];

                string formulaG3 = worksheet.Range["G3"].Formula?.ToString() ?? "";
                if (!formulaG3.Contains("=[@[Current Value]]*Increase"))
                    return "False";

                string formulaG54 = worksheet.Range["G54"].Formula?.ToString() ?? "";
                if (!formulaG54.Contains("=[@[Current Value]]*Increase"))
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau81(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet = d.Worksheets["Prices"];

                string formulaJ5 = sheet.Range["J5"].Formula.ToString();
                if (!formulaJ5.Contains("=[@[Unit Price]]*$L$2"))
                    return "False";

                string formulaJ25 = sheet.Range["J25"].Formula.ToString();
                if (!formulaJ25.Contains("=[@[Unit Price]]*$L$2"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau82(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet = d.Worksheets["Prices"];

                string formulaG5 = sheet.Range["G5"].Formula.ToString();
                if (!formulaG5.Contains("=IF([@[Inventory Level]]<15%,\"Low\",\"\")"))
                    return "False";

                string formulaG25 = sheet.Range["G25"].Formula.ToString();
                if (!formulaG25.Contains("=IF([@[Inventory Level]]<15%,\"Low\",\"\")"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau88(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["roster"];

                string formulaC9 = worksheet.Range["C9"].Formula.ToString();
                if (!formulaC9.Contains("=UPPER(A9)"))
                    return "False";

                string formulaC66 = worksheet.Range["C66"].Formula.ToString();
                if (!formulaC66.Contains("=UPPER(A66)"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau89(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["roster"];

                string formulaB9 = worksheet.Range["B9"].Formula.ToString();
                if (!formulaB9.Contains("=LOWER(D9)"))
                    return "False";

                string formulaB66 = worksheet.Range["B66"].Formula.ToString();
                if (!formulaB66.Contains("=LOWER(D66)"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau90(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["roster"];
                Range c8 = worksheet.Range["C8"];
                string formula = c8.Formula.ToString();

                if (formula != "=PROPER(A8)")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau94(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Projections"];

                string formulaC4 = worksheet.Range["C4"].Formula?.ToString() ?? "";
                if (!formulaC4.Contains("=[@[Quarter 1]]*Q2_Increase"))
                    return "False";

                string formulaC11 = worksheet.Range["C11"].Formula?.ToString() ?? "";
                if (!formulaC11.Contains("=[@[Quarter 1]]*Q2_Increase"))
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau95(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Summary"];
                string formula = worksheet.Range["B15"].Formula.ToString();

                if (formula != "=MAX(F4:F11)")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau99(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet = d.Worksheets["Sales"];

                string formulaE2 = sheet.Range["E2"].Formula.ToString();
                if (!formulaE2.Contains("=UPPER(LEFT([@City],3))"))
                    return "False";

                string formulaE20 = sheet.Range["E20"].Formula.ToString();
                if (!formulaE20.Contains("=UPPER(LEFT([@City],3))"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau100(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet = d.Worksheets["Historical Sales"];

                // Kiểm tra xem cửa sổ đang hiển thị có bật chế độ hiện công thức không
                if (!sheet.Application.ActiveWindow.DisplayFormulas)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau101(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet = d.Worksheets["Authors"];

                string formulaD2 = sheet.Range["D2"].Formula.ToString();
                if (!formulaD2.Contains("=IF([@[Books Sold]]>10000,500,100)"))
                    return "False";

                string formulaD37 = sheet.Range["D37"].Formula.ToString();
                if (!formulaD37.Contains("=IF([@[Books Sold]]>10000,500,100)"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau106(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Grade Criteria"];
                string formulaB28 = worksheet.Range["B28"].Formula?.ToString() ?? "";

                if (!formulaB28.Contains("=SUM(Total1,Total2,Total3)") &&
                    !formulaB28.Contains("=Total1+Total2+Total3"))
                {
                    return "False";
                }
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau107(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Exams"];
                string formulaE35 = worksheet.Range["E35"].Formula?.ToString() ?? "";

                if (!formulaE35.Contains("=COUNTBLANK(Table3[Exam 3])"))
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau8(Application a, Workbook d)
        {
            Worksheet worksheet;

            // 1. Lấy worksheet "Orders"
            try
            {
                worksheet = d.Worksheets["Orders"];
            }
            catch
            {
                return "Fales (worksheet)";
            }

            // 2. Lấy điều kiện định dạng có điều kiện trong vùng G2:G526
            FormatConditions formatConditions;
            try
            {
                formatConditions = worksheet.Range["G2", "G526"].FormatConditions;
            }
            catch
            {
                return "False";
            }

            if (formatConditions.Count != 1)
                return "False";

            // 3. Kiểm tra xem điều kiện là "AboveAverage"
            try
            {
                var aboveAverage = formatConditions[1] as AboveAverage;
                if (aboveAverage == null)
                    return "False";

                if (aboveAverage.AboveBelow != XlAboveBelow.xlAboveAverage)
                    return "False";

                string fontColor = aboveAverage.Font.Color?.ToString() ?? "";

                // Kiểm tra xem Font.Color có phải là 24832 không
                if (fontColor != "24832")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau52(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["London"] as Worksheet;
                if (worksheet == null)
                    return "False";

                Range range = worksheet.Range["D5", "D21"];
                FormatConditions formatConditions = range.FormatConditions;

                if (formatConditions == null)
                    return "False";

                return formatConditions.Count != 0
                    ? "False"
                    : "True";
            }
            catch
            {
                return "False";
            }
        }
        private string Cau74(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Products"] as Worksheet;
                if (worksheet == null)
                    return "False";

                Range range = worksheet.Range["E3", "E54"];
                FormatConditions formatConditions = range.FormatConditions;

                if (formatConditions == null)
                    return "False";
                if (formatConditions.Count != 1)
                    return "False";

                IconSetCondition iconSetCondition = formatConditions[1] as IconSetCondition;
                if (iconSetCondition == null)
                    return "False";

                IconSet iconSet = iconSetCondition.IconSet;
                if (iconSet == null)
                    return "False";
                if (iconSet.ID != XlIconSet.xl3TrafficLights1)
                    return "False";
                if (iconSet.Count != 3)
                    return "False";
                if (iconSetCondition.IconCriteria.Count != 3)
                    return "False";

                if (iconSetCondition.IconCriteria[1].Icon != XlIcon.xlIconRedCircleWithBorder)
                    return "False";
                if (iconSetCondition.IconCriteria[2].Icon != XlIcon.xlIconYellowCircle)
                    return "False";
                if (iconSetCondition.IconCriteria[3].Icon != XlIcon.xlIconGreenCircle)
                    return "False";

                if (iconSetCondition.IconCriteria[1].Type != XlConditionValueTypes.xlConditionValuePercent ||
                    iconSetCondition.IconCriteria[2].Type != XlConditionValueTypes.xlConditionValuePercent ||
                    iconSetCondition.IconCriteria[3].Type != XlConditionValueTypes.xlConditionValuePercent)
                    return "False";

                if (iconSetCondition.IconCriteria[1].Operator != 7 ||
                    iconSetCondition.IconCriteria[2].Operator != 7 ||
                    iconSetCondition.IconCriteria[3].Operator != 7)
                    return "False";

                return "True";
            }
            catch
            {
                return "False";
            }
        }
        private string Cau96(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Summary"] as Worksheet;
            }
            catch (Exception)
            {
                return "False";
            }

            FormatConditions formatConditions;
            try
            {
                formatConditions = worksheet.Range["F4", "F11"].FormatConditions;
            }
            catch (Exception)
            {
                return "False";
            }

            if (formatConditions.Count != 1)
                return "False";

            FormatCondition formatCondition;
            try
            {
                formatCondition = formatConditions[1] as FormatCondition;
            }
            catch (Exception)
            {
                return "False";
            }

            // Sửa lỗi kiểu ở đây
            if ((int)formatCondition.Type != (int)XlFormatConditionType.xlCellValue)
                return "False";

            if ((int)formatCondition.Operator != (int)XlFormatConditionOperator.xlGreater)
                return "False";

            if (formatCondition.Formula1 != "=5000000")
                return "False";

            object colorValue;
            try
            {
                colorValue = formatCondition.Font.Color.ToString();
            }
            catch (Exception)
            {
                return "False";
            }

            if (colorValue.ToString() != "22428")
                return "False";

            return "True";
        }
        private string Cau25(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Fiction"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                if (worksheet.Shapes.Count != 1)
                    return "False";

                Microsoft.Office.Interop.Excel.Shape picture = worksheet.Shapes.Item("picture 1");
                if (picture == null)
                    return "False";

                if (picture.Rotation != 0)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau26(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Lấy worksheet "Non_Fiction"
            try
            {
                worksheet = d.Worksheets["Non_Fiction"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                // Kiểm tra số lượng hình vẽ
                if (worksheet.Shapes.Count != 1)
                    return "False";

                var shape = worksheet.Shapes.Item(1);

                // Kiểm tra hiệu ứng hình ảnh
                if (shape.Fill.PictureEffects.Count != 1)
                    return "False";
                if (shape.Fill.PictureEffects[1].Type != mscore.MsoPictureEffectType.msoEffectBackgroundRemoval)
                    return "False";

                // Kiểm tra Pattern Fill
                if (shape.Fill.Pattern != mscore.MsoPatternType.msoPattern20Percent)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau39(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Summary"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                if (worksheet.Shapes.Count != 1)
                    return "False";

                Microsoft.Office.Interop.Excel.Shape shape = worksheet.Shapes.Item(1);
                if (shape.AlternativeText != "Renewal data")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau69(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["October"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                if (worksheet.Shapes.Count != 2)
                    return "False";

                var shape = worksheet.Shapes.Item(2);
                if ((double)shape.Left < 400.0)
                    return "False";
                if ((double)shape.Top > 100.0)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau3(Application a, Workbook d)
        {
            try
            {
                // Truy cập worksheet tên "Classes"
                Worksheet worksheet = d.Worksheets["Classes"] as Worksheet;
                if (worksheet == null)
                    return "False";

                // Truy cập vùng A4:F25
                Range range = worksheet.Range["A4", "F25"];

                // Kiểm tra xem vùng có chứa ListObject (Excel Table) hay không
                if (range.ListObject == null)
                    return "True"; // Không có bảng
                else
                    return "False"; // Có bảng
            }
            catch
            {
                return "False";
            }
        }
        private string Cau19(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Truy cập trang tính "Revenue"
            try
            {
                worksheet = d.Worksheets["Revenue"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            // Bước 2: Kiểm tra ô A3 có trong một bảng hay không
            ListObject a3Object;
            try
            {
                Range rangeA3 = worksheet.Range["A3"];
                a3Object = rangeA3.ListObject;
                if (a3Object == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            // Bước 3: Kiểm tra ô B7 có thuộc bảng, và bảng đó có đúng vùng không
            ListObject listObject;
            try
            {
                Range rangeB7 = worksheet.Range["B7"];
                listObject = rangeB7.ListObject;
                if (listObject == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            // Bước 4: Kiểm tra vùng của bảng có phải là $A$3:$B$7 không
            try
            {
                string address = listObject.Range.get_Address(
                    Missing.Value, Missing.Value,
                    XlReferenceStyle.xlA1,
                    Type.Missing, Type.Missing
                );

                if (address != "$A$3:$B$7")
                    return "False";
            }
            catch
            {
                return "False";
            }

            //return listObject.TableStyle.Name.ToString();

            // Bước 5: Kiểm tra kiểu của bảng có phải là "TableStyleLight14"
            try
            {
                if (listObject.TableStyle != null &&
                    listObject.TableStyle.Name.ToString() != "TableStyleLight14")
                {
                    return "False";
                }
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau20(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Truy cập trang tính "Last Semester"
            try
            {
                worksheet = d.Worksheets["Last Semester"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            // Bước 2: Kiểm tra xem ô B6 có thuộc một bảng không
            try
            {
                Range rangeB6 = worksheet.Range["B6"];
                if (rangeB6.ListObject == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            // Bước 3: So sánh nội dung ô B6
            try
            {
                Range rangeB6 = worksheet.Range["B6"];
                string textB6 = rangeB6.Text.ToString();

                if (textB6 != "Health & Beauty")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau40(Application a, Workbook d)
        {
            Worksheet ws;
            try
            {
                ws = d.Worksheets["March"];
            }
            catch
            {
                return "False";
            }

            ListObject tbl;
            try
            {
                tbl = ws.Range["A4"].ListObject;
            }
            catch
            {
                return "False";
            }

            try
            {
                var filter = tbl.AutoFilter.Filters[6]; // Cột G (vì chỉ số bắt đầu từ 0)
                string criteria = filter.Criteria1?.ToString();
                if (criteria != "=MP")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau46(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["New Accounts"] as Worksheet;
                if (worksheet == null)
                    return "False";

                try
                {
                    var listObject = worksheet.Range["A3"].ListObject;
                }
                catch
                {
                    return "False";
                }

                string cellA6 = worksheet.Range["A6"].Text.ToString();
                if (cellA6 != "Fabrikam, Inc.")
                    return "False";

                return "True";
            }
            catch
            {
                return "False";
            }
        }
        private string Cau53(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["New York City"] as Worksheet;
                if (worksheet == null)
                    return "False";

                string country = worksheet.Range["A7"].Text.ToString();
                if (country != "China")
                    return "False";

                string city = worksheet.Range["B7"].Text.ToString();
                if (city != "Beijing")
                    return "False";

                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau64(Application a, Workbook d)
        {
            Worksheet ws;
            try
            {
                ws = d.Worksheets["New Policies"];
            }
            catch
            {
                return "False";
            }

            ListObject tbl;
            try
            {
                tbl = ws.Range["A4"].ListObject;
            }
            catch
            {
                return "False";
            }

            if (!tbl.ShowTotals)
                return "False";

            string b14 = ws.Range["B14"].Formula?.ToString();
            if (b14 != "=SUBTOTAL(109,[January])")
                return "False";

            string h14 = ws.Range["H14"].Formula?.ToString();
            if (h14 != "=SUBTOTAL(109,[Total])")
                return "False";

            string i14 = ws.Range["I14"].Formula?.ToString();
            if (!string.IsNullOrEmpty(i14))
                return "False";

            string j14 = ws.Range["J14"].Formula?.ToString();
            if (!string.IsNullOrEmpty(j14))
                return "False";

            return "True";
        }
        private string Cau75(Application a, Workbook d)
        {
            Worksheet ws;
            try
            {
                ws = d.Worksheets["Products"];
            }
            catch
            {
                return "False";
            }

            ListObject tbl;
            try
            {
                tbl = ws.Range["A4"].ListObject;
            }
            catch
            {
                return "False";
            }

            try
            {
                if (tbl.TableStyle.Name != "TableStyleMedium1")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau83(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Truy cập trang tính "Orders"
            try
            {
                worksheet = d.Worksheets["Orders"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            // Bước 2: Lấy bảng bắt đầu từ ô A1
            ListObject listObject;
            try
            {
                Range range = worksheet.Range["A1"];
                listObject = range.ListObject;
                if (listObject == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            // Bước 3: Kiểm tra filter trên cột đầu tiên
            try
            {
                if (listObject.AutoFilter == null)
                    return "False";

                Filter filter = listObject.AutoFilter.Filters[1];
                if (!filter.On)
                    return "False";

                string criteria = filter.Criteria1?.ToString();
                if (criteria != "=Alpine Ski House")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau97(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Region 1"] as Worksheet;
                if (worksheet == null)
                    return "False";

                ListObject listObject = worksheet.Range["A3"]?.ListObject;
                if (listObject == null)
                    return "False";

                var sortFields = listObject.Sort.SortFields;
                if (sortFields.Count < 2)
                    return "False";

                var field1 = sortFields[1];
                var field2 = sortFields[2];

                // Sửa ở đây: truyền đúng kiểu XlReferenceStyle
                string address1 = field1?.Key?.get_Address(
                    Type.Missing,
                    Type.Missing,
                    Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1,
                    Type.Missing,
                    Type.Missing
                );

                string address2 = field2?.Key?.get_Address(
                    Type.Missing,
                    Type.Missing,
                    Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1,
                    Type.Missing,
                    Type.Missing
                );

                if (address1 != "$A$4:$A$11")
                    return "False";
                if (field1.Order != XlSortOrder.xlAscending)
                    return "False";

                if (address2 != "$F$4:$F$11")
                    return "False";
                if (field2.Order != XlSortOrder.xlDescending)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau108(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Tasks"] as Worksheet;
                if (worksheet == null)
                    return "False";

                Range cell = worksheet.get_Range("A3", "A3");
                if (cell == null)
                    return "False";

                ListObject listObject = cell.ListObject;
                if (listObject == null)
                    return "False";

                if (!listObject.ShowTableStyleRowStripes)
                    return "False";

                return "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
        private string Cau109(Application a, Workbook d)
        {
            try
            {
                //Worksheet worksheet = null;

                //// Tìm sheet "Tasks"
                //foreach (Worksheet ws in d.Worksheets)
                //{
                //    if (ws.Name == "Tasks")
                //    {
                //        worksheet = ws;
                //        break;
                //    }
                //}

                //if (worksheet == null)
                //    return "False (trang tính Tasks)";

                Worksheet worksheet = d.Worksheets["Tasks"] as Worksheet;
                if (worksheet == null)
                    return "False";

                // Lấy ô A3
                Range cell = worksheet.Range["A3"];
                if (cell == null)
                    return "False";

                // Tìm bảng chứa ô A3
                ListObject listObject = null;
                foreach (ListObject lo in worksheet.ListObjects)
                {
                    Range intersect = worksheet.Application.Intersect(lo.Range, cell);
                    if (intersect != null)
                    {
                        listObject = lo;
                        break;
                    }
                }

                if (listObject == null)
                    return "False";

                // Kiểm tra tên bảng
                if (!string.Equals(listObject.Name, "Tasks", StringComparison.OrdinalIgnoreCase))
                    return "False";

                return "True";
            }
            catch (Exception)
            {
                return "False";
            }
        }
        private string Cau4(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Tìm sheet "Instructional Hours"
            try
            {
                worksheet = (Worksheet)d.Worksheets["Instructional Hours"];
            }
            catch
            {
                return "False";
            }

            // Bước 2: Phải có đúng 1 shape (biểu đồ)
            if (worksheet.Shapes.Count != 1)
                return "False";

            // Bước 3: Lấy biểu đồ từ shape
            Chart chart;
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False";
            }

            // Bước 4: Kiểm tra tiêu đề trục tung (Value Axis)
            try
            {
                Axis yAxis = chart.Axes(XlAxisType.xlValue);
                if (yAxis.AxisTitle.Text != "Hours")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau5(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Kiểm tra tồn tại sheet "Graduation"
            try
            {
                worksheet = (Worksheet)d.Worksheets["Graduation"];
            }
            catch
            {
                return "False";
            }

            // Bước 2: Kiểm tra không còn biểu đồ trong sheet "Graduation"
            if (worksheet.Shapes.Count != 0)
                return "False";

            // Bước 3: Kiểm tra tổng số worksheet là 5
            try
            {
                if (d.Worksheets.Count != 5)
                    return "False";

                // Bước 4: Kiểm tra tồn tại sheet "Graduation Chart"
                object chartSheet = d.Sheets["Graduation Chart"];
                string name = chartSheet.ToString(); // Kiểm tra không lỗi
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau6(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Lấy worksheet "Enrollment"
            try
            {
                worksheet = (Worksheet)d.Worksheets["Enrollment"];
            }
            catch
            {
                return "Fales (Enrollment worksheet not found)";
            }

            try
            {
                Range sparkRange = worksheet.get_Range("G5", "G25");
                SparklineGroups sparklineGroups = sparkRange.SparklineGroups;

                // Bước 2: Kiểm tra có đúng 1 SparklineGroup
                if (sparklineGroups.Count != 1)
                    return "False";

                SparklineGroup group = sparklineGroups[1];

                // Bước 3: Kiểm tra vùng nguồn dữ liệu
                if (group.SourceData != "D5:F25")
                    return "False";

                // Bước 4: Kiểm tra loại Sparkline
                if (group.Type != XlSparkType.xlSparkColumn)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau21(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Next Semester"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            if (worksheet.Shapes.Count != 1)
                return "False";

            Chart chart;
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False";
            }

            try
            {
                // Kiểm tra loại biểu đồ
                if (chart.ChartType != XlChartType.xlColumnClustered)
                    return "False";

                SeriesCollection seriesCollection = chart.SeriesCollection() as SeriesCollection;
                if (seriesCollection == null || seriesCollection.Count != 1)
                    return "False";

                Series series = seriesCollection.Item(1);
                if (series.FormulaR1C1 != "=SERIES('Next Semester'!R3C5,'Next Semester'!R4C1:R21C1,'Next Semester'!R4C5:R21C5,1)")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau22(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Enrollment Summary"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            if (worksheet.Shapes.Count != 1)
                return "False";

            Chart chart;
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False";
            }

            try
            {
                // Kiểm tra ChartStyle
                string style = chart.ChartStyle?.ToString();
                if (style != "268")
                    return "False";

                // Kiểm tra ChartColor
                string color = chart.ChartColor?.ToString();
                if (color != "19")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau47(Application a, Workbook d)
        {
            string result = "True";
            Worksheet worksheet;

            // Bước 1: Kiểm tra sheet "New Accounts" tồn tại
            try
            {
                worksheet = (Worksheet)d.Worksheets["New Accounts"];
            }
            catch
            {
                return "False";
            }

            // Bước 2: Sheet phải có đúng 1 shape
            if (worksheet.Shapes.Count != 1)
                return "False";

            Chart chart;

            // Bước 3: Lấy chart từ shape
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False";
            }

            // Bước 4: Kiểm tra PlotBy có vẽ theo dòng không (phải là theo cột)
            try
            {
                if (chart.PlotBy == XlRowCol.xlRows)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return result;
        }
        private string Cau54(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Kiểm tra có tồn tại sheet "London"
            try
            {
                worksheet = (Worksheet)d.Worksheets["London"];
            }
            catch
            {
                return "False";
            }

            // Bước 2: Sheet phải có đúng 2 shape
            if (worksheet.Shapes.Count != 2)
                return "False";

            Chart chart;

            // Bước 3: Lấy biểu đồ từ shape thứ 2
            try
            {
                Shape shape = worksheet.Shapes.Item(2);
                chart = shape.Chart;
            }
            catch
            {
                return "False";
            }

            // Bước 4: Kiểm tra DataTable không hiển thị LegendKey
            try
            {
                if (chart.DataTable.ShowLegendKey)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau65(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["New Policies"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            if (worksheet.Shapes.Count != 1)
                return "False";

            Chart chart;
            try
            {
                chart = worksheet.Shapes.Item(1).Chart;
            }
            catch
            {
                return "False";
            }

            try
            {
                if (chart.DataTable != null && chart.DataTable.ShowLegendKey)
                    return "False";
            }
            catch
            {
                // Có thể biểu đồ không có DataTable → kiểm tra legend thay thế
                try
                {
                    if (chart.Legend == null || chart.Legend.Position != XlLegendPosition.xlLegendPositionBottom)
                        return "False";
                }
                catch
                {
                    return "False";
                }

                return "True";
            }

            return "False";
        }
        private string Cau66(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["New Policies"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                Range sparkRange = worksheet.Range["J5", "J13"];
                var sparkGroups = sparkRange.SparklineGroups;

                if (sparkGroups.Count != 1)
                    return "False";

                var sparkGroup = sparkGroups[1];

                if (sparkGroup.SourceData != "B5:G13")
                    return "False";

                if (sparkGroup.Type != XlSparkType.xlSparkColumnStacked100)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau76(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Summary"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            if (worksheet.Shapes.Count != 1)
                return "False";

            try
            {
                Chart chart = worksheet.Shapes.Item(1).Chart;

                string chartColor = chart.ChartColor?.ToString();
                if (chartColor != "11")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau84(Application a, Workbook d)
        {
            string result = "True";
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Inventory"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            if (worksheet.Shapes.Count != 1)
                return "False";

            Chart chart;
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False";
            }

            try
            {
                if (chart.ChartTitle.Top != 2.0)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                Series series = chart.SeriesCollection(1) as Series;
                if (series == null || !series.HasDataLabels)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return result;
        }
        private string Cau98(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Comparison"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            if (worksheet.Shapes.Count != 1)
                return "False";

            try
            {
                Chart chart = worksheet.Shapes.Item(1).Chart;
                if (chart.PlotBy != XlRowCol.xlRows)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau102(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Lấy worksheet "Summary"
            try
            {
                worksheet = (Worksheet)d.Worksheets["Summary"];
            }
            catch
            {
                return "False";
            }

            // Bước 2: Kiểm tra số shape
            if (worksheet.Shapes.Count != 1)
                return "False";

            Chart chart;

            // Bước 3: Lấy chart từ shape
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False";
            }

            try
            {
                // Bước 4: Kiểm tra loại biểu đồ
                if (chart.ChartType != XlChartType.xlColumnClustered)
                    return "False";

                // Bước 5: Lấy series và kiểm tra số lượng
                SeriesCollection seriesCollection = (SeriesCollection)chart.SeriesCollection(Type.Missing);
                if (seriesCollection.Count != 2)
                    return "False";

                // Bước 6: Kiểm tra công thức của series thứ 2
                string expectedFormula = "=SERIES(Summary!R5C3,Summary!R6C1:R12C1,Summary!R6C3:R12C3,2)";
                if (seriesCollection.Item(2).FormulaR1C1 != expectedFormula)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau110(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Score Distribution"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            if (worksheet.Shapes.Count != 1)
                return "False";

            Chart chart;
            try
            {
                chart = worksheet.Shapes.Item(1).Chart;
            }
            catch
            {
                return "False";
            }

            try
            {
                // Kiểm tra nếu có Legend thì sai
                if (chart.Legend != null)
                    return "False";
            }
            catch
            {
                // Nếu có lỗi khi kiểm tra Legend, kiểm tra Series[1] có bật DataLabels không
                try
                {
                    Series series = chart.SeriesCollection(1);
                    if (!series.HasDataLabels)
                        return "False";
                }
                catch
                {
                    return "False";
                }
            }

            return "True";
        }
        private string Cau34(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Materials"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                Range cell = worksheet.Range["A6"];
                Hyperlinks hyperlinks = cell.Hyperlinks;

                if (hyperlinks.Count != 1)
                    return "False";

                Hyperlink link = hyperlinks[1];
                if (link.SubAddress != "Categories!A18")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau57(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Materials"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                Range cell = worksheet.Range["A6"];
                Hyperlinks hyperlinks = cell.Hyperlinks;

                if (hyperlinks.Count != 1)
                    return "False";

                Hyperlink link = hyperlinks[1];
                if (link.SubAddress != "Categories!A18")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau58(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Materials"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                if (worksheet.PageSetup.RightHeader != "Confidential")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau59(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Materials"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                if (worksheet.PageSetup.CenterFooter != "Page &P of &N")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau103(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Summary"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                Range cell = worksheet.Range["A2"];
                Hyperlinks hyperlinks = cell.Hyperlinks;

                if (hyperlinks.Count != 1)
                    return "False";

                Hyperlink link = hyperlinks[1];

                if (!link.Address.Contains("www.nodpublishers.com"))
                    return "False";

                if (link.ScreenTip != "Company Website")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau27(Application a, Workbook d)
        {
            try
            {
                var props = d.BuiltinDocumentProperties;
                var companyProp = props["Company"];
                string companyName = companyProp.GetType()
                    .InvokeMember("Value", BindingFlags.GetProperty, null, companyProp, new object[0])
                    ?.ToString();

                if (companyName != "Lucerne Publishing")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau35(Application a, Workbook d)
        {
            try
            {
                foreach (Worksheet ws in d.Worksheets)
                {
                    try
                    {
                        var ps = ws.PageSetup;

                        // Kiểm tra Zoom phải là False
                        if (ps.Zoom != false)
                            return "False";

                        // Kiểm tra FitToPagesWide == 1
                        if (ps.FitToPagesWide != 1)
                            return "False";

                        // Kiểm tra FitToPagesTall == 1
                        if (ps.FitToPagesTall != 1)
                            return "False";
                    }
                    catch
                    {
                        return "False";
                    }
                }
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau36(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Games"] as Worksheet;
                if (worksheet == null)
                    return "Fales (ten trang tinh)";

                if (worksheet.PageSetup.TopMargin != 72.0)
                    return "False";
                if (worksheet.PageSetup.BottomMargin != 72.0)
                    return "False";
                if (worksheet.PageSetup.LeftMargin != 108.0)
                    return "False";
                if (worksheet.PageSetup.RightMargin != 108.0)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau41(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                // Truy cập worksheet "January"
                worksheet = (Worksheet)d.Worksheets["January"];
            }
            catch
            {
                return "False";
            }

            try
            {
                // Kiểm tra vùng in
                if (worksheet.PageSetup.PrintArea != "$A$4:$F$20")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau60(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Materials"] as Worksheet;
                if (worksheet == null)
                    return "False";

                var pageSetup = worksheet.PageSetup;

                //return pageSetup.Orientation.ToString() + XlPageOrientation.xlLandscape.ToString();

                if (pageSetup.Orientation.ToString() != XlPageOrientation.xlLandscape.ToString())
                    return "False";

                //return pageSetup.FitToPagesWide.ToString() + pageSetup.FitToPagesTall.ToString();

                if (pageSetup.FitToPagesWide != 1)
                    return "False";

                if (pageSetup.FitToPagesTall == true)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau77(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Q2 Sales"] as Worksheet;
                if (worksheet == null)
                    return "False";

                if (worksheet.PageSetup.BottomMargin != 54.0 || worksheet.PageSetup.LeftMargin != 18.0)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau91(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["roster"] as Worksheet;
                if (worksheet == null)
                    return "False";

                if (worksheet.PageSetup.PrintTitleRows != "$7:$7")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }
        private string Cau93(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Inbound call"];

                if (worksheet.PageSetup.PrintArea != "$A$1:$C$19")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
        private string Cau104(Application a, Workbook d)
        {
            try
            {
                // Lấy danh sách các thuộc tính tài liệu tích hợp sẵn (Built-in)
                object documentProperties = d.BuiltinDocumentProperties;

                // Lấy thuộc tính "Subject"
                object subjectProperty = documentProperties.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty,
                    null,
                    documentProperties,
                    new object[] { "Subject" });

                // Lấy giá trị của "Subject"
                string subjectValue = subjectProperty.GetType().InvokeMember(
                    "Value",
                    BindingFlags.GetProperty,
                    null,
                    subjectProperty,
                    new object[0])?.ToString();

                // Nếu Subject có giá trị khác rỗng, trả về False
                if (!string.IsNullOrEmpty(subjectValue))
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private string Cau111(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["New York City"] as Worksheet;
                if (ws == null)
                    return "False";

                ChartObjects chartObjects = (ChartObjects)ws.ChartObjects();
                if (chartObjects.Count == 0)
                    return "False";

                //bool correct = false;

                Range dataRange = ws.Range["A3"].CurrentRegion; // vùng bảng
                double tableBottom = dataRange.Top + dataRange.Height;

                // ✅ Phải có 1 biểu đồ
                if (ws.ChartObjects().Count == 0)
                    return "False";

                ChartObject chartObj = ws.ChartObjects(1);
                Chart chart = chartObj.Chart;

                // ✅ Biểu đồ phải là Clustered Column
                if (chart.ChartType != XlChartType.xlColumnClustered)
                    return "False";

                // ✅ Kiểm tra Series
                if (chart.SeriesCollection().Count == 0)
                    return "False";

                //Series series = chart.SeriesCollection(1);

                //// ✅ X-axis = City (cột B)
                //Range xRange = series.XValues as Range;
                //if (xRange.Column != 2)
                //    return "False (Trục ngang không phải City)";

                //// ✅ Values = Air Miles (cột D)
                //Range valRange = series.Values as Range;
                //if (valRange.Column != 4)
                //    return "False (Dữ liệu không phải Air Miles)";

                // ✅ Biểu đồ phải nằm *dưới bảng*
                if (chartObj.Top <= tableBottom)
                    return "False";

                return "True";

                //foreach (ChartObject chObj in chartObjects)
                //{
                //    Chart chart = chObj.Chart;

                //    // 1️⃣ Loại biểu đồ phải là Clustered Column
                //    if (chart.ChartType != XlChartType.xlColumnClustered)
                //        continue;

                //    Series series = chart.SeriesCollection(1);
                //    if (series == null)
                //        continue;

                //    // 2️⃣ Kiểm tra tên Series chứa "Air Miles"
                //    string sName = series.Name.ToLower();
                //    if (!sName.Contains("air") || !sName.Contains("mile"))
                //        continue;

                //    // 3️⃣ Check XValues must come from column B (City)
                //    string xRange = series.XValues as string;
                //    if (string.IsNullOrEmpty(xRange))
                //        continue;

                //    // Ví dụ địa chỉ sẽ kiểu như: ='New York City'!$B$3:$B$17
                //    if (!xRange.ToUpper().Contains("$B$"))
                //        continue;

                //    if (chartObjects.Top <= tableBottom)
                //        return "False (Biểu đồ không nằm dưới bảng)";


                //    // ✅ Passed all checks
                //    correct = true;
                //    break;
                //}

                //return correct ? "True" : "False (Biểu đồ chưa đúng dữ liệu)";
            }
            catch (Exception ex)
            {
                return "False";
            }

        }
    }
}
