// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Sec1
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
  internal class Sec1
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec1.cau3(a, d);
        case 2:
          return Sec1.cau5(a, d);
        case 3:
          return Sec1.cau6(a, d);
        case 4:
          return Sec1.cau11(a, d);
        case 5:
          return Sec1.cau5_5(a, d);
        case 6:
          return Sec1.cau1(a, d);
        case 7:
          return Sec1.cau10(a, d);
        case 8:
          return Sec1.cau2(a, d);
        case 9:
          return Sec1.cau4(a, d);
        case 10:
          return Sec1.cau8(a, d);
        case 11:
          return Sec1.cau9(a, d);
        case 12:
          return Sec1.goto001(a, d);
        case 13:
          return Sec1.goto002(a, d);
        default:
          return "False";
      }
    }

        private static string cau1(Application excelApp, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Customers by Order"] as Worksheet;
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
                return "False (Something not finish!)";
            }
        }

        private static string cau2(Application excelApp, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Social Media Ads"] as Worksheet;
                if (worksheet == null)
                    return "False";

                // Lấy giá trị ô A17
                Range cellA17 = worksheet.Range["A17"];
                string valueA17 = cellA17?.Text?.ToString()?.Trim() ?? "";

                if (!string.Equals(valueA17, "RANCH", StringComparison.OrdinalIgnoreCase))
                    return "False (comma)";

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
                Worksheet worksheet = workbook.Worksheets["Top Donors"] as Worksheet;
                if (worksheet == null)
                    return "False";

                // Kiểm tra ô A2 có phải là "Mark Bebbington" hay không
                string cellA2 = worksheet.Range["A2"].Text?.ToString()?.Trim();
                if (!string.Equals(cellA2, "Mark Bebbington", StringComparison.OrdinalIgnoreCase))
                    return "False (A2)";

                // Kiểm tra ô C6 có phải là "Platinum" hay không
                string cellC6 = worksheet.Range["C6"].Text?.ToString()?.Trim();
                if (!string.Equals(cellC6, "Platinum", StringComparison.OrdinalIgnoreCase))
                    return "False (C6)";

                return "True";
            }
            catch
            {
                return "False (Something not finish!)";
            }
        }

        private static string cau4(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = (Worksheet)d.Worksheets["Top 25"];

                string cellA2 = worksheet.get_Range("A2", "A2").Text.ToString();
                if (cellA2 != "Rank")
                    return "False (A2)";

                string cellG27 = worksheet.get_Range("G27", "G27").Text.ToString();
                if (cellG27 != "London")
                    return "False (G27)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau5(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Top Donors"] as Worksheet;

                string a2 = ws.Range["A2"].Text.ToString();
                if (a2 != "Daniel P. Taylor")
                    return "False (A2)";

                string c2 = ws.Range["C2"].Text.ToString();
                if (c2 != "Charles Fitzgerald")
                    return "False (C2)";
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau6(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Donor Contact Info"] as Worksheet;

                string f2 = ws.Range["F2"].Text.ToString();
                string f3 = ws.Range["F3"].Text.ToString();
                string h46 = ws.Range["H46"].Text.ToString();

                if (f2 != "64043" || f3 != "64043")
                    return "False(thay thế tất cả)";

                if (h46 != "314-555-64044")
                    return "False(Chỉ thay thế trên cột F)";
            }
            catch
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        public static string Cau7(Workbook workbook)
        {
            // Lấy sheet "Donor List"
            Worksheet worksheet = workbook.Worksheets["Donor List"];

            // Kiểm tra nội dung các ô C16, C27, C67
            if (worksheet.Range["C16"].Text.ToString() != "Bronze Total" ||
                worksheet.Range["C27"].Text.ToString() != "Silver Total" ||
                worksheet.Range["C67"].Text.ToString() != "Platinum Total")
            {
                return "False(chọn group trên Annual Donations Hàm Sum)";
            }

            // Kiểm tra giá trị ô B67
            string cellB67 = worksheet.Range["B67"].Text.ToString();
            if (cellB67 != "$38,021.00 ")
            {
                return "False(Sum trên Donation Level)";
            }

            // Kiểm tra số lượng page break ngang
            if (worksheet.HPageBreaks.Count != 4)
            {
                return "False(check vào Page break between group)";
            }

            return "True";
        }

        private static string cau8(Application excelApp, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Donor List"];

                string c16 = worksheet.Range["C16"].Text?.ToString();
                if (c16 != "Bronze Count")
                    return "False(chọn group trên Annual Donations hàm count)";

                string c27 = worksheet.Range["C27"].Text?.ToString();
                if (c27 != "Silver Count")
                    return "False(chọn group trên Annual Donations hàm count)";

                string c67 = worksheet.Range["C67"].Text?.ToString();
                if (c67 != "Platinum Count")
                    return "False(chọn group trên Annual Donations hàm count)";

                string a67 = worksheet.Range["A67"].Text?.ToString();
                if (a67 != "28")
                    return "False(Count trên Name)";

                if (worksheet.HPageBreaks.Count != 4)
                    return "False(check vào Page break between group)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau9(Application a, Workbook d)
        {
            try
            {
                Worksheet sheet1 = a.ActiveWorkbook.Sheets["Sheet1"] as Worksheet;
                if (sheet1 == null)
                    return "";

                Range cell = sheet1.get_Range("B5");
                if (cell == null || cell.Value2 == null)
                    return "";

                return cell.Value2.ToString();
            }
            catch
            {
                return "";
            }
        }

        private static string cau10(Application a, Workbook d)
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
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string cau11(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Quote"];

                string e12 = ws.Range["E12", "E12"].Text.ToString();
                if (e12 != "Quantity")
                    return "False (E12)";

                string f12 = ws.Range["F12", "F12"].Text.ToString();
                if (f12 != "Unit Price")
                    return "False (F12)";

                string f30 = ws.Range["F30", "F30"].Text.ToString();
                if (f30 != "Total")
                    return "False (F30)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau5_5(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["Quote"];

                string e12 = sheet.Range["E12"].Text?.ToString().Trim();
                if (e12 != "Quantity")
                    return "False (E12)";

                string f12 = sheet.Range["F12"].Text?.ToString();
                if (f12 != "Unit Price ") // có dấu cách ở cuối
                    return "False (F12)";

                string f30 = sheet.Range["F30"].Text?.ToString().Trim();
                if (f30 != "Total")
                    return "False (F30)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }

        private static string goto001(Application a, Workbook d)
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
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string goto002(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Demographics"];
            }
            catch (Exception)
            {
                return "Fales (Ten trang tinh)"; // Có thể bạn muốn viết là "False"?
            }

            try
            {
                var value = worksheet.Range["C5", "C5"].Value;

                string cellValue = value != null ? value.ToString() : "";

                if (cellValue != "300")
                {
                    return "False(C5)";
                }
            }
            catch (Exception)
            {
                return "False (something wrong)";
            }

            return "True";
        }

    }
}
