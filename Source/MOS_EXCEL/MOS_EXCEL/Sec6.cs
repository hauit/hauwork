// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Sec6
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
  internal class Sec6
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec6.cau12(a, d);
        case 2:
          return Sec6.cau16(a, d);
        case 3:
          return Sec6.cau9(a, d);
        case 4:
          return Sec6.cau13(a, d);
        case 5:
          return Sec6.cau19(a, d);
        case 6:
          return Sec6.cau14(a, d);
        case 7:
          return Sec6.cau17(a, d);
        case 8:
          return Sec6.cau8(a, d);
        case 9:
          return Sec6.cau11(a, d);
        case 10:
          return Sec6.cau15(a, d);
        case 11:
          return Sec6.cau18(a, d);
        case 12:
          return Sec6.cau10(a, d);
        case 13:
          return Sec6.cau13(a, d);
        case 14:
          return Sec6.cau14(a, d);
        case 15:
          return Sec6.cau15(a, d);
        case 16:
          return Sec6.cau16(a, d);
        case 17:
          return Sec6.cau17(a, d);
        case 18:
          return Sec6.cau18(a, d);
        case 19:
          return Sec6.cau19(a, d);
        default:
          return "False";
      }
    }

        private static string cau1(Application app, Workbook workbook)
        {
            try
            {
                // 1. Tìm worksheet có tên "Customers"
                Worksheet worksheet = workbook.Worksheets["Customers"] as Worksheet;
                if (worksheet == null)
                    return "False (ten trang tinh)";

                // 2. Lấy bảng đầu tiên trong worksheet
                if (worksheet.ListObjects.Count < 1)
                    return "False (Table)";

                ListObject table = worksheet.ListObjects[1];

                // 3. Kiểm tra banded rows (hàng xen kẽ màu)
                if (!table.ShowTableStyleRowStripes)
                    return "False (banded rows)";

                return "True";
            }
            catch
            {
                return "False (Exception)";
            }
        }

        private static string cau2(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["Customers"];
                if (sheet == null)
                    return "False (ten trang tinh)";

                // Lấy dữ liệu các ô cần kiểm tra
                string I499 = sheet.Range["I499"].Text.ToString();
                string I500 = sheet.Range["I500"].Text.ToString();
                string I501 = sheet.Range["I501"].Text.ToString();

                string G499 = sheet.Range["G499"].Text.ToString();
                string G2 = sheet.Range["G2"].Text.ToString();

                string H2 = sheet.Range["H2"].Text.ToString();
                string H4 = sheet.Range["H4"].Text.ToString();

                // Kiểm tra điều kiện sort Level 1
                if (I499 != "Canada" || I500 != "Canada" || I501 != "Canada")
                    return "False(sort sai Level 1)";

                // Kiểm tra điều kiện sort Level 2
                if (G499 != "QC" || G2 != "AK")
                    return "False(sort sai Level 2)";

                // Kiểm tra điều kiện sort Level 3
                if (H2 != "16202" || H4 != "64577")
                    return "False(sort sai Level 3)";

                return "True";
            }
            catch
            {
                return "False (ten trang tinh)";
            }
        }

        private static string cau3(Application app, Workbook wb)
        {
            try
            {
                Worksheet ws = wb.Worksheets["Inbound call"] as Worksheet;
                if (ws == null)
                    return "False (ten trang tinh)";

                ListObject table = ws.ListObjects[1];
                if (table == null)
                    return "False (Table)";

                if (!table.ShowTotals)
                    return "False (Totals)";

                return "True";
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return "False (COM Error)";
            }
            catch (Exception)
            {
                return "False (Exception)";
            }
        }

        private static string cau4(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Kiểm tra sheet "Donor List" có tồn tại không
            try
            {
                worksheet = d.Worksheets["Donor List"] as Worksheet;
                if (worksheet == null)
                    return "False (ten trang tinh)";
            }
            catch
            {
                return "False (ten trang tinh)";
            }

            // Bước 2: Kiểm tra xem có Table nào trong sheet không
            try
            {
                if (worksheet.ListObjects.Count > 0)
                    return "False (convert table to range)";
            }
            catch
            {
                return "False (something wrong)";
            }

            return "True";
        }

        private static string cau5(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Donor List"] as Worksheet;
                if (worksheet == null)
                    return "False (ten trang tinh)";
            }
            catch (Exception)
            {
                return "False (ten trang tinh)";
            }

            try
            {
                if (worksheet.ListObjects.Count != 2)
                    return "False (chuyen tu day o sang table)";

                ListObject listObject = worksheet.ListObjects[2]; // index 2 là bảng thứ hai (1-based index)

                if (listObject.ListRows.Count != 4)
                    return "False (check vao table has headerRow khi insert table)";
            }
            catch (Exception)
            {
                return "False (something wrong)";
            }

            return "True";
        }

        private static string cau6(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Lấy worksheet "Demographics"
            try
            {
                worksheet = d.Worksheets["Demographics"] as Worksheet;
                if (worksheet == null)
                    return "False (Tên trang tính)";
            }
            catch
            {
                return "False (Tên trang tính)";
            }

            try
            {
                Range cell = worksheet.Range["C5"];
                string value = cell.Value?.ToString().Trim();

                if (value != "300")
                    return "False (C5)";
            }
            catch
            {
                return "False (Lỗi khi đọc ô C5)";
            }

            return "True";
        }

        private static string cau7(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Donor List"] as Worksheet;
                if (worksheet == null)
                    return "False (Tên trang tính)";

                ListObject listObject = worksheet.ListObjects["Table2"];
                if (listObject == null)
                    return "False (Tên bảng)";

                string altText = listObject.AlternativeText;
                if (string.IsNullOrWhiteSpace(altText))
                    return "False (Chưa thêm alt text)";

                if (altText != "Donor")
                    return "False (Donor - Chuột phải chọn bảng → Alt Text)";

                return "True";
            }
            catch (Exception ex)
            {
                return "False (Lỗi không xác định)";
            }
        }

        private static string cau8(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["New York City"] as Worksheet;
                if (worksheet == null)
                    return "False (Tên trang tính)";

                string country = worksheet.Range["A7"].Text.ToString();
                if (country != "China")
                    return "False (Sai ở sort cấp 1)";

                string city = worksheet.Range["B7"].Text.ToString();
                if (city != "Beijing")
                    return "False (Sai ở sort cấp 2)";

                return "True";
            }
            catch (Exception ex)
            {
                return "False (Lỗi không xác định)";
            }
        }

        private static string cau9(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["New Accounts"] as Worksheet;
                if (worksheet == null)
                    return "False (Trang tính 'New Accounts' không tồn tại)";

                try
                {
                    var listObject = worksheet.Range["A3"].ListObject;
                }
                catch
                {
                    return "False (Table bị chỉnh sửa hoặc không tồn tại tại A3)";
                }

                string cellA6 = worksheet.Range["A6"].Text.ToString();
                if (cellA6 != "Fabrikam, Inc.")
                    return "False";

                return "True";
            }
            catch
            {
                return "False (Lỗi không xác định)";
            }
        }

        private static string cau10(Application excelApp, Workbook workbook)
        {
            try
            {
                // Truy cập worksheet tên "Classes"
                Worksheet worksheet = workbook.Worksheets["Classes"] as Worksheet;
                if (worksheet == null)
                    return "Fales (trang tính Classes)";

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
                return "Fales (trang tính Classes)";
            }
        }

        private static string cau11(Application excelApp, Workbook workbook)
        {
            Worksheet worksheet;

            // Bước 1: Truy cập trang tính "Orders"
            try
            {
                worksheet = workbook.Worksheets["Orders"] as Worksheet;
                if (worksheet == null)
                    return "Fales (trang tính Orders)";
            }
            catch
            {
                return "Fales (trang tính Orders)";
            }

            // Bước 2: Lấy bảng bắt đầu từ ô A1
            ListObject listObject;
            try
            {
                Range range = worksheet.Range["A1"];
                listObject = range.ListObject;
                if (listObject == null)
                    return "False (Table was modify)";
            }
            catch
            {
                return "False (Table was modify)";
            }

            // Bước 3: Kiểm tra filter trên cột đầu tiên
            try
            {
                if (listObject.AutoFilter == null)
                    return "False (Chưa áp dụng filter)";

                Filter filter = listObject.AutoFilter.Filters[1];
                if (!filter.On)
                    return "False(filter trên cột 1 chưa bật)";

                string criteria = filter.Criteria1?.ToString();
                if (criteria != "=Alpine Ski House")
                    return "False(filter cột 1 không chọn Alpine Ski House)";
            }
            catch
            {
                return "False()";
            }

            return "True";
        }

        private static string cau12(Application app, Workbook workbook)
        {
            Worksheet worksheet;

            // Bước 1: Truy cập trang tính "Revenue"
            try
            {
                worksheet = workbook.Worksheets["Revenue"] as Worksheet;
                if (worksheet == null)
                    return "False (trang tính Revenue)";
            }
            catch
            {
                return "False (trang tính Revenue)";
            }

            // Bước 2: Kiểm tra ô A3 có trong một bảng hay không
            ListObject a3Object;
            try
            {
                Range rangeA3 = worksheet.Range["A3"];
                a3Object = rangeA3.ListObject;
                if (a3Object == null)
                    return "False (A3)";
            }
            catch
            {
                return "False (A3)";
            }

            // Bước 3: Kiểm tra ô B7 có thuộc bảng, và bảng đó có đúng vùng không
            ListObject listObject;
            try
            {
                Range rangeB7 = worksheet.Range["B7"];
                listObject = rangeB7.ListObject;
                if (listObject == null)
                    return "False (B7)";
            }
            catch
            {
                return "False (B7)";
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
                    return $"False({address})";
            }
            catch
            {
                return "False(chưa chuyển sang table)";
            }

            //return listObject.TableStyle.Name.ToString();

            // Bước 5: Kiểm tra kiểu của bảng có phải là "TableStyleLight14"
            try
            {
                if (listObject.TableStyle != null &&
                    listObject.TableStyle.Name.ToString() != "TableStyleLight14")
                {
                    return "False(sai kiểu)";
                }
            }
            catch
            {
                return "False(không xác định liên quan đến Kiểu)";
            }

            return "True";
        }

        private static string cau13(Application app, Workbook workbook)
        {
            Worksheet worksheet;

            // Bước 1: Truy cập trang tính "Last Semester"
            try
            {
                worksheet = workbook.Worksheets["Last Semester"] as Worksheet;
                if (worksheet == null)
                    return "Fales (trang tính Last Semester)";
            }
            catch
            {
                return "Fales (trang tính Last Semester)";
            }

            // Bước 2: Kiểm tra xem ô B6 có thuộc một bảng không
            try
            {
                Range rangeB6 = worksheet.Range["B6"];
                if (rangeB6.ListObject == null)
                    return "False (Table was modify)";
            }
            catch
            {
                return "False (Table was modify)";
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

        private static string cau14(Application app, Workbook wb)
        {
            Worksheet ws;
            try
            {
                ws = wb.Worksheets["New Policies"];
            }
            catch
            {
                return "False (trang tính New Policies)";
            }

            ListObject tbl;
            try
            {
                tbl = ws.Range["A4"].ListObject;
            }
            catch
            {
                return "False (Table was modified)";
            }

            if (!tbl.ShowTotals)
                return "False (show dòng tổng)";

            string b14 = ws.Range["B14"].Formula?.ToString();
            if (b14 != "=SUBTOTAL(109,[January])")
                return "False (B14)";

            string h14 = ws.Range["H14"].Formula?.ToString();
            if (h14 != "=SUBTOTAL(109,[Total])")
                return "False (H14)";

            string i14 = ws.Range["I14"].Formula?.ToString();
            if (!string.IsNullOrEmpty(i14))
                return "False (I14)";

            string j14 = ws.Range["J14"].Formula?.ToString();
            if (!string.IsNullOrEmpty(j14))
                return "False (J14)";

            return "True";
        }

        private static string cau15(Application app, Workbook wb)
        {
            Worksheet ws;
            try
            {
                ws = wb.Worksheets["March"];
            }
            catch
            {
                return "False (trang tính March)";
            }

            ListObject tbl;
            try
            {
                tbl = ws.Range["A4"].ListObject;
            }
            catch
            {
                return "False (Table was modify)";
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
                return "False (không xác định)";
            }

            return "True";
        }

        private static string cau16(Application app, Workbook wb)
        {
            Worksheet ws;
            try
            {
                ws = wb.Worksheets["Products"];
            }
            catch
            {
                return "False (trang tính Products)";
            }

            ListObject tbl;
            try
            {
                tbl = ws.Range["A4"].ListObject;
            }
            catch
            {
                return "False (Table was modify)";
            }

            try
            {
                if (tbl.TableStyle.Name != "TableStyleMedium1")
                    return "False (sai style)";
            }
            catch
            {
                return "False (style không xác định)";
            }

            return "True";
        }

        private static string cau17(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Region 1"] as Worksheet;
                if (worksheet == null)
                    return "False (trang tính Region 1)";

                ListObject listObject = worksheet.Range["A3"]?.ListObject;
                if (listObject == null)
                    return "False (Table was modify)";

                var sortFields = listObject.Sort.SortFields;
                if (sortFields.Count < 2)
                    return "False(chưa sort đủ 2 trường cùng lúc)";

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
                    return "False(Product)";
                if (field1.Order != XlSortOrder.xlAscending)
                    return "False(Product->A to Z)";

                if (address2 != "$F$4:$F$11")
                    return "False(Total Sales)";
                if (field2.Order != XlSortOrder.xlDescending)
                    return "False(Total Sales->lớn đến nhỏ)";
            }
            catch
            {
                return "False(chưa sort đủ 2 trường cùng lúc)";
            }

            return "True";
        }

        private static string cau19(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Tasks"] as Worksheet;
                if (worksheet == null)
                    return "False (trang tính Tasks)";

                Range cell = worksheet.get_Range("A3", "A3");
                if (cell == null)
                    return "False (ô A3 không tồn tại)";

                ListObject listObject = cell.ListObject;
                if (listObject == null)
                    return "False (Table không tồn tại ở A3)";

                if (!listObject.ShowTableStyleRowStripes)
                    return "False (banded rows)";

                return "True";
            }
            catch (Exception ex)
            {
                return "False (" + ex.Message + ")";
            }
        }

        private static string cau18(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = null;

                // Tìm sheet "Tasks"
                foreach (Worksheet ws in d.Worksheets)
                {
                    if (ws.Name == "Tasks")
                    {
                        worksheet = ws;
                        break;
                    }
                }

                if (worksheet == null)
                    return "False (trang tính Tasks)";

                // Lấy ô A3
                Range cell = worksheet.Range["A3"];
                if (cell == null)
                    return "False (không tìm thấy ô A3)";

                // Lấy ListObject từ ô A3
                ListObject listObject = null;
                foreach (ListObject lo in worksheet.ListObjects)
                {
                    if (lo.Range.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1) == cell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1))
                    {
                        listObject = lo;
                        break;
                    }
                }

                if (listObject == null)
                    return "False (Table was modified)";

                // Kiểm tra tên bảng
                if (listObject.Name != "Tasks")
                    return "False (tên table)";

                return "True";
            }
            catch (Exception)
            {
                return "False (lỗi không xác định)";
            }
        }
    }
}
