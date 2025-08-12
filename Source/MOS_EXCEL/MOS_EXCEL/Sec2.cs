// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Sec2
// Assembly: MOS_EXCEL_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace MOS_EXCEL_LEARN
{
  internal class Sec2
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec2.cau6(a, d);
        case 2:
          return Sec2.cau7(a, d);
        case 3:
          return Sec2.cau9(a, d);
        case 4:
          return Sec2.cau2(a, d);
        case 5:
          return Sec2.cau4(a, d);
        case 6:
          return Sec2.cau6(a, d);
        case 7:
          return Sec2.cau7(a, d);
        case 8:
          return Sec2.cau8(a, d);
        case 9:
          return Sec2.cau9(a, d);
        default:
          return "False";
      }
    }

        private static string cau1(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Products"] as Worksheet;
                if (worksheet == null)
                    return "False (Không tìm thấy sheet 'Products')";

                Range range = worksheet.get_Range("D2", "D32");
                if (range == null)
                    return "False (Không lấy được vùng D2:D32)";

                string numberFormat = range.NumberFormat?.ToString();
                if (numberFormat != "0.000")
                    return "False(chuỗi định dạng phải là: 0.000)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau2(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Materials"] as Worksheet;
                if (worksheet == null)
                    return "False (Không tìm thấy sheet 'Materials')";

                Range range = worksheet.get_Range("A1", "N1");
                if (range == null)
                    return "False (Không lấy được vùng A1:N1)";

                // Kiểm tra MergeCells
                if (!(range.MergeCells is bool merged) || !merged)
                    return "False(MergeCell)";

                // Kiểm tra canh lề ngang (1 = xlLeft)
                if (!(range.HorizontalAlignment is int alignment) || alignment != 1)
                    return "False(không thay đổi canh lề)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau3(Application app, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Materials"];
                if (worksheet == null)
                    return "False";

                string[] columns = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N" };

                foreach (string col in columns)
                {
                    double width = (double)(worksheet.Range[$"{col}4"].Width);
                    // Kiểm tra nếu KHÔNG nằm trong khoảng [159, 161]
                    if (width < 159.0 || width > 161.0)
                        return "False";
                }
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau4(Application app, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Games"];
                if (worksheet == null)
                    return "False";

                // Nếu cả vùng A12:B18 bị merge lại thì sai
                Range mergedRange = worksheet.Range["A12", "B18"];
                if ((bool)mergedRange.MergeCells == true)
                    return "False(chọn merge cross)";

                // Nếu A12:B12 không được merge thì sai
                Range topRow = worksheet.Range["A12", "B12"];
                if ((bool)topRow.MergeCells != true)
                    return "False(chọn merge cross)";

                // Nếu A18:B18 không được merge thì sai
                Range bottomRow = worksheet.Range["A18", "B18"];
                if ((bool)bottomRow.MergeCells != true)
                    return "False(chọn merge cross)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau5(Application app, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Materials"];
                if (worksheet == null)
                    return "False";

                Range cellA2 = worksheet.Range["A2"];
                if (!(bool)cellA2.WrapText)
                    return "False(dong 2)";

                Range cellA3 = worksheet.Range["A3"];
                if (!(bool)cellA3.WrapText)
                    return "False(dong 3)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau6(Application app, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Exchange Rates"];
                if (worksheet == null)
                    return "False";

                Range range = worksheet.Range["B4", "D8"];
                string numberFormat = range.NumberFormat as string;

                if (numberFormat != "0.00")
                    return "False (hien thi so duoi dang 0.00)";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau7(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Prices"];
                Range range = worksheet.Range["A1", "A1"];
                string styleName = range.Style.Name;

                if (styleName != "Title")
                    return $"False({styleName})";
            }
            catch (Exception)
            {
                return "False (Something not finish!)";
            }

            return "True";
        }


        private static string cau8(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Products"];
                Range range = worksheet.Range["A1", "A1"];
                int horizontalAlign = (int)range.HorizontalAlignment;

                if (horizontalAlign != -4131)  // -4131 == xlLeft
                    return "False(Left)";
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
                Worksheet worksheet = d.Worksheets["Projects"];

                Range a1 = worksheet.Range["A1"];
                Range a2 = worksheet.Range["A2"];

                // Kiểm tra A1 đã merge chưa
                if (!(bool)a1.MergeCells)
                    return "False(copy định dang)";

                // Kiểm tra A1 căn giữa (1 == xlCenter)
                if ((int)a1.HorizontalAlignment != 1)
                    return "False(copy định dang)";

                // Kiểm tra A1 có màu nền là 14408667 không
                if ((int)(a1.Interior.Color) != 14408667)
                    return "False(copy định dang)";

                // Kiểm tra A2 đã merge chưa
                if (!(bool)a2.MergeCells)
                    return "False(copy định dang)";

                // Kiểm tra A2 căn giữa
                if ((int)a2.HorizontalAlignment != 1)
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
