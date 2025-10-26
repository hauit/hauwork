// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Sec9
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
  internal class Sec9
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec9.cau17(a, d);
        case 2:
          return Sec9.cau8(a, d);
        case 3:
          return Sec9.cau16(a, d);
        case 4:
          return Sec9.cau1(a, d);
        case 5:
          return Sec9.cau2(a, d);
        case 6:
          return Sec9.cau7(a, d);
        case 7:
          return Sec9.cau13(a, d);
        case 8:
          return Sec9.cau14(a, d);
        case 9:
          return Sec9.cau3(a, d);
        case 10:
          return Sec9.cau10New(a, d);
        case 11:
          return Sec9.cau11(a, d);
        case 12:
          return Sec9.cau12(a, d);
        case 13:
          return Sec9.cau13(a, d);
        case 14:
          return Sec9.cau14(a, d);
        case 15:
          return Sec9.cau15(a, d);
        case 16:
          return Sec9.cau16(a, d);
        case 17:
          return Sec9.cau17(a, d);
        case 18:
          return Sec9.cau18(a, d);
        case 19:
          return Sec9.cau19(a, d);
        case 20:
          return Sec9.cau20(a, d);
        default:
          return "False";
      }
    }

        private static string cau1(Application a, Workbook d)
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

        private static string cau2(Application a, Workbook d)
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

        private static string cau3(Application a, Workbook d)
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

        private static string cau4(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Living Online"];
                string textInE32 = ws.Range["E32"].Text;

                if (textInE32 == "=AVERAGE(E2:E30)")
                    return "False";
            }
            catch (Exception ex)
            {
                return ex.Message.Contains("Living Online")
                    ? "False"
                    : "False";
            }

            return "True";
        }

        private static string cau5(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Q2 Sales"];
                string textInF6 = ws.Range["F6"].Text;

                if (textInF6 == "=AVERAGE(Table2[@[April]:[June]])")
                    return "False";
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Q2 Sales"))
                    return "False";
                return "False";
            }

            return "True";
        }

        private static string cau6(Application a, Workbook d)
        {
            try
            {
                var properties = d.BuiltinDocumentProperties;
                var titleProp = properties.GetType().InvokeMember("Item",
                                  BindingFlags.GetProperty, null, properties, new object[] { "Title" });
                string titleValue = titleProp.GetType().InvokeMember("Value",
                                  BindingFlags.GetProperty, null, titleProp, new object[0]).ToString();

                if (titleValue != "2015")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau7(Application a, Workbook d)
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

        private static string cau8(Application a, Workbook d)
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

        private static string cau9(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Q2 Sales"];
            }
            catch
            {
                return "False";
            }

            try
            {
                string cellText = worksheet.get_Range("F6", "F6").Text.ToString();
                if (cellText != "=AVERAGE(Table2[@[April]:[June]])")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau10(Application excelApp, Workbook workbook)
        {
            try
            {
                object properties = workbook.BuiltinDocumentProperties;
                object subjectProp = properties.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty,
                    null,
                    properties,
                    new object[] { "Subject" });

                object subjectValue = subjectProp.GetType().InvokeMember(
                    "Value",
                    BindingFlags.GetProperty,
                    null,
                    subjectProp,
                    new object[0]);

                if (!string.IsNullOrEmpty(subjectValue?.ToString()))
                    return "False"; // Subject có giá trị
            }
            catch
            {
                return "False";
            }

            return "True"; // Subject không có giá trị
        }

        private static string cau10New(Application a, Workbook d)
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
                return "False " + ex.ToString();
            }

        }

        private static string cau11(Application app, Workbook workbook)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets["Expenses"] as Worksheet;
                if (worksheet == null)
                    return "False";

                string printArea = worksheet.PageSetup.PrintArea;
                if (printArea != "$B$5:$D$52")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau12(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Scholarships"] as Worksheet;
                if (worksheet == null)
                    return "False";

                if (worksheet.PageSetup.PrintTitleColumns != "$A:$A")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau13(Application a, Workbook d)
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

        private static string cau14(Application a, Workbook d)
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

        private static string cau15(Application a, Workbook d)
        {
            try
            {
                // Sheet 1: "Games"
                Worksheet worksheet1 = d.Worksheets["Games"] as Worksheet;
                if (worksheet1 == null)
                    return "False";

                if (worksheet1.ListObjects["Table2"].AlternativeText != "data")
                    return "False";
                if (worksheet1.ListObjects["Table3"].AlternativeText != "data")
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                // Sheet 2: "Shareholders Info"
                Worksheet worksheet2 = d.Worksheets["Shareholders Info"] as Worksheet;
                if (worksheet2 == null)
                    return "False";

                if (worksheet2.ListObjects["Table1"].AlternativeText != "data")
                    return "False";

                if (worksheet2.Shapes.Item("Chart 1").Title != "data")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau16(Application a, Workbook d)
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

        private static string cau17(Application a, Workbook d)
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

        private static string cau18(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                // Lấy worksheet tên "Q2 Sales"
                worksheet = (Worksheet)d.Worksheets["Q2 Sales"];
            }
            catch
            {
                return "False";
            }

            try
            {
                // Lấy nội dung hiển thị của ô F6 (Text), không phải giá trị tính toán
                string cellText = worksheet.Range["F6"].Text.ToString();

                // Nếu không đúng công thức hiển thị, tức là chưa bật chế độ "Show formulas"
                if (cellText != "=AVERAGE(Table2[@[April]:[June]])")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau19(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                // Lấy worksheet "Q2 Sales"
                worksheet = (Worksheet)d.Worksheets["Q2 Sales"];
            }
            catch
            {
                return "False";
            }

            try
            {
                // Lấy nội dung text hiển thị trong ô F6
                string cellText = worksheet.Range["F6"].Text.ToString();

                // So sánh với công thức mong muốn (ở chế độ hiển thị công thức)
                if (cellText != "=AVERAGE(Table2[@[April]:[June]])")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau20(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = (Worksheet)d.Worksheets["Q2 Sales"];
            }
            catch
            {
                return "False";
            }

            try
            {
                // Lấy nội dung hiển thị của ô F6
                string cellText = worksheet.Range["F6"].Text.ToString();

                // Nếu vẫn đang hiển thị công thức, thì sai
                if (cellText == "=AVERAGE(Table2[@[April]:[June]])")
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }
    }
}
