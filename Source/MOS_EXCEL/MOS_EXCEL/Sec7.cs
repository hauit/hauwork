// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Sec7
// Assembly: MOS_EXCEL_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;

namespace MOS_EXCEL_LEARN
{
  internal class Sec7
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec7.cau16(a, d);
        case 2:
          return Sec7.cau22(a, d);
        case 3:
          return Sec7.cau15(a, d);
        case 4:
          return Sec7.cau9(a, d);
        case 5:
          return Sec7.cau14(a, d);
        case 6:
          return Sec7.cau19(a, d);
        case 7:
          return Sec7.cau20(a, d);
        case 8:
          return Sec7.cau17(a, d);
        case 9:
          return Sec7.cau21(a, d);
        case 10:
          return Sec7.cau10(a, d);
        case 11:
          return Sec7.cau13(a, d);
        case 12:
          return Sec7.cau11(a, d);
        case 13:
          return Sec7.cau18(a, d);
        case 14:
          return Sec7.cau12(a, d);
        case 15:
          return Sec7.cau15(a, d);
        case 16:
          return Sec7.cau16(a, d);
        case 17:
          return Sec7.cau17(a, d);
        case 18:
          return Sec7.cau18(a, d);
        case 19:
          return Sec7.cau19(a, d);
        case 20:
          return Sec7.cau20(a, d);
        case 21:
          return Sec7.cau21(a, d);
        case 22:
          return Sec7.cau22(a, d);
        default:
          return "False Out Range";
      }
    }

        private static string cau1(Application a, Workbook d)
        {
            try
            {
                // Tìm worksheet "Key Applications"
                Worksheet worksheet = null;
                foreach (Worksheet ws in d.Worksheets)
                {
                    if (ws.Name == "Key Applications")
                    {
                        worksheet = ws;
                        break;
                    }
                }

                if (worksheet == null)
                    return "False (Tên trang tính)";

                // Lấy ô J2
                Range cell = worksheet.Range["J2"];

                // Kiểm tra Sparkline
                SparklineGroups sparkGroups = cell.SparklineGroups;
                if (sparkGroups == null || sparkGroups.Count != 1)
                    return "False(chen Sparkline)";

                SparklineGroup sparkGroup = sparkGroups[1];

                if (sparkGroup.SourceData != "B2:G2")
                    return "False(B2:G2)";

                if (sparkGroup.Type != XlSparkType.xlSparkColumn)
                    return "False(Column)";

                return "True";
            }
            catch (Exception)
            {
                return "False (something wrong)";
            }
        }

        private static string cau2(Application a, Workbook d)
        {
            try
            {
                // Lấy worksheet "Key Applications"
                Worksheet worksheet = null;
                foreach (Worksheet ws in d.Worksheets)
                {
                    if (ws.Name == "Key Applications")
                    {
                        worksheet = ws;
                        break;
                    }
                }

                if (worksheet == null)
                    return "False (Tên Trang Tính)";

                if (worksheet.Shapes.Count != 1)
                    return "False (Number of shape)";

                // Lấy shape đầu tiên
                Shape shape = worksheet.Shapes.Item(1);

                // Kiểm tra có phải là biểu đồ không
                if (shape.Type != Microsoft.Office.Core.MsoShapeType.msoChart)
                    return "False (Not Chart)";

                Chart chart = shape.Chart;

                // Kiểm tra legend width
                if (chart.Legend == null || chart.Legend.Width < 100.0)
                    return "False (add series)";

                // Lấy các series
                SeriesCollection seriesCollection = (SeriesCollection)chart.SeriesCollection();

                if (seriesCollection.Count != 3)
                    return "False (co 3 series)";

                Series series3 = seriesCollection.Item(3);
                if (!series3.FormulaR1C1.Contains("R32C2:R32C7"))
                    return "False (add series sai)";

                return "True";
            }
            catch (Exception)
            {
                return "False (Something wrong)";
            }
        }

        private static string cau3(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = (Worksheet)d.Worksheets["Q1 Sales"];
            }
            catch
            {
                return "Fales (Ten trang tinh)";
            }

            try
            {
                Range sparkRange = worksheet.get_Range("H6", "H36");
                SparklineGroups sparkGroups = sparkRange.SparklineGroups;

                if (sparkGroups == null || sparkGroups.Count != 1)
                    return "False(chen Sparkline)";

                SparklineGroup sparkGroup = sparkGroups[1];

                if (sparkGroup.SourceData != "C6:E36")
                    return "False(C6:E6)";

                if (sparkGroup.Type != XlSparkType.xlSparkLine)
                    return "False(Column)";
            }
            catch
            {
                return "False (something wrong)";
            }

            return "True";
        }

        private static string cau4(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = (Worksheet)d.Worksheets["Q2 Sales"];
            }
            catch
            {
                return "Fales (Tên Trang Tính)";
            }

            if (worksheet.Shapes.Count != 1)
                return "False (Number of shape)";

            Chart chart;
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False (Not Chart)";
            }

            try
            {
                if (chart.Legend.Width < 90.0)
                    return "False (add series)";

                SeriesCollection seriesCollection = (SeriesCollection)chart.SeriesCollection(Type.Missing);

                if (seriesCollection.Count != 3)
                    return "False(co 3 series)";

                string formulaR1C1 = seriesCollection.Item(3).FormulaR1C1;

                if (!formulaR1C1.Contains("'Q2 Sales'!R6C5:R36C5,3)"))
                    return "False(add series sai)";
            }
            catch
            {
                return "False(Something wrong)";
            }

            return "True";
        }

        private static string cau5(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                // Kiểm tra số lượng sheet tổng và worksheet
                if (d.Sheets.Count != 5)
                    return "False(Duy chuyen chart qua trang tin moi)";
                if (d.Worksheets.Count != 4)
                    return "False(dung Move chart, khong tao trang tinh moi roi cut qua)";

                worksheet = (Worksheet)d.Worksheets["Outbound Calls"];
            }
            catch
            {
                return "Fales (Outbound Calls)";
            }

            try
            {
                if (worksheet.Shapes.Count != 1)
                    return "False (Move not copy)";

                Shape shape = worksheet.Shapes.Item(1);
                Chart chart = shape.Chart;

                if (chart.ChartType == XlChartType.xlLine)
                    return "False (Move line chart)";
            }
            catch
            {
                return "False (Not Chart)";
            }

            try
            {
                object sheet = d.Sheets["Whale Tour Sales"];
            }
            catch
            {
                return "False(Whale Tour Sales)";
            }

            return "True";
        }

        private static string cau6(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Kiểm tra tồn tại sheet "Inbound call"
            try
            {
                worksheet = (Worksheet)d.Worksheets["Inbound call"];
            }
            catch
            {
                return "Fales (ten trang tinh)";
            }

            Chart chart;

            // Bước 2: Kiểm tra shape "Chart 1" tồn tại và là biểu đồ
            try
            {
                Shape shape = worksheet.Shapes.Item("Chart 1");
                chart = shape.Chart;
            }
            catch
            {
                return "False (Not Chart)";
            }

            try
            {
                // Bước 3: Kiểm tra legend width
                if (chart.Legend.Width < 200.0)
                    return "False (Swap colunms)";

                SeriesCollection seriesCollection = (SeriesCollection)chart.SeriesCollection(Type.Missing);

                // Bước 4: Kiểm tra số lượng series
                if (seriesCollection.Count != 6)
                    return "False(co 6 series)";

                // Bước 5: Kiểm tra công thức R1C1 của series thứ 3
                string expectedFormula = "=SERIES('Inbound call'!R17C2,'Inbound call'!R2C3:R2C6,'Inbound call'!R17C3:R17C6,3)";
                if (seriesCollection.Item(3).FormulaR1C1 != expectedFormula)
                    return "False(Swap colunms)";
            }
            catch
            {
                return "False(Something wrong)";
            }

            return "True";
        }

        private static string cau7(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Kiểm tra tồn tại trang tính "Demographics"
            try
            {
                worksheet = (Worksheet)d.Worksheets["Demographics"];
            }
            catch
            {
                return "Fales (ten trang tinh)";
            }

            Chart chart;

            // Bước 2: Lấy biểu đồ từ shape
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False (insert a chart)";
            }

            // Bước 3: Kiểm tra chi tiết biểu đồ
            try
            {
                SeriesCollection seriesCollection = (SeriesCollection)chart.SeriesCollection(Type.Missing);

                if (seriesCollection.Count != 1)
                    return "False(co 1 series)";

                string expectedFormula = "=SERIES(Demographics!R3C8,Demographics!R4C7:R9C7,Demographics!R4C8:R9C8,1)";
                if (seriesCollection.Item(1).FormulaR1C1 != expectedFormula)
                    return "False(series)";

                if (chart.ChartTitle.Text != "Donations by Age Group")
                    return "False(Donations by Age Group)";

                if (chart.ChartType != XlChartType.xl3DColumnClustered)
                    return "False(chartType)";
            }
            catch
            {
                return "False(Something wrong)";
            }

            return "True";
        }

        private static string cau8(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Kiểm tra xem có tồn tại sheet "New York City" không
            try
            {
                worksheet = (Worksheet)d.Worksheets["New York City"];
            }
            catch
            {
                return "Fales (New York City worksheet not found)";
            }

            // Bước 2: Sheet phải chứa đúng 1 biểu đồ
            if (worksheet.Shapes.Count != 1)
                return "False (chèn chart)";

            Chart chart;

            // Bước 3: Lấy chart từ shape
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False (Not Chart)";
            }

            // Bước 4: Kiểm tra loại biểu đồ và series
            try
            {
                if (chart.ChartType != XlChartType.xlColumnClustered)
                    return "False (Chart phải là Clustered Column)";

                SeriesCollection seriesCollection = (SeriesCollection)chart.SeriesCollection(Type.Missing);

                if (seriesCollection.Count != 1)
                    return "False(co 1 series)";

                string expectedFormula = "=SERIES('New York City'!R4C4,'New York City'!R5C2:R21C2,'New York City'!R5C4:R21C4,1)";
                if (seriesCollection.Item(1).FormulaR1C1 != expectedFormula)
                    return "False(series)";
            }
            catch
            {
                return "False(khong xác định)";
            }

            return "True";
        }

        private static string cau9(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Kiểm tra có tồn tại sheet "London"
            try
            {
                worksheet = (Worksheet)d.Worksheets["London"];
            }
            catch
            {
                return "Fales (London worksheet not found)";
            }

            // Bước 2: Sheet phải có đúng 2 shape
            if (worksheet.Shapes.Count != 2)
                return "False (Number of shape)";

            Chart chart;

            // Bước 3: Lấy biểu đồ từ shape thứ 2
            try
            {
                Shape shape = worksheet.Shapes.Item(2);
                chart = shape.Chart;
            }
            catch
            {
                return "False (not Chart)";
            }

            // Bước 4: Kiểm tra DataTable không hiển thị LegendKey
            try
            {
                if (chart.DataTable.ShowLegendKey)
                    return "False(không show LegendKey)";
            }
            catch
            {
                return "Fasle(show dataTable)";
            }

            return "True";
        }

        private static string cau10(Application a, Workbook d)
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
                return "Fales (New Accounts worksheet not found)";
            }

            // Bước 2: Sheet phải có đúng 1 shape
            if (worksheet.Shapes.Count != 1)
                return "False (Number of shape)";

            Chart chart;

            // Bước 3: Lấy chart từ shape
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False (Not Chart)";
            }

            // Bước 4: Kiểm tra PlotBy có vẽ theo dòng không (phải là theo cột)
            try
            {
                if (chart.PlotBy == XlRowCol.xlRows)
                    return "False (biểu đồ dòng)";
            }
            catch
            {
                return "False (not add title)";
            }

            return result;
        }

        private static string cau11(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Lấy worksheet "Summary"
            try
            {
                worksheet = (Worksheet)d.Worksheets["Summary"];
            }
            catch
            {
                return "Fales (Summary worksheet not found)";
            }

            // Bước 2: Kiểm tra số shape
            if (worksheet.Shapes.Count != 1)
                return "False (Number of shape)";

            Chart chart;

            // Bước 3: Lấy chart từ shape
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False (Not Chart)";
            }

            try
            {
                // Bước 4: Kiểm tra loại biểu đồ
                if (chart.ChartType != XlChartType.xlColumnClustered)
                    return "False (not 3DPie)";  // ❗️Ghi chú: Câu thông báo sai — đây là ColumnClustered, không phải 3DPie

                // Bước 5: Lấy series và kiểm tra số lượng
                SeriesCollection seriesCollection = (SeriesCollection)chart.SeriesCollection(Type.Missing);
                if (seriesCollection.Count != 2)
                    return "False(co 2 series)";

                // Bước 6: Kiểm tra công thức của series thứ 2
                string expectedFormula = "=SERIES(Summary!R5C3,Summary!R6C1:R12C1,Summary!R6C3:R12C3,2)";
                if (seriesCollection.Item(2).FormulaR1C1 != expectedFormula)
                    return $"False ({expectedFormula})";
            }
            catch
            {
                return "False (không xát định)";
            }

            return "True";
        }

        private static string cau12(Application a, Workbook d)
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
                    return "False(chen Sparkline)";

                SparklineGroup group = sparklineGroups[1];

                // Bước 3: Kiểm tra vùng nguồn dữ liệu
                if (group.SourceData != "D5:F25")
                    return "False(D5:F25)";

                // Bước 4: Kiểm tra loại Sparkline
                if (group.Type != XlSparkType.xlSparkColumn)
                    return "False(Column)";
            }
            catch
            {
                return "False (Not Chart)";
            }

            return "True";
        }

        private static string cau13(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Kiểm tra tồn tại sheet "Graduation"
            try
            {
                worksheet = (Worksheet)d.Worksheets["Graduation"];
            }
            catch
            {
                return "Fales (Graduation worksheet not found)";
            }

            // Bước 2: Kiểm tra không còn biểu đồ trong sheet "Graduation"
            if (worksheet.Shapes.Count != 0)
                return "False (Move Chart)";

            // Bước 3: Kiểm tra tổng số worksheet là 5
            try
            {
                if (d.Worksheets.Count != 5)
                    return "False(dùng chức năng move chart)";

                // Bước 4: Kiểm tra tồn tại sheet "Graduation Chart"
                object chartSheet = d.Sheets["Graduation Chart"];
                string name = chartSheet.ToString(); // Kiểm tra không lỗi
            }
            catch
            {
                return "False (Graduation Chart)";
            }

            return "True";
        }

        private static string cau14(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Tìm sheet "Instructional Hours"
            try
            {
                worksheet = (Worksheet)d.Worksheets["Instructional Hours"];
            }
            catch
            {
                return "Fales (Instructional Hours worksheet not found)";
            }

            // Bước 2: Phải có đúng 1 shape (biểu đồ)
            if (worksheet.Shapes.Count != 1)
                return "False (Number of shape)";

            // Bước 3: Lấy biểu đồ từ shape
            Chart chart;
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False (Not Chart)";
            }

            // Bước 4: Kiểm tra tiêu đề trục tung (Value Axis)
            try
            {
                Axis yAxis = chart.Axes(XlAxisType.xlValue);
                if (yAxis.AxisTitle.Text != "Hours")
                    return "False(Hours)";
            }
            catch
            {
                return "False (not add title)";
            }

            return "True";
        }

        private static string cau15(Application a, Workbook d)
        {
            string result = "True";
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Inventory"] as Worksheet;
                if (worksheet == null)
                    return "False (Inventory)";
            }
            catch
            {
                return "False (Inventory)";
            }

            if (worksheet.Shapes.Count != 1)
                return "False (Number of shape)";

            Chart chart;
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False (Not Chart)";
            }

            try
            {
                if (chart.ChartTitle.Top != 2.0)
                    return "False()";
            }
            catch
            {
                return "False(show Title)";
            }

            try
            {
                Series series = chart.SeriesCollection(1) as Series;
                if (series == null || !series.HasDataLabels)
                    return "False(show Data)";
            }
            catch
            {
                return "False()";
            }

            return result;
        }

        private static string cau16(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Next Semester"] as Worksheet;
                if (worksheet == null)
                    return "False (Next Semester worksheet not found)";
            }
            catch
            {
                return "False (Next Semester worksheet not found)";
            }

            if (worksheet.Shapes.Count != 1)
                return "False (Number of shape)";

            Chart chart;
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False (Not Chart)";
            }

            try
            {
                // Kiểm tra loại biểu đồ
                if (chart.ChartType != XlChartType.xlColumnClustered)
                    return "False (ChartType not xlColumnClustered)";

                SeriesCollection seriesCollection = chart.SeriesCollection() as SeriesCollection;
                if (seriesCollection == null || seriesCollection.Count != 1)
                    return "False (co 2 series)";

                Series series = seriesCollection.Item(1);
                if (series.FormulaR1C1 != "=SERIES('Next Semester'!R3C5,'Next Semester'!R4C1:R21C1,'Next Semester'!R4C5:R21C5,1)")
                    return "False (series formula not matched)";
            }
            catch
            {
                return "False (không xác định)";
            }

            return "True";
        }

        private static string cau17(Application a, Workbook d)
        {
            Worksheet worksheet;

            try
            {
                worksheet = d.Worksheets["Enrollment Summary"] as Worksheet;
                if (worksheet == null)
                    return "False (Enrollment Summary worksheet not found)";
            }
            catch
            {
                return "False (Enrollment Summary worksheet not found)";
            }

            if (worksheet.Shapes.Count != 1)
                return "False (Number of shape)";

            Chart chart;
            try
            {
                Shape shape = worksheet.Shapes.Item(1);
                chart = shape.Chart;
            }
            catch
            {
                return "False (Not Chart)";
            }

            try
            {
                // Kiểm tra ChartStyle
                string style = chart.ChartStyle?.ToString();
                if (style != "268")
                    return "False (ChartStyle)";

                // Kiểm tra ChartColor
                string color = chart.ChartColor?.ToString();
                if (color != "19")
                    return "False(Color)";
            }
            catch
            {
                return "False (không xác định)";
            }

            return "True";
        }

        private static string cau18(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["New Policies"] as Worksheet;
                if (worksheet == null)
                    return "False (New Policies worksheet not found)";
            }
            catch
            {
                return "False (New Policies worksheet not found)";
            }

            try
            {
                Range sparkRange = worksheet.Range["J5", "J13"];
                var sparkGroups = sparkRange.SparklineGroups;

                if (sparkGroups.Count != 1)
                    return "False (Missing or too many Sparkline groups)";

                var sparkGroup = sparkGroups[1];

                if (sparkGroup.SourceData != "B5:G13")
                    return $"False (Wrong SourceData: {sparkGroup.SourceData})";

                if (sparkGroup.Type != XlSparkType.xlSparkColumnStacked100)
                    return $"False (Wrong Sparkline Type: {sparkGroup.Type})";
            }
            catch
            {
                return "False (Error checking Sparkline)";
            }

            return "True";
        }

        private static string cau19(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["New Policies"] as Worksheet;
                if (worksheet == null)
                    return "False (New Policies worksheet not found)";
            }
            catch
            {
                return "False (New Policies worksheet not found)";
            }

            if (worksheet.Shapes.Count != 1)
                return "False (Chart count is not 1)";

            Chart chart;
            try
            {
                chart = worksheet.Shapes.Item(1).Chart;
            }
            catch
            {
                return "False (Not a Chart)";
            }

            try
            {
                if (chart.DataTable != null && chart.DataTable.ShowLegendKey)
                    return "False (Legend key shown in DataTable)";
            }
            catch
            {
                // Có thể biểu đồ không có DataTable → kiểm tra legend thay thế
                try
                {
                    if (chart.Legend == null || chart.Legend.Position != XlLegendPosition.xlLegendPositionBottom)
                        return "False (Legend not at bottom)";
                }
                catch
                {
                    return "False (Cannot determine layout)";
                }

                return "True";
            }

            return "False (Layout not correct)";
        }

        private static string cau20(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Summary"] as Worksheet;
                if (worksheet == null)
                    return "False (Summary worksheet not found)";
            }
            catch
            {
                return "False (Summary worksheet not found)";
            }

            if (worksheet.Shapes.Count != 1)
                return "False (không thêm xóa chart)";

            try
            {
                Chart chart = worksheet.Shapes.Item(1).Chart;

                string chartColor = chart.ChartColor?.ToString();
                if (chartColor != "11")
                    return "False(sai màu)";
            }
            catch
            {
                return "False (Graduation Chart)";
            }

            return "True";
        }

        private static string cau21(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Comparison"] as Worksheet;
                if (worksheet == null)
                    return "False (Comparison worksheet not found)";
            }
            catch
            {
                return "False (Comparison worksheet not found)";
            }

            if (worksheet.Shapes.Count != 1)
                return "False (không thêm xóa chart)";

            try
            {
                Chart chart = worksheet.Shapes.Item(1).Chart;
                if (chart.PlotBy != XlRowCol.xlRows)
                    return "False(Biểu đồ Cột)";
            }
            catch
            {
                return "False (Chart)";
            }

            return "True";
        }

        private static string cau22(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Score Distribution"] as Worksheet;
                if (worksheet == null)
                    return "False (Score Distribution worksheet not found)";
            }
            catch
            {
                return "False (Score Distribution worksheet not found)";
            }

            if (worksheet.Shapes.Count != 1)
                return "False (không thêm xóa chart)";

            Chart chart;
            try
            {
                chart = worksheet.Shapes.Item(1).Chart;
            }
            catch
            {
                return "False (Chart)";
            }

            try
            {
                // Kiểm tra nếu có Legend thì sai
                if (chart.Legend != null)
                    return "False(Legend)";
            }
            catch
            {
                // Nếu có lỗi khi kiểm tra Legend, kiểm tra Series[1] có bật DataLabels không
                try
                {
                    Series series = chart.SeriesCollection(1);
                    if (!series.HasDataLabels)
                        return "False(show Data)";
                }
                catch
                {
                    return "False(show Data)";
                }
            }

            return "True";
        }
    }
}
