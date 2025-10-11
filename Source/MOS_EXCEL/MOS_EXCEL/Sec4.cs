// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Sec4
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
    internal class Sec4
    {
        public static string CheckCau(int causo, Application a, Workbook d)
        {
            switch (causo)
            {
                case 1:
                    return Sec4.cau1(a, d);
                case 2:
                    return Sec4.cau2(a, d);
                case 3:
                    return Sec4.cau3(a, d);
                case 4:
                    return Sec4.cau4(a, d);
                default:
                    return "False";
            }
        }

        private static string cau1(Application a, Workbook d)
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

        private static string cau2(Application a, Workbook d)
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

        private static string cau3(Application a, Workbook d)
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

        private static string cau4(Application a, Workbook d)
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
    }
}
