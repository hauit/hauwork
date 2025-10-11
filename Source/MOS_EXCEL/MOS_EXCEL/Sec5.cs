// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Sec5
// Assembly: MOS_EXCEL_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.CompilerServices;

namespace MOS_EXCEL_LEARN
{
  internal class Sec5
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec5.cau1(a, d);
        case 2:
          return Sec5.cau2(a, d);
        case 3:
          return Sec5.cau3(a, d);
        case 4:
          return Sec5.cau4(a, d);
        case 5:
          return Sec5.cau5(a, d);
        case 6:
          return Sec5.cau6(a, d);
        default:
          return "False";
      }
    }

        private static string cau1(Application a, Workbook d)
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


        private static string cau2(Application a, Workbook d)
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

        private static string cau3(Application a, Workbook d)
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
                if (shape.Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectBackgroundRemoval)
                    return "False";

                // Kiểm tra Pattern Fill
                if (shape.Fill.Pattern != MsoPatternType.msoPattern20Percent)
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau4(Application app, Workbook workbook)
        {
            Worksheet worksheet;

            try
            {
                worksheet = workbook.Worksheets["Summary"] as Worksheet;
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

        private static string cau5(Application a, Workbook d)
        {
            Worksheet worksheet;

            // Bước 1: Lấy worksheet tên "Summary"
            try
            {
                worksheet = d.Worksheets["Summary"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            // Bước 2: Kiểm tra số lượng shape
            try
            {
                if (worksheet.Shapes.Count != 1)
                    return "False";

                Microsoft.Office.Interop.Excel.Shape shape = worksheet.Shapes.Item(1);

                if (shape.Type != MsoShapeType.msoAutoShape)
                    return "False";

                if (shape.AutoShapeType != MsoAutoShapeType.msoShapeVerticalScroll)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau6(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Olympic Men Single Sculls"] as Worksheet;
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

                var shape = worksheet.Shapes.Item(1);
                if (shape.Type != MsoShapeType.msoAutoShape)
                    return "False";

                if (shape.AutoShapeType != MsoAutoShapeType.msoShapeVerticalScroll)
                    return "False";

                var text = shape.TextFrame2.TextRange.Text;
                var expectedText = "Top experts - from trail runners to CEOs to beloved authors - reveal the trails that fuel their dreams.";
                if (text != expectedText)
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
