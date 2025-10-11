// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Sec8
// Assembly: MOS_EXCEL_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace MOS_EXCEL_LEARN
{
  internal class Sec8
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec8.cau2(a, d);
        case 2:
          return Sec8.cau3(a, d);
        case 3:
          return Sec8.cau4(a, d);
        case 4:
          return Sec8.cau1(a, d);
        case 5:
          return Sec8.cau5(a, d);
        default:
          return "False";
      }
    }

        private static string cau1(Application a, Workbook d)
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

        private static string cau2(Application a, Workbook d)
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

        private static string cau3(Application a, Workbook d)
        {
            Worksheet worksheet;
            try
            {
                worksheet = d.Worksheets["Shareholders Info"] as Worksheet;
                if (worksheet == null)
                    return "False";
            }
            catch
            {
                return "False";
            }

            try
            {
                Range cell = worksheet.Range["C5"];
                Hyperlinks hyperlinks = cell.Hyperlinks;

                if (hyperlinks.Count != 1)
                    return "False";

                Hyperlink link = hyperlinks[1];

                if (!link.Address.Contains("tailspintoys.com/beyond.html"))
                    return "False";

                if (link.TextToDisplay != "More Info")
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

        private static string cau5(Application a, Workbook d)
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

        private static string cau6(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Materials"] as Worksheet;
                if (worksheet == null)
                    return "False (Không tìm thấy trang 'Materials')";

                worksheet.Activate();

                if (!a.ActiveWindow.FreezePanes)
                    return "False (Chưa bật Freeze Panes)";

                if (a.ActiveCell.Address != "$A$6")
                    return "False (Chưa chọn ô A6 trước khi Freeze Panes)";
            }
            catch
            {
                return "False (Lỗi trong khi kiểm tra Freeze Panes)";
            }

            return "True";
        }
    }
}
