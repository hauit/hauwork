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
                    return "False (Ten trang tinh)";
            }
            catch
            {
                return "False (Ten trang tinh)";
            }

            try
            {
                if (worksheet.PageSetup.RightHeader != "Confidential")
                    return "False(Confidential)";
            }
            catch
            {
                return "False (Right header)";
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
                    return "False (Materials worksheet not found)";
            }
            catch
            {
                return "False (Materials worksheet not found)";
            }

            try
            {
                Range cell = worksheet.Range["A6"];
                Hyperlinks hyperlinks = cell.Hyperlinks;

                if (hyperlinks.Count != 1)
                    return "False (Number of hyperlink)";

                Hyperlink link = hyperlinks[1];
                if (link.SubAddress != "Categories!A18")
                    return "False (" + link.SubAddress + ")";
            }
            catch
            {
                return "False (Not apply hyperlink to A6)";
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
                    return "False (Shareholders Info worksheet not found)";
            }
            catch
            {
                return "False (Shareholders Info worksheet not found)";
            }

            try
            {
                Range cell = worksheet.Range["C5"];
                Hyperlinks hyperlinks = cell.Hyperlinks;

                if (hyperlinks.Count != 1)
                    return "False (Number of hyperlink)";

                Hyperlink link = hyperlinks[1];

                if (!link.Address.Contains("tailspintoys.com/beyond.html"))
                    return "False (" + link.Address + ")";

                if (link.TextToDisplay != "More Info")
                    return "False (More Info)";
            }
            catch
            {
                return "False (Not apply hyperlink to C5)";
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
                    return "False (Summary worksheet not found)";
            }
            catch
            {
                return "False (Summary worksheet not found)";
            }

            try
            {
                Range cell = worksheet.Range["A2"];
                Hyperlinks hyperlinks = cell.Hyperlinks;

                if (hyperlinks.Count != 1)
                    return "False (chèn hyperlink tại A2)";

                Hyperlink link = hyperlinks[1];

                if (!link.Address.Contains("www.nodpublishers.com"))
                    return $"False ({link.Address})";

                if (link.ScreenTip != "Company Website")
                    return "False (Company Website)";
            }
            catch
            {
                return "False (Not apply hyperlink to A2)";
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
                    return "False (Tên trang tính)";
            }
            catch
            {
                return "False (Tên trang tính)";
            }

            try
            {
                if (worksheet.PageSetup.CenterFooter != "Page &P of &N")
                    return "False (Page 1 of ?)";
            }
            catch
            {
                return "False (Lỗi khi kiểm tra footer)";
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
