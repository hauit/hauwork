// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Sec3
// Assembly: MOS_EXCEL_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace MOS_EXCEL_LEARN
{
  internal class Sec3
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec3.cau17(a, d);
        case 2:
          return Sec3.cau4(a, d);
        case 3:
          return Sec3.cau26(a, d);
        case 4:
          return Sec3.cau31(a, d);
        case 5:
          return Sec3.cau32(a, d);
        case 6:
          return Sec3.cau24(a, d);
        case 7:
          return Sec3.cau7(a, d);
        case 8:
          return Sec3.cau34(a, d);
        case 9:
          return Sec3.cau35(a, d);
        case 10:
          return Sec3.cau27(a, d);
        case 11:
          return Sec3.cau18(a, d);
        case 12:
          return Sec3.cau19(a, d);
        case 13:
          return Sec3.cau30(a, d);
        case 14:
          return Sec3.cau23(a, d);
        case 15:
          return Sec3.cau11(a, d);
        case 16:
          return Sec3.cau10(a, d);
        case 17:
          return Sec3.cau3(a, d);
        case 18:
          return Sec3.cau28(a, d);
        case 19:
          return Sec3.cau20(a, d);
        case 20:
          return Sec3.cau21(a, d);
        case 21:
          return Sec3.cau29(a, d);
        case 22:
          return Sec3.cau25(a, d);
        case 23:
          return Sec3.cau22(a, d);
        case 24:
          return Sec3.cau6(a, d);
        case 25:
          return Sec3.cau5(a, d);
        case 26:
          return Sec3.cau9(a, d);
        case 27:
          return Sec3.cau27(a, d);
        case 28:
          return Sec3.cau28(a, d);
        case 29:
          return Sec3.cau29(a, d);
        case 30:
          return Sec3.cau30(a, d);
        case 31:
          return Sec3.cau31(a, d);
        case 32:
          return Sec3.cau32(a, d);
        case 33:
          return Sec3.cau33(a, d);
        case 34:
          return Sec3.cau34(a, d);
        case 35:
          return Sec3.cau35(a, d);
        default:
          return "case out of";
      }
    }

        private static string cau1(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Customers"];
                Range n2 = worksheet.Range["N2"];

                string formula = n2.Formula.ToString();
                if (formula != "=AVERAGE(Table1[CurrenAge])")
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
                Worksheet worksheet = d.Worksheets["Orders"];
                Range j2 = worksheet.Range["J2"];
                string formula = j2.Formula.ToString();

                if (formula != "=MAX(G2:G526)")
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

        private static string cau4(Application a, Workbook d)
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

        private static string cau5(Application a, Workbook d)
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

        private static string cau6(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Key Applications"];

                string formulaI2 = worksheet.Range["I2"].Formula.ToString();
                if (formulaI2 != "=IF(H2>719,\"Yes\",\"No\")")
                    return "False";

                string formulaI30 = worksheet.Range["I30"].Formula.ToString();
                if (formulaI30 != "=IF(H30>719,\"Yes\",\"No\")")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau7(Application a, Workbook d)
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

        private static string cau8(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["October"];
                string formula = worksheet.Range["G35"].Formula.ToString();

                if (formula != "=G34-F35-E35")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau9(Application a, Workbook d)
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

        private static string cau10(Application a, Workbook d)
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

        private static string cau11(Application a, Workbook d)
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

        private static string cau12(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["October"];
                string formula = worksheet.Range["F37"].Formula.ToString();

                if (!formula.Contains("=CONCATENATE(D6\", \",G6)"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau13(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Demographics"];
                string formulaK4 = ws.Range["K4"].Formula.ToString();
                string formulaK5 = ws.Range["K5"].Formula.ToString();

                if (!formulaK4.Contains("=COUNTIF(Table5[Size Group],\">=12\")"))
                    return "False";

                if (!formulaK5.Contains("=SUMIF(Table5[Size Group],\">=12\",Table5[Total Donations])"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau14(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Cars Price"];
                string formula = ws.Range["B13"].Formula.ToString();

                if (!formula.Contains("=AVERAGEIF(Table14[Discount],\">0\",Table14[Discount])"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau15(Application a, Workbook d)
        {
            try
            {
                Worksheet ws = d.Worksheets["Cars Price"];
                Range cell = ws.Range["E4"];

                string formula = cell.Formula.ToString();
                string numberFormat = cell.NumberFormat.ToString();

                if (!formula.Contains("[@Price]*[@[Discount % ]]"))
                    return "False";

                if (!numberFormat.Contains("$"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau16(Application app, Workbook workbook)
        {
            try
            {
                Worksheet sheet = workbook.Worksheets["Donor List"];

                // Kiểm tra công thức COUNTIF ở ô H9
                string formulaH9 = sheet.Range["H9"].Formula.ToString();
                if (!formulaH9.Contains("COUNTIF(Table2[Donation Level],G9)"))
                    return "False";

                // Kiểm tra công thức COUNTIF được tự động điền ở H12
                string formulaH12 = sheet.Range["H12"].Formula.ToString();
                if (!formulaH12.Contains("COUNTIF(Table2[Donation Level],G12)"))
                    return "False";

                // Kiểm tra công thức SUMIF ở ô I9
                string formulaI9 = sheet.Range["I9"].Formula.ToString();
                if (!formulaI9.Contains("=SUMIF(Table2[Donation Level],G9,Table2[Annual Donations])"))
                    return "False";

                // Kiểm tra công thức SUMIF được tự động điền ở I12
                string formulaI12 = sheet.Range["I12"].Formula.ToString();
                if (!formulaI12.Contains("=SUMIF(Table2[Donation Level],G12,Table2[Annual Donations])"))
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau17(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["London"];

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

        private static string cau18(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["New York City"];
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

        private static string cau19(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["Key Accounts"];

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

        private static string cau20(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["Contact"];

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

        private static string cau21(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["Historical Sales"];

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

        private static string cau22(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["Authors"];

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

        private static string cau23(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["Sales"];

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

        private static string cau24(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["Prices"];

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

        private static string cau25(Application app, Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets["Prices"];

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

        private static string cau26(Application a, Workbook d)
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

        private static string cau27(Application a, Workbook d)
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

        private static string cau28(Application a, Workbook d)
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

        private static string cau29(Application a, Workbook d)
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

        private static string cau30(Application a, Workbook d)
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

        private static string cau31(Application a, Workbook d)
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

        private static string cau32(Application a, Workbook d)
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

        private static string cau33(Application a, Workbook d)
        {
            try
            {
                Worksheet worksheet = d.Worksheets["Summary"];

                string formulaB15 = worksheet.Range["B15"].Formula?.ToString() ?? "";
                if (!formulaB15.Contains("=MAX(F4:F11)"))
                    return "False";
            }
            catch
            {
                return "False";
            }

            return "True";
        }

        private static string cau34(Application a, Workbook d)
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

        private static string cau35(Application a, Workbook d)
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
    }
}
