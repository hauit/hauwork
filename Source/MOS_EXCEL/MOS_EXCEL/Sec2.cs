// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Sec2
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
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
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, d.Worksheets[(object) "Products"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite2 = Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3.Target((CallSite) Sec2.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3, worksheet.get_Range((object) "D2", (object) "D32").NumberFormat);
        if (target((CallSite) pSite2, obj) != "0.000")
          return "False(chuỗi định dạng phải la: 0.000)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau2(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site5 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site5 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site5.Target((CallSite) Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site5, d.Worksheets[(object) "Materials"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site6 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site6 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site6.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite6 = Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site6;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site7 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site7 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site7.Target((CallSite) Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site7, worksheet.get_Range((object) "A1", (object) "N1").MergeCells);
        if (target1((CallSite) pSite6, obj1) != "True")
          return "False(MergeCell)";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site8 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site8 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site8.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite8 = Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site8;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site9 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site9 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site9.Target((CallSite) Sec2.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site9, worksheet.get_Range((object) "A1", (object) "N1").HorizontalAlignment);
        if (target2((CallSite) pSite8, obj2) != "1")
          return "False(không thay đổi canh lề)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau3(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Siteb == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Siteb = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Siteb.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Siteb, d.Worksheets[(object) "Materials"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitec == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitec = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num1 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitec.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitec, worksheet.get_Range((object) "A4", (object) "A4").Width);
        if (num1 < 159.0 && num1 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sited == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sited = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num2 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sited.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sited, worksheet.get_Range((object) "B4", (object) "B4").Width);
        if (num2 < 159.0 && num2 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitee == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitee = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num3 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitee.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitee, worksheet.get_Range((object) "C4", (object) "C4").Width);
        if (num3 < 159.0 && num3 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitef == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitef = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num4 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitef.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Sitef, worksheet.get_Range((object) "D4", (object) "D4").Width);
        if (num4 < 159.0 && num4 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site10 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site10 = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num5 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site10.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site10, worksheet.get_Range((object) "E4", (object) "E4").Width);
        if (num5 < 159.0 && num5 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site11 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site11 = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num6 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site11.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site11, worksheet.get_Range((object) "F4", (object) "F4").Width);
        if (num6 < 159.0 && num6 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site12 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site12 = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num7 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site12.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site12, worksheet.get_Range((object) "G4", (object) "G4").Width);
        if (num7 < 159.0 && num7 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site13 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site13 = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num8 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site13.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site13, worksheet.get_Range((object) "H4", (object) "H4").Width);
        if (num8 < 159.0 && num8 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site14 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site14 = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num9 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site14.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site14, worksheet.get_Range((object) "I4", (object) "I4").Width);
        if (num9 < 159.0 && num9 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site15 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site15 = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num10 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site15.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site15, worksheet.get_Range((object) "J4", (object) "J4").Width);
        if (num10 < 159.0 && num10 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site16 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site16 = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num11 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site16.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site16, worksheet.get_Range((object) "K4", (object) "K4").Width);
        if (num11 < 159.0 && num11 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site17 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site17 = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num12 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site17.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site17, worksheet.get_Range((object) "L4", (object) "L4").Width);
        if (num12 < 159.0 && num12 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site18 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site18 = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num13 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site18.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site18, worksheet.get_Range((object) "M4", (object) "M4").Width);
        if (num13 < 159.0 && num13 > 161.0)
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site19 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site19 = CallSite<Func<CallSite, object, double>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (double), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        double num14 = Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site19.Target((CallSite) Sec2.\u003Ccau3\u003Eo__SiteContainera.\u003C\u003Ep__Site19, worksheet.get_Range((object) "N4", (object) "N4").Width);
        if (num14 < 159.0 && num14 > 161.0)
          return "False";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau4(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1b = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1b.Target((CallSite) Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1b, d.Worksheets[(object) "Games"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1c = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1c.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite1c = Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1c;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1d = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1d.Target((CallSite) Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1d, worksheet.get_Range((object) "A12", (object) "B18").MergeCells);
        if (target1((CallSite) pSite1c, obj1) == "True")
          return "False(chọn merge cross)";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1e = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite1e = Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1e;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1f = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1f.Target((CallSite) Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1f, worksheet.get_Range((object) "A12", (object) "B12").MergeCells);
        if (target2((CallSite) pSite1e, obj2) != "True")
          return "False(chọn merge cross)";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site20 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site20 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target3 = Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site20.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite20 = Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site20;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site21 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site21 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site21.Target((CallSite) Sec2.\u003Ccau4\u003Eo__SiteContainer1a.\u003C\u003Ep__Site21, worksheet.get_Range((object) "A18", (object) "B18").MergeCells);
        if (target3((CallSite) pSite20, obj3) != "True")
          return "False(chọn merge cross)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau5(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site23 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site23 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site23.Target((CallSite) Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site23, d.Worksheets[(object) "Materials"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site24 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site24 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site24.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite24 = Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site24;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site25 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site25 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site25.Target((CallSite) Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site25, worksheet.get_Range((object) "A2", (object) "A2").WrapText);
        if (target1((CallSite) pSite24, obj1) != "True")
          return "False(dong 2)";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site26 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site26 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site26.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite26 = Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site26;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site27 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site27 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site27.Target((CallSite) Sec2.\u003Ccau5\u003Eo__SiteContainer22.\u003C\u003Ep__Site27, worksheet.get_Range((object) "A3", (object) "A3").WrapText);
        if (target2((CallSite) pSite26, obj2) != "True")
          return "False(dong 3)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau6(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site29 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site29 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site29.Target((CallSite) Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site29, d.Worksheets[(object) "Exchange Rates"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite2a = Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site2b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site2b = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site2b.Target((CallSite) Sec2.\u003Ccau6\u003Eo__SiteContainer28.\u003C\u003Ep__Site2b, worksheet.get_Range((object) "B4", (object) "D8").NumberFormat);
        if (target((CallSite) pSite2a, obj) != "0.00")
          return "False (hien thi so duoi dang 0.00)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau7(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau7\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau7\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec2.\u003Ccau7\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d.Target((CallSite) Sec2.\u003Ccau7\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d, d.Worksheets[(object) "Prices"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau7\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau7\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2e = CallSite<Func<CallSite, object, Style>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Style), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        Style style = Sec2.\u003Ccau7\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2e.Target((CallSite) Sec2.\u003Ccau7\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2e, worksheet.get_Range((object) "A1", (object) "A1").Style);
        string name = style.Name;
        if (name != "Title")
          return "False(" + name + ")";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau8(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site30 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site30 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site30.Target((CallSite) Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site30, d.Worksheets[(object) "Products"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site31 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site31 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site31.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite31 = Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site31;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site32 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site32 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site32.Target((CallSite) Sec2.\u003Ccau8\u003Eo__SiteContainer2f.\u003C\u003Ep__Site32, worksheet.get_Range((object) "A1", (object) "A1").HorizontalAlignment);
        if (target((CallSite) pSite31, obj) != "-4131")
          return "False(Left)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau9(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site34 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site34 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site34.Target((CallSite) Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site34, d.Worksheets[(object) "Projects"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site35 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site35 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site35.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite35 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site35;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site36 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site36 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site36.Target((CallSite) Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site36, worksheet.get_Range((object) "A1", (object) "A1").MergeCells);
        if (target1((CallSite) pSite35, obj1) != "True")
          return "False(copy định dang)";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site37 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site37 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site37.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite37 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site37;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site38 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site38 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site38.Target((CallSite) Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site38, worksheet.get_Range((object) "A1", (object) "A1").HorizontalAlignment);
        if (target2((CallSite) pSite37, obj2) != "1")
          return "False(copy định dang)";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site39 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site39 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target3 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site39.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite39 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site39;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3a = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3a.Target((CallSite) Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3a, worksheet.get_Range((object) "A1", (object) "A1").Interior.Color);
        if (target3((CallSite) pSite39, obj3) != "14408667")
          return "False(copy định dang)";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3b = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target4 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3b.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite3b = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3b;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3c = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj4 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3c.Target((CallSite) Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3c, worksheet.get_Range((object) "A2", (object) "A2").MergeCells);
        if (target4((CallSite) pSite3b, obj4) != "True")
          return "False(copy định dang)";
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3d = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec2)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target5 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3d.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite3d = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3d;
        // ISSUE: reference to a compiler-generated field
        if (Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3e = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec2), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj5 = Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3e.Target((CallSite) Sec2.\u003Ccau9\u003Eo__SiteContainer33.\u003C\u003Ep__Site3e, worksheet.get_Range((object) "A2", (object) "A2").HorizontalAlignment);
        if (target5((CallSite) pSite3d, obj5) != "1")
          return "False";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }
  }
}
