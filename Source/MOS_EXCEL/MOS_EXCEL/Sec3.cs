// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Sec3
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
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
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, d.Worksheets[(object) "Customers"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite2 = Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3.Target((CallSite) Sec3.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3, worksheet.get_Range((object) "N2", (object) "N2").Formula);
        if (target((CallSite) pSite2, obj) != "=AVERAGE(Table1[CurrenAge])")
          return "False(=AVERAGE(Table1[CurrenAge]))";
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
        if (Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site5 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site5 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site5.Target((CallSite) Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site5, d.Worksheets[(object) "Orders"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site6 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site6 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site6.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite6 = Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site6;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site7 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site7 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site7.Target((CallSite) Sec3.\u003Ccau2\u003Eo__SiteContainer4.\u003C\u003Ep__Site7, worksheet.get_Range((object) "J2", (object) "J2").Formula);
        if (target((CallSite) pSite6, obj) != "=MAX(G2:G526)")
          return "False(==MAX(G2:G526)))";
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
        if (Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Site9 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Site9 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Site9.Target((CallSite) Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Site9, d.Worksheets[(object) "roster"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSitea = Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb.Target((CallSite) Sec3.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb, worksheet.get_Range((object) "C8", (object) "C8").Formula);
        if (target((CallSite) pSitea, obj) != "=PROPER(A8)")
          return "False(=PROPER(A8))";
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
        if (Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sited == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sited = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sited.Target((CallSite) Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sited, d.Worksheets[(object) "Non_Fiction"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSitee = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef.Target((CallSite) Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef, worksheet.get_Range((object) "H5", (object) "H5").Formula);
        if (target1((CallSite) pSitee, obj1) != "=F5-G5")
          return "False(=F5-G5)";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site10 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site10 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site10.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite10 = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site10;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site11 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site11 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site11.Target((CallSite) Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site11, worksheet.get_Range((object) "H35", (object) "H35").Formula);
        if (target2((CallSite) pSite10, obj2) != "=F35-G35")
          return "False(=F35-G35)";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site12 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site12 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target3 = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site12.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite12 = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site12;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site13 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site13 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site13.Target((CallSite) Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site13, worksheet.get_Range((object) "H5", (object) "H5").NumberFormat);
        if (target3((CallSite) pSite12, obj3) != "General")
          return "False(General)";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site14 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site14 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target4 = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site14.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite14 = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site14;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site15 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site15 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj4 = Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site15.Target((CallSite) Sec3.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site15, worksheet.get_Range((object) "H6", (object) "H6").Interior.Color);
        if (target4((CallSite) pSite14, obj4) != "16777215")
          return "False(without format)";
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
        if (Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site17 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site17 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site17.Target((CallSite) Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site17, d.Worksheets[(object) "Non_Fiction"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site18 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site18 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site18.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite18 = Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site18;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site19 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site19 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site19.Target((CallSite) Sec3.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site19, worksheet.get_Range((object) "F37", (object) "F37").Formula);
        string str = target((CallSite) pSite18, obj);
        if (!str.Contains("=AVERAGEIF(D5:D35,\"Lucerne Publishing\",F5:F35)") && !str.Contains("=AVERAGEIF($D$5:$D$35,\"Lucerne Publishing\",$F$5:$F$35)"))
          return "False(=AVERAGEIF(D5:D35,\"Lucerne Publishing\",F5:F35))";
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
        if (Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1b = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1b.Target((CallSite) Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1b, d.Worksheets[(object) "Key Applications"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1c = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1c.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite1c = Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1c;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1d = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1d.Target((CallSite) Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1d, worksheet.get_Range((object) "I2", (object) "I2").Formula);
        if (target1((CallSite) pSite1c, obj1) != "=IF(H2>719,\"Yes\",\"No\")")
          return "False(=IF(H2>719,\"Yes\",\"No\"))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1e = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite1e = Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1e;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1f = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1f.Target((CallSite) Sec3.\u003Ccau6\u003Eo__SiteContainer1a.\u003C\u003Ep__Site1f, worksheet.get_Range((object) "I30", (object) "I30").Formula);
        if (target2((CallSite) pSite1e, obj2) != "=IF(H30>719,\"Yes\",\"No\")")
          return "False(Auto Fill)";
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
        if (Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site21 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site21 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site21.Target((CallSite) Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site21, d.Worksheets[(object) "Summary"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site22 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site22 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site22.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite22 = Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site22;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site23 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site23 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site23.Target((CallSite) Sec3.\u003Ccau7\u003Eo__SiteContainer20.\u003C\u003Ep__Site23, worksheet.get_Range((object) "B15", (object) "B15").Formula);
        if (target((CallSite) pSite22, obj) != "=MAX(F4:F11)")
          return "False(=MAX(F4:F11))";
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
        if (Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site25 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site25 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site25.Target((CallSite) Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site25, d.Worksheets[(object) "October"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site26 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site26 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site26.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite26 = Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site26;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site27 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site27 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site27.Target((CallSite) Sec3.\u003Ccau8\u003Eo__SiteContainer24.\u003C\u003Ep__Site27, worksheet.get_Range((object) "G35", (object) "G35").Formula);
        if (target((CallSite) pSite26, obj) != "=G34-F35-E35")
          return "False(Auto Fill)";
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
        if (Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site29 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site29 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site29.Target((CallSite) Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site29, d.Worksheets[(object) "October"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite2a = Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site2a;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site2b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site2b = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site2b.Target((CallSite) Sec3.\u003Ccau9\u003Eo__SiteContainer28.\u003C\u003Ep__Site2b, worksheet.get_Range((object) "E37", (object) "E37").Formula);
        string str = target((CallSite) pSite2a, obj);
        if (!str.Contains("=AVERAGEIF(E11:E35,\">300\",E11:E35)") && !str.Contains("=AVERAGEIF($E$11:$E$35,\">300\",$E$11:$E$35)"))
          return "False(=AVERAGEIF(E11:E35,\">300\",E11:E35))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau10(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d.Target((CallSite) Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d, d.Worksheets[(object) "roster"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2e = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite2e = Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2e;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2f = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2f.Target((CallSite) Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2f, worksheet.get_Range((object) "B9", (object) "B9").Formula);
        if (!target1((CallSite) pSite2e, obj1).Contains("=LOWER(D9)"))
          return "False(=LOWER(D9))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site30 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site30 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site30.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite30 = Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site30;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site31 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site31 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site31.Target((CallSite) Sec3.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site31, worksheet.get_Range((object) "B66", (object) "B66").Formula);
        if (!target2((CallSite) pSite30, obj2).Contains("=LOWER(D66)"))
          return "False(Auto Fill)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau11(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site33 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site33 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site33.Target((CallSite) Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site33, d.Worksheets[(object) "roster"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site34 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site34 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site34.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite34 = Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site34;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site35 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site35 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site35.Target((CallSite) Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site35, worksheet.get_Range((object) "C9", (object) "C9").Formula);
        if (!target1((CallSite) pSite34, obj1).Contains("=UPPER(A9)"))
          return "False(=UPPER(A9))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site36 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site36 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site36.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite36 = Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site36;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site37 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site37 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site37.Target((CallSite) Sec3.\u003Ccau11\u003Eo__SiteContainer32.\u003C\u003Ep__Site37, worksheet.get_Range((object) "C66", (object) "C66").Formula);
        if (!target2((CallSite) pSite36, obj2).Contains("=UPPER(A66)"))
          return "False(Auto Fill)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau12(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site39 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site39 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site39.Target((CallSite) Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site39, d.Worksheets[(object) "October"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site3a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site3a = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site3a.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite3a = Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site3a;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site3b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site3b = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site3b.Target((CallSite) Sec3.\u003Ccau12\u003Eo__SiteContainer38.\u003C\u003Ep__Site3b, worksheet.get_Range((object) "F37", (object) "F37").Formula);
        if (!target((CallSite) pSite3a, obj).Contains("=CONCATENATE(D6,\", \",G6)"))
          return "False(=CONCATENATE(D6,\", \",G6))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau13(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3d = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3d.Target((CallSite) Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3d, d.Worksheets[(object) "Demographics"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3e = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite3e = Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3e;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3f = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3f.Target((CallSite) Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site3f, worksheet.get_Range((object) "K4", (object) "K4").Formula);
        if (!target1((CallSite) pSite3e, obj1).Contains("=COUNTIF(Table5[Size Group],\">=12\")"))
          return "False(=COUNTIF(Table5[Size Group],\">=12\"))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site40 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site40 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site40.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite40 = Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site40;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site41 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site41 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site41.Target((CallSite) Sec3.\u003Ccau13\u003Eo__SiteContainer3c.\u003C\u003Ep__Site41, worksheet.get_Range((object) "K5", (object) "K5").Formula);
        if (!target2((CallSite) pSite40, obj2).Contains("=SUMIF(Table5[Size Group],\">=12\",Table5[Total Donations])"))
          return "False(=SUMIF(Table5[Size Group],\">=12\",Table5[Total Donations]))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau14(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site43 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site43 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site43.Target((CallSite) Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site43, d.Worksheets[(object) "Cars Price"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site44 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site44 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site44.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite44 = Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site44;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site45 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site45 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site45.Target((CallSite) Sec3.\u003Ccau14\u003Eo__SiteContainer42.\u003C\u003Ep__Site45, worksheet.get_Range((object) "B13", (object) "B13").Formula);
        if (!target((CallSite) pSite44, obj).Contains("=AVERAGEIF(Table14[Discount],\">0\""))
          return "False(=AVERAGEIF(Table14[Discount],\">0\",Table14[Discount]))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau15(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site47 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site47 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site47.Target((CallSite) Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site47, d.Worksheets[(object) "Cars Price"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site48 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site48 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site48.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite48 = Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site48;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site49 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site49 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site49.Target((CallSite) Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site49, worksheet.get_Range((object) "E4", (object) "E4").Formula);
        if (!target1((CallSite) pSite48, obj1).Contains("[@Price]*[@[Discount % ]]"))
          return "False([@Price]*[@[Discount % ]])";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site4a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site4a = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site4a.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite4a = Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site4a;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site4b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site4b = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site4b.Target((CallSite) Sec3.\u003Ccau15\u003Eo__SiteContainer46.\u003C\u003Ep__Site4b, worksheet.get_Range((object) "E4", (object) "E4").NumberFormat);
        if (!target2((CallSite) pSite4a, obj2).Contains("$"))
          return "False(apply Accounting format)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau16(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4d = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet1 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4d.Target((CallSite) Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4d, d.Worksheets[(object) "Donor List"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4e = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite4e = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4e;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4f = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4f.Target((CallSite) Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4f, worksheet1.get_Range((object) "H9", (object) "H9").Formula);
        if (!target1((CallSite) pSite4e, obj1).Contains("COUNTIF(Table2[Donation Level],G9)"))
          return "False(COUNTIF(Table2[Donation Level],G9))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site50 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site50 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site50.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite50 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site50;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site51 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site51 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site51.Target((CallSite) Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site51, worksheet1.get_Range((object) "H12", (object) "H12").Formula);
        if (!target2((CallSite) pSite50, obj2).Contains("COUNTIF(Table2[Donation Level],G12)"))
          return "False(Auto Fill)";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site52 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site52 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet2 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site52.Target((CallSite) Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site52, d.Worksheets[(object) "Donor List"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site53 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site53 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target3 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site53.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite53 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site53;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site54 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site54 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site54.Target((CallSite) Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site54, worksheet2.get_Range((object) "I9", (object) "I9").Formula);
        if (!target3((CallSite) pSite53, obj3).Contains("=SUMIF(Table2[Donation Level],G9,Table2[Annual Donations])"))
          return "False(=SUMIF(Table2[Donation Level],G9,Table2[Annual Donations]))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site55 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site55 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target4 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site55.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite55 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site55;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site56 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site56 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj4 = Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site56.Target((CallSite) Sec3.\u003Ccau16\u003Eo__SiteContainer4c.\u003C\u003Ep__Site56, worksheet2.get_Range((object) "I12", (object) "I12").Formula);
        if (!target4((CallSite) pSite55, obj4).Contains("=SUMIF(Table2[Donation Level],G12,Table2[Annual Donations])"))
          return "False(Auto Fill)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau17(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site58 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site58 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site58.Target((CallSite) Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site58, d.Worksheets[(object) "London"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site59 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site59 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site59.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite59 = Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site59;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5a = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5a.Target((CallSite) Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5a, worksheet.get_Range((object) "E21", (object) "E21").Formula);
        if (target1((CallSite) pSite59, obj1) != "=[@[Air Miles]]*0.08")
          return "False(Auto Fill)";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5b = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5b.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite5b = Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5b;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5c = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5c.Target((CallSite) Sec3.\u003Ccau17\u003Eo__SiteContainer57.\u003C\u003Ep__Site5c, worksheet.get_Range((object) "E21", (object) "E21").NumberFormat);
        if (target2((CallSite) pSite5b, obj2) != "General")
          return "False(không lấy định dạng)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau18(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site5e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site5e = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site5e.Target((CallSite) Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site5e, d.Worksheets[(object) "New York City"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site5f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site5f = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site5f.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite5f = Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site5f;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site60 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site60 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site60.Target((CallSite) Sec3.\u003Ccau18\u003Eo__SiteContainer5d.\u003C\u003Ep__Site60, worksheet.get_Range((object) "D23", (object) "D23").Formula);
        if (!target((CallSite) pSite5f, obj).Contains("=MAX(Table1[Air Miles]"))
          return "False(=MAX(Table1[Air Miles])";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau19(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site62 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site62 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site62.Target((CallSite) Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site62, d.Worksheets[(object) "Key Accounts"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site63 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site63 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site63.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite63 = Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site63;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site64 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site64 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site64.Target((CallSite) Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site64, worksheet.get_Range((object) "C4", (object) "C4").Formula);
        if (!target1((CallSite) pSite63, obj1).Contains("=AVERAGE(Table1[@[January]:[April]])"))
          return "False(=AVERAGE(Table1[@[January]:[April]]))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site65 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site65 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site65.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite65 = Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site65;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site66 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site66 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site66.Target((CallSite) Sec3.\u003Ccau19\u003Eo__SiteContainer61.\u003C\u003Ep__Site66, worksheet.get_Range((object) "C12", (object) "C12").Formula);
        if (!target2((CallSite) pSite65, obj2).Contains("=AVERAGE(Table1[@[January]:[April]])"))
          return "False(=AVERAGE(Table1[@[January]:[April]]))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau20(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site68 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site68 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site68.Target((CallSite) Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site68, d.Worksheets[(object) "Contact"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site69 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site69 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site69.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite69 = Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site69;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6a = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6a.Target((CallSite) Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6a, worksheet.get_Range((object) "C5", (object) "C5").Formula);
        if (!target1((CallSite) pSite69, obj1).Contains("=CONCATENATE([@[First Name]],\"@woodgrovebank.com\")"))
          return "False(=CONCATENATE([@[First Name]],\"@woodgrovebank.com\"))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6b = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6b.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite6b = Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6b;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6c = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6c.Target((CallSite) Sec3.\u003Ccau20\u003Eo__SiteContainer67.\u003C\u003Ep__Site6c, worksheet.get_Range((object) "C19", (object) "C19").Formula);
        if (!target2((CallSite) pSite6b, obj2).Contains("=CONCATENATE([@[First Name]],\"@woodgrovebank.com\")"))
          return "False(=CONCATENATE([@[First Name]],\"@woodgrovebank.com\"))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau21(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau21\u003Eo__SiteContainer6d.\u003C\u003Ep__Site6e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau21\u003Eo__SiteContainer6d.\u003C\u003Ep__Site6e = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau21\u003Eo__SiteContainer6d.\u003C\u003Ep__Site6e.Target((CallSite) Sec3.\u003Ccau21\u003Eo__SiteContainer6d.\u003C\u003Ep__Site6e, d.Worksheets[(object) "Historical Sales"]);
        if (!worksheet.Application.ActiveWindow.DisplayFormulas)
          return "False(bật show Formulas)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau22(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site70 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site70 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site70.Target((CallSite) Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site70, d.Worksheets[(object) "Authors"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site71 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site71 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site71.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite71 = Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site71;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site72 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site72 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site72.Target((CallSite) Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site72, worksheet.get_Range((object) "D2", (object) "D2").Formula);
        if (!target1((CallSite) pSite71, obj1).Contains("=IF([@[Books Sold]]>10000,500,100)"))
          return "False(=IF([@[Books Sold]]>10000,500,100))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site73 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site73 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site73.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite73 = Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site73;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site74 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site74 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site74.Target((CallSite) Sec3.\u003Ccau22\u003Eo__SiteContainer6f.\u003C\u003Ep__Site74, worksheet.get_Range((object) "D37", (object) "D37").Formula);
        if (!target2((CallSite) pSite73, obj2).Contains("=IF([@[Books Sold]]>10000,500,100)"))
          return "False(=IF([@[Books Sold]]>10000,500,100))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau23(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site76 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site76 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site76.Target((CallSite) Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site76, d.Worksheets[(object) "Sales"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site77 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site77 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site77.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite77 = Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site77;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site78 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site78 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site78.Target((CallSite) Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site78, worksheet.get_Range((object) "E2", (object) "E2").Formula);
        if (!target1((CallSite) pSite77, obj1).Contains("=UPPER(LEFT([@City],3))"))
          return "False(=UPPER(LEFT([@City],3))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site79 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site79 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site79.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite79 = Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site79;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site7a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site7a = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site7a.Target((CallSite) Sec3.\u003Ccau23\u003Eo__SiteContainer75.\u003C\u003Ep__Site7a, worksheet.get_Range((object) "E20", (object) "E20").Formula);
        if (!target2((CallSite) pSite79, obj2).Contains("=UPPER(LEFT([@City],3))"))
          return "False(=UPPER(LEFT([@City],3))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau24(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7c = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7c.Target((CallSite) Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7c, d.Worksheets[(object) "Prices"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7d = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7d.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite7d = Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7d;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7e = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7e.Target((CallSite) Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7e, worksheet.get_Range((object) "J5", (object) "J5").Formula);
        if (!target1((CallSite) pSite7d, obj1).Contains("=[@[Unit Price]]*$L$2"))
          return "False(=[@[Unit Price]]*$L$2)";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7f = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7f.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite7f = Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site7f;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site80 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site80 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site80.Target((CallSite) Sec3.\u003Ccau24\u003Eo__SiteContainer7b.\u003C\u003Ep__Site80, worksheet.get_Range((object) "J25", (object) "J25").Formula);
        if (!target2((CallSite) pSite7f, obj2).Contains("=[@[Unit Price]]*$L$2"))
          return "False(=[@[Unit Price]]*$L$2)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau25(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site82 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site82 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site82.Target((CallSite) Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site82, d.Worksheets[(object) "Prices"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site83 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site83 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site83.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite83 = Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site83;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site84 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site84 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site84.Target((CallSite) Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site84, worksheet.get_Range((object) "G5", (object) "G5").Formula);
        if (!target1((CallSite) pSite83, obj1).Contains("=IF([@[Inventory Level]]<15%,\"Low\",\"\")"))
          return "False(=IF([@[Inventory Level]]<15%,\"Low\",\"\"))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site85 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site85 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site85.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite85 = Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site85;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site86 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site86 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site86.Target((CallSite) Sec3.\u003Ccau25\u003Eo__SiteContainer81.\u003C\u003Ep__Site86, worksheet.get_Range((object) "G25", (object) "G25").Formula);
        if (!target2((CallSite) pSite85, obj2).Contains("=IF([@[Inventory Level]]<15%,\"Low\",\"\")"))
          return "False(=IF([@[Inventory Level]]<15%,\"Low\",\"\"))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau26(Application a, Workbook d)
    {
      try
      {
        if (d.Names.Count != 1)
          return "False(Tạo 1 range name)";
        // ISSUE: reference to a compiler-generated method
        if (d.Names.Item((object) 1, (object) Missing.Value, (object) Missing.Value).Name != "Enrollment")
          return "False(Enrollment)";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau26\u003Eo__SiteContainer87.\u003C\u003Ep__Site88 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau26\u003Eo__SiteContainer87.\u003C\u003Ep__Site88 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau26\u003Eo__SiteContainer87.\u003C\u003Ep__Site88.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite88 = Sec3.\u003Ccau26\u003Eo__SiteContainer87.\u003C\u003Ep__Site88;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau26\u003Eo__SiteContainer87.\u003C\u003Ep__Site89 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau26\u003Eo__SiteContainer87.\u003C\u003Ep__Site89 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau26\u003Eo__SiteContainer87.\u003C\u003Ep__Site89.Target((CallSite) Sec3.\u003Ccau26\u003Eo__SiteContainer87.\u003C\u003Ep__Site89, d.Names.Item((object) 1, (object) Missing.Value, (object) Missing.Value).RefersToLocal);
        string str = target((CallSite) pSite88, obj);
        if (str != "='Enrollment Summary'!$A$3:$B$7")
          return "False(" + str + ")";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau27(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8b = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8b.Target((CallSite) Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8b, d.Worksheets[(object) "New Policies"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8c = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8c.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite8c = Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8c;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8d = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8d.Target((CallSite) Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8d, worksheet.get_Range((object) "I5", (object) "I5").Formula);
        if (!target1((CallSite) pSite8c, obj1).Contains("=COUNTBLANK(Table1[@[January]:[June]])"))
          return "False(=COUNTBLANK(Table1[@[January]:[June]]))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8e = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite8e = Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8e;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8f = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8f.Target((CallSite) Sec3.\u003Ccau27\u003Eo__SiteContainer8a.\u003C\u003Ep__Site8f, worksheet.get_Range((object) "I13", (object) "I13").Formula);
        if (!target2((CallSite) pSite8e, obj2).Contains("=COUNTBLANK(Table1[@[January]:[June]])"))
          return "False(=COUNTBLANK(Table1[@[January]:[June]]))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau28(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site91 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site91 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site91.Target((CallSite) Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site91, d.Worksheets[(object) "Contact"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site92 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site92 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site92.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite92 = Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site92;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site93 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site93 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site93.Target((CallSite) Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site93, worksheet.get_Range((object) "C5", (object) "C5").Formula);
        if (!target1((CallSite) pSite92, obj1).Contains("=CONCATENATE([@[First Name]],\"@humongousinsurance.com\")"))
          return "False(=CONCATENATE([@[First Name]],\"@humongousinsurance.com\"))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site94 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site94 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site94.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite94 = Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site94;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site95 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site95 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site95.Target((CallSite) Sec3.\u003Ccau28\u003Eo__SiteContainer90.\u003C\u003Ep__Site95, worksheet.get_Range((object) "C13", (object) "C13").Formula);
        if (!target2((CallSite) pSite94, obj2).Contains("=CONCATENATE([@[First Name]],\"@humongousinsurance.com\")"))
          return "False(=CONCATENATE([@[First Name]],\"@humongousinsurance.com\"))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau29(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site97 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site97 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site97.Target((CallSite) Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site97, d.Worksheets[(object) "February"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site98 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site98 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site98.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite98 = Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site98;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site99 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site99 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site99.Target((CallSite) Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site99, worksheet.get_Range((object) "G5", (object) "G5").Formula);
        if (!target1((CallSite) pSite98, obj1).Contains("=IF([@[Years as Member]]>3,\"Yes\",\"No\")"))
          return "False(=IF([@[Years as Member]]>3,\"Yes\",\"No\"))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site9a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site9a = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site9a.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite9a = Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site9a;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site9b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site9b = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site9b.Target((CallSite) Sec3.\u003Ccau29\u003Eo__SiteContainer96.\u003C\u003Ep__Site9b, worksheet.get_Range((object) "G18", (object) "G18").Formula);
        if (!target2((CallSite) pSite9a, obj2).Contains("=IF([@[Years as Member]]>3,\"Yes\",\"No\")"))
          return "False(=IF([@[Years as Member]]>3,\"Yes\",\"No\"))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau30(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9d = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9d.Target((CallSite) Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9d, d.Worksheets[(object) "February"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9e = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite9e = Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9e;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9f = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9f.Target((CallSite) Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Site9f, worksheet.get_Range((object) "F5", (object) "F5").Formula);
        if (!target1((CallSite) pSite9e, obj1).Contains("=LEFT([@[Policy Number ]],2)"))
          return "False(=LEFT([@[Policy Number ]],2))";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Sitea0 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Sitea0 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Sitea0.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSitea0 = Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Sitea0;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Sitea1 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Sitea1 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Sitea1.Target((CallSite) Sec3.\u003Ccau30\u003Eo__SiteContainer9c.\u003C\u003Ep__Sitea1, worksheet.get_Range((object) "F18", (object) "F18").Formula);
        if (!target2((CallSite) pSitea0, obj2).Contains("=LEFT([@[Policy Number ]],2)"))
          return "False(=LEFT([@[Policy Number ]],2))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau31(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea3 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea3.Target((CallSite) Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea3, d.Worksheets[(object) "Products"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea4 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea4 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea4.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSitea4 = Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea4;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea5 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea5 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea5.Target((CallSite) Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea5, worksheet.get_Range((object) "G3", (object) "G3").Formula);
        if (!target1((CallSite) pSitea4, obj1).Contains("=[@[Current Value]]*Increase"))
          return "False(=[@[Current Value]]*Increase)";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea6 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea6 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea6.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSitea6 = Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea6;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea7 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea7 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea7.Target((CallSite) Sec3.\u003Ccau31\u003Eo__SiteContainera2.\u003C\u003Ep__Sitea7, worksheet.get_Range((object) "G54", (object) "G54").Formula);
        if (!target2((CallSite) pSitea6, obj2).Contains("=[@[Current Value]]*Increase"))
          return "False(=[@[Current Value]]*Increase)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau32(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Sitea9 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Sitea9 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Sitea9.Target((CallSite) Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Sitea9, d.Worksheets[(object) "Projections"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteaa == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteaa = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteaa.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSiteaa = Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteaa;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteab == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteab = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteab.Target((CallSite) Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteab, worksheet.get_Range((object) "C4", (object) "C4").Formula);
        if (!target1((CallSite) pSiteaa, obj1).Contains("=[@[Quarter 1]]*Q2_Increase"))
          return "False(=[@[Quarter 1]]*Q2_Increase)";
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteac == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteac = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteac.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSiteac = Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Siteac;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Sitead == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Sitead = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Sitead.Target((CallSite) Sec3.\u003Ccau32\u003Eo__SiteContainera8.\u003C\u003Ep__Sitead, worksheet.get_Range((object) "C11", (object) "C11").Formula);
        if (!target2((CallSite) pSiteac, obj2).Contains("=[@[Quarter 1]]*Q2_Increase"))
          return "False(=[@[Quarter 1]]*Q2_Increase)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau33(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteaf == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteaf = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteaf.Target((CallSite) Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteaf, d.Worksheets[(object) "Summary"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteb0 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteb0 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteb0.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSiteb0 = Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteb0;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteb1 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteb1 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteb1.Target((CallSite) Sec3.\u003Ccau33\u003Eo__SiteContainerae.\u003C\u003Ep__Siteb1, worksheet.get_Range((object) "B15", (object) "B15").Formula);
        if (!target((CallSite) pSiteb0, obj).Contains("=MAX(F4:F11)"))
          return "False(=MAX(F4:F11))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau34(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb3 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb3.Target((CallSite) Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb3, d.Worksheets[(object) "Grade Criteria"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb4 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb4 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb4.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSiteb4 = Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb4;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb5 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb5 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb5.Target((CallSite) Sec3.\u003Ccau34\u003Eo__SiteContainerb2.\u003C\u003Ep__Siteb5, worksheet.get_Range((object) "B28", (object) "B28").Formula);
        string str = target((CallSite) pSiteb4, obj);
        if (!str.Contains("=SUM(Total1,Total2,Total3)") && !str.Contains("=Total1+Total2+Total3"))
          return "False(=SUM(Total1,Total2,Total3))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau35(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb7 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb7 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb7.Target((CallSite) Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb7, d.Worksheets[(object) "Exams"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb8 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb8 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec3)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb8.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSiteb8 = Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb8;
        // ISSUE: reference to a compiler-generated field
        if (Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb9 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb9 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec3), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj = Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb9.Target((CallSite) Sec3.\u003Ccau35\u003Eo__SiteContainerb6.\u003C\u003Ep__Siteb9, worksheet.get_Range((object) "E35", (object) "E35").Formula);
        if (!target((CallSite) pSiteb8, obj).Contains("=COUNTBLANK(Table3[Exam 3])"))
          return "False(=COUNTBLANK(Table3[Exam 3]))";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }
  }
}
