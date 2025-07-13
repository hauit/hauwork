// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Sec6
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
{
  internal class Sec6
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec6.cau12(a, d);
        case 2:
          return Sec6.cau16(a, d);
        case 3:
          return Sec6.cau9(a, d);
        case 4:
          return Sec6.cau13(a, d);
        case 5:
          return Sec6.cau19(a, d);
        case 6:
          return Sec6.cau14(a, d);
        case 7:
          return Sec6.cau17(a, d);
        case 8:
          return Sec6.cau8(a, d);
        case 9:
          return Sec6.cau11(a, d);
        case 10:
          return Sec6.cau15(a, d);
        case 11:
          return Sec6.cau18(a, d);
        case 12:
          return Sec6.cau10(a, d);
        case 13:
          return Sec6.cau13(a, d);
        case 14:
          return Sec6.cau14(a, d);
        case 15:
          return Sec6.cau15(a, d);
        case 16:
          return Sec6.cau16(a, d);
        case 17:
          return Sec6.cau17(a, d);
        case 18:
          return Sec6.cau18(a, d);
        case 19:
          return Sec6.cau19(a, d);
        default:
          return "False";
      }
    }

    private static string cau1(Application a, Workbook d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) Sec6.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, d.Worksheets[(object) "Customers"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject;
      try
      {
        listObject = worksheet.ListObjects[(object) 1];
      }
      catch (Exception ex)
      {
        return "False (Table)";
      }
      if (!listObject.ShowTableStyleRowStripes)
        return "False (banded rows)";
      try
      {
      }
      catch (Exception ex)
      {
        return "False (Wrong range)";
      }
      return str;
    }

    private static string cau2(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3.Target((CallSite) Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site3, d.Worksheets[(object) "Customers"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site4 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site4 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target1 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site4.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite4 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site4;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site5 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site5 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site5.Target((CallSite) Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site5, worksheet.get_Range((object) "I499", (object) "I499").Text);
        if (target1((CallSite) pSite4, obj1) != "Canada")
          return "False(sort sai Level 1)";
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site6 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site6 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target2 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site6.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite6 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site6;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site7 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site7 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj2 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site7.Target((CallSite) Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site7, worksheet.get_Range((object) "I500", (object) "I500").Text);
        if (target2((CallSite) pSite6, obj2) != "Canada")
          return "False(sort sai Level 1)";
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site8 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site8 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target3 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site8.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite8 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site8;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site9 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site9 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site9.Target((CallSite) Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site9, worksheet.get_Range((object) "I501", (object) "I501").Text);
        if (target3((CallSite) pSite8, obj3) != "Canada")
          return "False(sort sai Level 1)";
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitea == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitea = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target4 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitea.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSitea = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitea;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Siteb == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Siteb = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj4 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Siteb.Target((CallSite) Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Siteb, worksheet.get_Range((object) "G499", (object) "G499").Text);
        if (target4((CallSite) pSitea, obj4) != "QC")
          return "False(sort sai Level 2)";
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitec == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitec = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target5 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitec.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSitec = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitec;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sited == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sited = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj5 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sited.Target((CallSite) Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sited, worksheet.get_Range((object) "G2", (object) "G2").Text);
        if (target5((CallSite) pSitec, obj5) != "AK")
          return "False(sort sai Level 2)";
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitee == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitee = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target6 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitee.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSitee = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitee;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitef == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitef = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj6 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitef.Target((CallSite) Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Sitef, worksheet.get_Range((object) "h2", (object) "h2").Text);
        if (target6((CallSite) pSitee, obj6) != "16202")
          return "False(sort sai Level 3)";
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site10 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site10 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target7 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site10.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite10 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site10;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site11 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site11 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj7 = Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site11.Target((CallSite) Sec6.\u003Ccau2\u003Eo__SiteContainer2.\u003C\u003Ep__Site11, worksheet.get_Range((object) "h4", (object) "h4").Text);
        if (target7((CallSite) pSite10, obj7) != "64577")
          return "False(sort sai Level 3)";
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      return "True";
    }

    private static string cau3(Application a, Workbook d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau3\u003Eo__SiteContainer12.\u003C\u003Ep__Site13 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau3\u003Eo__SiteContainer12.\u003C\u003Ep__Site13 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau3\u003Eo__SiteContainer12.\u003C\u003Ep__Site13.Target((CallSite) Sec6.\u003Ccau3\u003Eo__SiteContainer12.\u003C\u003Ep__Site13, d.Worksheets[(object) "Inbound call"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject;
      try
      {
        listObject = worksheet.ListObjects[(object) 1];
      }
      catch (Exception ex)
      {
        return "False (Table)";
      }
      try
      {
        if (!listObject.ShowTotals)
          return "False (Tatals)";
      }
      catch (Exception ex)
      {
        return "False (Table)";
      }
      return str;
    }

    private static string cau4(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau4\u003Eo__SiteContainer14.\u003C\u003Ep__Site15 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau4\u003Eo__SiteContainer14.\u003C\u003Ep__Site15 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau4\u003Eo__SiteContainer14.\u003C\u003Ep__Site15.Target((CallSite) Sec6.\u003Ccau4\u003Eo__SiteContainer14.\u003C\u003Ep__Site15, d.Worksheets[(object) "Donor List"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.ListObjects.Count.ToString() != "0")
          return "False(conver table to range)";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
      return "True";
    }

    private static string cau5(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site17 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site17 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site17.Target((CallSite) Sec6.\u003Ccau5\u003Eo__SiteContainer16.\u003C\u003Ep__Site17, d.Worksheets[(object) "Donor List"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.ListObjects.Count.ToString() != "2")
          return "False(chuyen tu day o sang table)";
        // ISSUE: variable of a compiler-generated type
        ListObject listObject = worksheet.ListObjects[(object) 2];
        if (listObject.ListRows.Count != 4)
          return "False(check vao table has headerRow khi insert table)";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
      return "True";
    }

    private static string cau6(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site19 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site19 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site19.Target((CallSite) Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site19, d.Worksheets[(object) "Demographics"]);
      }
      catch (Exception ex)
      {
        return "Fales (Ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1a = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1a.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite1a = Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1a;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1b = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1b.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite1b = Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1b;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1c = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1c.Target((CallSite) Sec6.\u003Ccau6\u003Eo__SiteContainer18.\u003C\u003Ep__Site1c, worksheet.get_Range((object) "C5", (object) "C5").get_Value((object) Missing.Value));
        object obj2 = target2((CallSite) pSite1b, obj1, "300");
        if (target1((CallSite) pSite1a, obj2))
          return "False(C5)";
      }
      catch (Exception ex)
      {
        return "False (something wrong)";
      }
      return "True";
    }

    private static string cau7(Application a, Workbook d)
    {
      string str = "True";
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau7\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau7\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1e = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau7\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1e.Target((CallSite) Sec6.\u003Ccau7\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1e, d.Worksheets[(object) "Donor List"]);
      }
      catch (Exception ex)
      {
        return "Fales (Ten trang tinh)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject;
      try
      {
        listObject = worksheet.ListObjects[(object) "Table2"];
      }
      catch (Exception ex)
      {
        return "False (Name of table)";
      }
      try
      {
        if (listObject.AlternativeText != "Donor")
          return "False (Donor chuot phai chon table->alt text)";
      }
      catch (Exception ex)
      {
        return "False (add alt text)";
      }
      return str;
    }

    private static string cau8(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site20 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site20 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site20.Target((CallSite) Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site20, d.Worksheets[(object) "New York City"]);
      }
      catch (Exception ex)
      {
        return "Fales (Ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site21 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site21 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site21.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite21 = Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site21;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site22 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site22 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site22.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite22 = Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site22;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site23 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site23 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site23.Target((CallSite) Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site23, worksheet.get_Range((object) "A7", (object) "A7").Text);
        object obj2 = target2((CallSite) pSite22, obj1, "China");
        if (target1((CallSite) pSite21, obj2))
          return "False(sai ở sort cấp 1)";
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site24 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site24 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site24.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite24 = Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site24;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site25 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site25 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site25.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite25 = Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site25;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site26 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site26 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site26.Target((CallSite) Sec6.\u003Ccau8\u003Eo__SiteContainer1f.\u003C\u003Ep__Site26, worksheet.get_Range((object) "B7", (object) "B7").Text);
        object obj4 = target4((CallSite) pSite25, obj3, "Beijing");
        if (target3((CallSite) pSite24, obj4))
          return "False(sai ở sort cấp 2)";
      }
      catch (Exception ex)
      {
        return "False (something wrong)";
      }
      return "True";
    }

    private static string cau9(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site28 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site28 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site28.Target((CallSite) Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site28, d.Worksheets[(object) "New Accounts"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tính New Accounts)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        ListObject listObject = worksheet.get_Range((object) "A3", (object) "A3").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site29 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site29 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target1 = Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site29.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite29 = Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site29;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site2a == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site2a = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target2 = Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site2a.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite2a = Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site2a;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site2b == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site2b = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj1 = Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site2b.Target((CallSite) Sec6.\u003Ccau9\u003Eo__SiteContainer27.\u003C\u003Ep__Site2b, worksheet.get_Range((object) "A6", (object) "A6").Text);
      object obj2 = target2((CallSite) pSite2a, obj1, "Fabrikam, Inc.");
      return target1((CallSite) pSite29, obj2) ? "False" : "True";
    }

    private static string cau10(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d.Target((CallSite) Sec6.\u003Ccau10\u003Eo__SiteContainer2c.\u003C\u003Ep__Site2d, d.Worksheets[(object) "Classes"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tính Classes)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject;
      try
      {
        // ISSUE: reference to a compiler-generated method
        listObject = worksheet.get_Range((object) "A4", (object) "F25").ListObject;
      }
      catch (Exception ex)
      {
        return "True";
      }
      return listObject == null ? "True" : "False";
    }

    private static string cau11(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site2f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site2f = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site2f.Target((CallSite) Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site2f, d.Worksheets[(object) "Orders"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tính Orders)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject;
      try
      {
        // ISSUE: reference to a compiler-generated method
        listObject = worksheet.get_Range((object) "A1", (object) "A1").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      try
      {
        // ISSUE: variable of a compiler-generated type
        Microsoft.Office.Interop.Excel.Filter filter = listObject.AutoFilter.Filters[1];
        if (!filter.On)
          return "False(filter tren cột 1)";
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site30 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site30 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site30.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite30 = Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site30;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site31 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site31 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site31.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite31 = Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site31;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site32 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site32 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        object obj1 = Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site32.Target((CallSite) Sec6.\u003Ccau11\u003Eo__SiteContainer2e.\u003C\u003Ep__Site32, filter.Criteria1);
        object obj2 = target2((CallSite) pSite31, obj1, "=Alpine Ski House");
        if (target1((CallSite) pSite30, obj2))
          return "False(filter cột 1 chọn Alpine Ski House)";
      }
      catch (Exception ex)
      {
        return "False()";
      }
      return "True";
    }

    private static string cau12(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau12\u003Eo__SiteContainer33.\u003C\u003Ep__Site34 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau12\u003Eo__SiteContainer33.\u003C\u003Ep__Site34 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau12\u003Eo__SiteContainer33.\u003C\u003Ep__Site34.Target((CallSite) Sec6.\u003Ccau12\u003Eo__SiteContainer33.\u003C\u003Ep__Site34, d.Worksheets[(object) "Revenue"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tính New Accounts)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        ListObject listObject = worksheet.get_Range((object) "A3", (object) "A3").ListObject;
      }
      catch (Exception ex)
      {
        return "False (A3)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject1;
      try
      {
        // ISSUE: reference to a compiler-generated method
        listObject1 = worksheet.get_Range((object) "B7", (object) "B7").ListObject;
      }
      catch (Exception ex)
      {
        return "False (B7)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (listObject1.Range.get_Address((object) Missing.Value, (object) Missing.Value, External: (object) Missing.Value, RelativeTo: (object) Missing.Value) != "$A$3:$B$7")
        {
          // ISSUE: reference to a compiler-generated method
          return "False(" + listObject1.Range.get_Address((object) Missing.Value, (object) Missing.Value, External: (object) Missing.Value, RelativeTo: (object) Missing.Value) + ")";
        }
      }
      catch (Exception ex)
      {
        return "False(chưa chuyển sang table)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau12\u003Eo__SiteContainer33.\u003C\u003Ep__Site35 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau12\u003Eo__SiteContainer33.\u003C\u003Ep__Site35 = CallSite<Func<CallSite, object, TableStyle>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (TableStyle), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        TableStyle tableStyle = Sec6.\u003Ccau12\u003Eo__SiteContainer33.\u003C\u003Ep__Site35.Target((CallSite) Sec6.\u003Ccau12\u003Eo__SiteContainer33.\u003C\u003Ep__Site35, listObject1.TableStyle);
        if (tableStyle.Name != "TableStyleLight14")
          return "False(sai kiểu)";
      }
      catch (Exception ex)
      {
        return "False(không xác định lien quan đến Kiểu)";
      }
      return "True";
    }

    private static string cau13(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site37 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site37 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site37.Target((CallSite) Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site37, d.Worksheets[(object) "Last Semester"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tính Last Semester)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        // ISSUE: variable of a compiler-generated type
        ListObject listObject = worksheet.get_Range((object) "B6", (object) "B6").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site38 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site38 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target1 = Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site38.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite38 = Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site38;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site39 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site39 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target2 = Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site39.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite39 = Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site39;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site3a == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site3a = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj1 = Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site3a.Target((CallSite) Sec6.\u003Ccau13\u003Eo__SiteContainer36.\u003C\u003Ep__Site3a, worksheet.get_Range((object) "B6", (object) "B6").Text);
      object obj2 = target2((CallSite) pSite39, obj1, "Health & Beauty");
      return target1((CallSite) pSite38, obj2) ? "False" : "True";
    }

    private static string cau14(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3c = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3c.Target((CallSite) Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3c, d.Worksheets[(object) "New Policies"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tính New Policies)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject;
      try
      {
        // ISSUE: reference to a compiler-generated method
        listObject = worksheet.get_Range((object) "A4", (object) "A4").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      if (!listObject.ShowTotals)
        return "False(show dồng tổng)";
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3d == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3d = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target1 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3d.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite3d = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3d;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3e == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3e = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target2 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3e.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite3e = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3e;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3f == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3f = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj1 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3f.Target((CallSite) Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site3f, worksheet.get_Range((object) "b14", (object) "b14").Formula);
      object obj2 = target2((CallSite) pSite3e, obj1, "=SUBTOTAL(109,[January])");
      if (target1((CallSite) pSite3d, obj2))
        return "False(B14)";
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site40 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site40 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target3 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site40.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite40 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site40;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site41 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site41 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target4 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site41.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite41 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site41;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site42 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site42 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj3 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site42.Target((CallSite) Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site42, worksheet.get_Range((object) "H14", (object) "H14").Formula);
      object obj4 = target4((CallSite) pSite41, obj3, "=SUBTOTAL(109,[Total])");
      if (target3((CallSite) pSite40, obj4))
        return "False(H14)";
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site43 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site43 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target5 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site43.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite43 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site43;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site44 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site44 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target6 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site44.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite44 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site44;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site45 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site45 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj5 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site45.Target((CallSite) Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site45, worksheet.get_Range((object) "I14", (object) "I14").Formula);
      object obj6 = target6((CallSite) pSite44, obj5, "");
      if (target5((CallSite) pSite43, obj6))
        return "False(I14)";
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site46 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site46 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target7 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site46.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSite46 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site46;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site47 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site47 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target8 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site47.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite47 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site47;
      // ISSUE: reference to a compiler-generated field
      if (Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site48 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site48 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated method
      object obj7 = Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site48.Target((CallSite) Sec6.\u003Ccau14\u003Eo__SiteContainer3b.\u003C\u003Ep__Site48, worksheet.get_Range((object) "J14", (object) "J14").Formula);
      object obj8 = target8((CallSite) pSite47, obj7, "");
      return target7((CallSite) pSite46, obj8) ? "False(J14)" : "True";
    }

    private static string cau15(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4a = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4a.Target((CallSite) Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4a, d.Worksheets[(object) "March"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tính March)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject;
      try
      {
        // ISSUE: reference to a compiler-generated method
        listObject = worksheet.get_Range((object) "A4", (object) "A4").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4b = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4b.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite4b = Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4b;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4c = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4c.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite4c = Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4c;
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4d = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec6), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        object obj1 = Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4d.Target((CallSite) Sec6.\u003Ccau15\u003Eo__SiteContainer49.\u003C\u003Ep__Site4d, listObject.AutoFilter.Filters[6].Criteria1);
        object obj2 = target2((CallSite) pSite4c, obj1, "=MP");
        if (target1((CallSite) pSite4b, obj2))
          return "False";
      }
      catch (Exception ex)
      {
        return "False(không xac đinh)";
      }
      return "True";
    }

    private static string cau16(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau16\u003Eo__SiteContainer4e.\u003C\u003Ep__Site4f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau16\u003Eo__SiteContainer4e.\u003C\u003Ep__Site4f = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau16\u003Eo__SiteContainer4e.\u003C\u003Ep__Site4f.Target((CallSite) Sec6.\u003Ccau16\u003Eo__SiteContainer4e.\u003C\u003Ep__Site4f, d.Worksheets[(object) "Products"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tính Products)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject;
      try
      {
        // ISSUE: reference to a compiler-generated method
        listObject = worksheet.get_Range((object) "A4", (object) "A4").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau16\u003Eo__SiteContainer4e.\u003C\u003Ep__Site50 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau16\u003Eo__SiteContainer4e.\u003C\u003Ep__Site50 = CallSite<Func<CallSite, object, TableStyle>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (TableStyle), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        TableStyle tableStyle = Sec6.\u003Ccau16\u003Eo__SiteContainer4e.\u003C\u003Ep__Site50.Target((CallSite) Sec6.\u003Ccau16\u003Eo__SiteContainer4e.\u003C\u003Ep__Site50, listObject.TableStyle);
        if (tableStyle.Name != "TableStyleMedium1")
          return "False(sai style)";
      }
      catch (Exception ex)
      {
        return "False(style không xac đinh)";
      }
      return "True";
    }

    private static string cau17(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau17\u003Eo__SiteContainer51.\u003C\u003Ep__Site52 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau17\u003Eo__SiteContainer51.\u003C\u003Ep__Site52 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau17\u003Eo__SiteContainer51.\u003C\u003Ep__Site52.Target((CallSite) Sec6.\u003Ccau17\u003Eo__SiteContainer51.\u003C\u003Ep__Site52, d.Worksheets[(object) "Region 1"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tính Region 1)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject;
      try
      {
        // ISSUE: reference to a compiler-generated method
        listObject = worksheet.get_Range((object) "A3", (object) "A3").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated method
        if (listObject.Sort.SortFields[(object) 1].Key.get_Address((object) Missing.Value, (object) Missing.Value, External: (object) Missing.Value, RelativeTo: (object) Missing.Value) != "$A$4:$A$11")
          return "False(Product)";
        if (listObject.Sort.SortFields[(object) 1].Order != XlSortOrder.xlAscending)
          return "False(Product->A to Z)";
        // ISSUE: reference to a compiler-generated method
        if (listObject.Sort.SortFields[(object) 2].Key.get_Address((object) Missing.Value, (object) Missing.Value, External: (object) Missing.Value, RelativeTo: (object) Missing.Value) != "$F$4:$F$11")
          return "False(Total Sales)";
        if (listObject.Sort.SortFields[(object) 2].Order != XlSortOrder.xlDescending)
          return "False(Total Sales->lớn đến nhỏ)";
      }
      catch (Exception ex)
      {
        return "False(chưa sort đủ 2 trường cùng lúc)";
      }
      return "True";
    }

    private static string cau19(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau19\u003Eo__SiteContainer53.\u003C\u003Ep__Site54 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau19\u003Eo__SiteContainer53.\u003C\u003Ep__Site54 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau19\u003Eo__SiteContainer53.\u003C\u003Ep__Site54.Target((CallSite) Sec6.\u003Ccau19\u003Eo__SiteContainer53.\u003C\u003Ep__Site54, d.Worksheets[(object) "Tasks"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tính Tasks)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject;
      try
      {
        // ISSUE: reference to a compiler-generated method
        listObject = worksheet.get_Range((object) "A3", (object) "A3").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      try
      {
        if (!listObject.ShowTableStyleRowStripes)
          return "False(banded rows)";
      }
      catch (Exception ex)
      {
        return "False(tên table)";
      }
      return "True";
    }

    private static string cau18(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec6.\u003Ccau18\u003Eo__SiteContainer55.\u003C\u003Ep__Site56 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec6.\u003Ccau18\u003Eo__SiteContainer55.\u003C\u003Ep__Site56 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec6)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec6.\u003Ccau18\u003Eo__SiteContainer55.\u003C\u003Ep__Site56.Target((CallSite) Sec6.\u003Ccau18\u003Eo__SiteContainer55.\u003C\u003Ep__Site56, d.Worksheets[(object) "Tasks"]);
      }
      catch (Exception ex)
      {
        return "Fales (trang tính Tasks)";
      }
      // ISSUE: variable of a compiler-generated type
      ListObject listObject;
      try
      {
        // ISSUE: reference to a compiler-generated method
        listObject = worksheet.get_Range((object) "A3", (object) "A3").ListObject;
      }
      catch (Exception ex)
      {
        return "False (Table was modify)";
      }
      try
      {
        if (listObject.Name != "Tasks")
          return "False(table name)";
      }
      catch (Exception ex)
      {
        return "False(tên table)";
      }
      return "True";
    }
  }
}
