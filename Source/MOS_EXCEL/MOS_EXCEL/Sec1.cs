// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Sec1
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
  internal class Sec1
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec1.cau3(a, d);
        case 2:
          return Sec1.cau5(a, d);
        case 3:
          return Sec1.cau6(a, d);
        case 4:
          return Sec1.cau11(a, d);
        case 5:
          return Sec1.cau5_5(a, d);
        case 6:
          return Sec1.cau1(a, d);
        case 7:
          return Sec1.cau10(a, d);
        case 8:
          return Sec1.cau2(a, d);
        case 9:
          return Sec1.cau4(a, d);
        case 10:
          return Sec1.cau8(a, d);
        case 11:
          return Sec1.cau9(a, d);
        case 12:
          return Sec1.goto001(a, d);
        case 13:
          return Sec1.goto002(a, d);
        default:
          return "False";
      }
    }

    private static string cau1(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, d.Worksheets[(object) "Customers by Order"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite2 = Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite3 = Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4.Target((CallSite) Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4, worksheet.get_Range((object) "C502", (object) "C502").Text);
        object obj2 = target2((CallSite) pSite3, obj1, "");
        if (target1((CallSite) pSite2, obj2))
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite5 = Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site6 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site6 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site6.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite6 = Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site6;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site7 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site7 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site7.Target((CallSite) Sec1.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site7, worksheet.get_Range((object) "B502", (object) "B502").Text);
        object obj4 = target4((CallSite) pSite6, obj3, "");
        if (target3((CallSite) pSite5, obj4))
          return "False";
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
        if (Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Site9 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Site9 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Site9.Target((CallSite) Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Site9, d.Worksheets[(object) "Social Media Ads"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSitea = Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSiteb = Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Sitec == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Sitec = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Sitec.Target((CallSite) Sec1.\u003Ccau2\u003Eo__SiteContainer8.\u003C\u003Ep__Sitec, worksheet.get_Range((object) "A17", (object) "A17").Text);
        object obj2 = target2((CallSite) pSiteb, obj1, "RANCH");
        if (target1((CallSite) pSitea, obj2))
          return "False (comma)";
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
        if (Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Sitee == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Sitee = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Sitee.Target((CallSite) Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Sitee, d.Worksheets[(object) "Top Donors"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Sitef == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Sitef = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Sitef.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSitef = Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Sitef;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site10 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site10 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site10.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite10 = Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site10;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site11 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site11 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site11.Target((CallSite) Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site11, worksheet.get_Range((object) "A2", (object) "A2").Text);
        object obj2 = target2((CallSite) pSite10, obj1, "Mark Bebbington");
        if (target1((CallSite) pSitef, obj2))
          return "False (A2)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site12 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site12 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site12.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite12 = Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site12;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site13 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site13 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site13.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite13 = Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site13;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site14 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site14 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site14.Target((CallSite) Sec1.\u003Ccau3\u003Eo__SiteContainerd.\u003C\u003Ep__Site14, worksheet.get_Range((object) "C6", (object) "C6").Text);
        object obj4 = target4((CallSite) pSite13, obj3, "Platinum");
        if (target3((CallSite) pSite12, obj4))
          return "False (C6)";
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
        if (Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site16 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site16 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site16.Target((CallSite) Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site16, d.Worksheets[(object) "Top 25"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site17 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site17 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site17.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite17 = Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site17;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site18 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site18 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site18.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite18 = Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site18;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site19 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site19 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site19.Target((CallSite) Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site19, worksheet.get_Range((object) "A2", (object) "A2").Text);
        object obj2 = target2((CallSite) pSite18, obj1, "Rank");
        if (target1((CallSite) pSite17, obj2))
          return "False (A2)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1a = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1a.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite1a = Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1a;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1b = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1b.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite1b = Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1b;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1c = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1c.Target((CallSite) Sec1.\u003Ccau4\u003Eo__SiteContainer15.\u003C\u003Ep__Site1c, worksheet.get_Range((object) "G27", (object) "G27").Text);
        object obj4 = target4((CallSite) pSite1b, obj3, "London");
        if (target3((CallSite) pSite1a, obj4))
          return "False (G27)";
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
        if (Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1e = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1e.Target((CallSite) Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1e, d.Worksheets[(object) "Top Donors"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1f = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1f.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite1f = Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site1f;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site20 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site20 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site20.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite20 = Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site20;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site21 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site21 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site21.Target((CallSite) Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site21, worksheet.get_Range((object) "A2", (object) "A2").Text);
        object obj2 = target2((CallSite) pSite20, obj1, "Daniel P. Taylor");
        if (target1((CallSite) pSite1f, obj2))
          return "False (A2)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site22 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site22 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site22.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite22 = Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site22;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site23 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site23 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site23.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite23 = Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site23;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site24 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site24 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site24.Target((CallSite) Sec1.\u003Ccau5\u003Eo__SiteContainer1d.\u003C\u003Ep__Site24, worksheet.get_Range((object) "C2", (object) "C2").Text);
        object obj4 = target4((CallSite) pSite23, obj3, "Charles Fitzgerald");
        if (target3((CallSite) pSite22, obj4))
          return "False (C2)";
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
        if (Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site26 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site26 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site26.Target((CallSite) Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site26, d.Worksheets[(object) "Donor Contact Info"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site27 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site27 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site27.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite27 = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site27;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site28 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site28 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site28.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite28 = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site28;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site29 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site29 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site29.Target((CallSite) Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site29, worksheet.get_Range((object) "F2", (object) "F2").Text);
        object obj2 = target2((CallSite) pSite28, obj1, "64043");
        if (target1((CallSite) pSite27, obj2))
          return "False(thay thế tất cả)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2a = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2a.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite2a = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2a;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2b = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2b.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite2b = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2b;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2c = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2c.Target((CallSite) Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2c, worksheet.get_Range((object) "F3", (object) "F3").Text);
        object obj4 = target4((CallSite) pSite2b, obj3, "64043");
        if (target3((CallSite) pSite2a, obj4))
          return "False(thay thế tất cả)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2d = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target5 = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2d.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite2d = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2d;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2e = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target6 = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite2e = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2e;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2f = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj5 = Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2f.Target((CallSite) Sec1.\u003Ccau6\u003Eo__SiteContainer25.\u003C\u003Ep__Site2f, worksheet.get_Range((object) "H46", (object) "H46").Text);
        object obj6 = target6((CallSite) pSite2e, obj5, "314-555-64044");
        if (target5((CallSite) pSite2d, obj6))
          return "False(Chỉ thay thế trên cột F)";
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
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site31 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site31 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site31.Target((CallSite) Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site31, d.Worksheets[(object) "Donor List"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site32 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site32 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site32.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite32 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site32;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site33 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site33 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site33.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite33 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site33;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site34 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site34 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site34.Target((CallSite) Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site34, worksheet.get_Range((object) "C16", (object) "C16").Text);
        object obj2 = target2((CallSite) pSite33, obj1, "Bronze Total");
        if (target1((CallSite) pSite32, obj2))
          return "False(chọn group trên Annual Donations Hàm Sum)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site35 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site35 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site35.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite35 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site35;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site36 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site36 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site36.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite36 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site36;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site37 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site37 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site37.Target((CallSite) Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site37, worksheet.get_Range((object) "C27", (object) "C27").Text);
        object obj4 = target4((CallSite) pSite36, obj3, "Silver Total");
        if (target3((CallSite) pSite35, obj4))
          return "False(chọn group trên Annual Donations Hàm Sum)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site38 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site38 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target5 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site38.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite38 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site38;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site39 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site39 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target6 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site39.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite39 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site39;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3a = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj5 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3a.Target((CallSite) Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3a, worksheet.get_Range((object) "C67", (object) "C67").Text);
        object obj6 = target6((CallSite) pSite39, obj5, "Platinum Total");
        if (target5((CallSite) pSite38, obj6))
          return "False(chọn group trên Annual Donations Hàm Sum)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3b = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target7 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3b.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite3b = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3b;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3c = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target8 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3c.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite3c = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3c;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3d = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj7 = Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3d.Target((CallSite) Sec1.\u003Ccau7\u003Eo__SiteContainer30.\u003C\u003Ep__Site3d, worksheet.get_Range((object) "B67", (object) "B67").Text);
        object obj8 = target8((CallSite) pSite3c, obj7, "$38,021.00 ");
        if (target7((CallSite) pSite3b, obj8))
          return "False(Sum trên Donation Level)";
        if (worksheet.HPageBreaks.Count != 4)
          return "False(check vào Page break between group)";
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
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site3f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site3f = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site3f.Target((CallSite) Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site3f, d.Worksheets[(object) "Donor List"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site40 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site40 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site40.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite40 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site40;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site41 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site41 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site41.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite41 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site41;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site42 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site42 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site42.Target((CallSite) Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site42, worksheet.get_Range((object) "C16", (object) "C16").Text);
        object obj2 = target2((CallSite) pSite41, obj1, "Bronze Count");
        if (target1((CallSite) pSite40, obj2))
          return "False(chọn group trên Annual Donations hàm count)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site43 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site43 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site43.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite43 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site43;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site44 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site44 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site44.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite44 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site44;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site45 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site45 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site45.Target((CallSite) Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site45, worksheet.get_Range((object) "C27", (object) "C27").Text);
        object obj4 = target4((CallSite) pSite44, obj3, "Silver Count");
        if (target3((CallSite) pSite43, obj4))
          return "False(chọn group trên Annual Donations hàm count)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site46 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site46 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target5 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site46.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite46 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site46;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site47 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site47 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target6 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site47.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite47 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site47;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site48 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site48 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj5 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site48.Target((CallSite) Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site48, worksheet.get_Range((object) "C67", (object) "C67").Text);
        object obj6 = target6((CallSite) pSite47, obj5, "Platinum Count");
        if (target5((CallSite) pSite46, obj6))
          return "False(chọn group trên Annual Donations hàm count)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site49 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site49 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target7 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site49.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite49 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site49;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site4a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site4a = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target8 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site4a.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite4a = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site4a;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site4b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site4b = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj7 = Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site4b.Target((CallSite) Sec1.\u003Ccau8\u003Eo__SiteContainer3e.\u003C\u003Ep__Site4b, worksheet.get_Range((object) "A67", (object) "A67").Text);
        object obj8 = target8((CallSite) pSite4a, obj7, "28");
        if (target7((CallSite) pSite49, obj8))
          return "False(Count trên Name)";
        if (worksheet.HPageBreaks.Count != 4)
          return "False(check vào Page break between group)";
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
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4d = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4d.Target((CallSite) Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4d, d.Worksheets[(object) "Exchange Rates"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4e = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite4e = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4e;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4f = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4f.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite4f = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site4f;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site50 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site50 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site50.Target((CallSite) Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site50, worksheet.get_Range((object) "A7", (object) "A7").Text);
        object obj2 = target2((CallSite) pSite4f, obj1, "");
        if (target1((CallSite) pSite4e, obj2))
          return "False(A7)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site51 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site51 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site51.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite51 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site51;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site52 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site52 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site52.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite52 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site52;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site53 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site53 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site53.Target((CallSite) Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site53, worksheet.get_Range((object) "A8", (object) "A8").Text);
        object obj4 = target4((CallSite) pSite52, obj3, "");
        if (target3((CallSite) pSite51, obj4))
          return "False(A8)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site54 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site54 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target5 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site54.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite54 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site54;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site55 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site55 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target6 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site55.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite55 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site55;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site56 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site56 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj5 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site56.Target((CallSite) Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site56, worksheet.get_Range((object) "B7", (object) "B7").Text);
        object obj6 = target6((CallSite) pSite55, obj5, "");
        if (target5((CallSite) pSite54, obj6))
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site57 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site57 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target7 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site57.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite57 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site57;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site58 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site58 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target8 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site58.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite58 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site58;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site59 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site59 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj7 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site59.Target((CallSite) Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site59, worksheet.get_Range((object) "B8", (object) "B8").Text);
        object obj8 = target8((CallSite) pSite58, obj7, "");
        if (target7((CallSite) pSite57, obj8))
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5a = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target9 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5a.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite5a = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5a;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5b = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target10 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5b.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite5b = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5b;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5c = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj9 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5c.Target((CallSite) Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5c, worksheet.get_Range((object) "C7", (object) "C7").Text);
        object obj10 = target10((CallSite) pSite5b, obj9, "");
        if (target9((CallSite) pSite5a, obj10))
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5d = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target11 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5d.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite5d = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5d;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5e = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target12 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite5e = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5e;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5f = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj11 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5f.Target((CallSite) Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site5f, worksheet.get_Range((object) "C8", (object) "C8").Text);
        object obj12 = target12((CallSite) pSite5e, obj11, "");
        if (target11((CallSite) pSite5d, obj12))
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site60 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site60 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target13 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site60.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite60 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site60;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site61 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site61 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target14 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site61.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite61 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site61;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site62 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site62 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj13 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site62.Target((CallSite) Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site62, worksheet.get_Range((object) "D7", (object) "D7").Text);
        object obj14 = target14((CallSite) pSite61, obj13, "");
        if (target13((CallSite) pSite60, obj14))
          return "False";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site63 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site63 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target15 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site63.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite63 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site63;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site64 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site64 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target16 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site64.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite64 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site64;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site65 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site65 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj15 = Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site65.Target((CallSite) Sec1.\u003Ccau9\u003Eo__SiteContainer4c.\u003C\u003Ep__Site65, worksheet.get_Range((object) "D8", (object) "D8").Text);
        object obj16 = target16((CallSite) pSite64, obj15, "");
        if (target15((CallSite) pSite63, obj16))
          return "False";
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
        if (Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site67 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site67 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site67.Target((CallSite) Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site67, d.Worksheets[(object) "Substitutes"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site68 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site68 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site68.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite68 = Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site68;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site69 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site69 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site69.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite69 = Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site69;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site6a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site6a = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site6a.Target((CallSite) Sec1.\u003Ccau10\u003Eo__SiteContainer66.\u003C\u003Ep__Site6a, worksheet.get_Range((object) "A1", (object) "A1").Text);
        object obj2 = target2((CallSite) pSite69, obj1, "Rank");
        if (target1((CallSite) pSite68, obj2))
          return "False";
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
        if (Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6c = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6c.Target((CallSite) Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6c, d.Worksheets[(object) "Quote"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6d = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6d.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite6d = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6d;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6e = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite6e = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6e;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6f = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6f.Target((CallSite) Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site6f, worksheet.get_Range((object) "E12", (object) "E12").Text);
        object obj2 = target2((CallSite) pSite6e, obj1, "Quantity");
        if (target1((CallSite) pSite6d, obj2))
          return "False (E12)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site70 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site70 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site70.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite70 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site70;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site71 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site71 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site71.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite71 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site71;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site72 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site72 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site72.Target((CallSite) Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site72, worksheet.get_Range((object) "F12", (object) "F12").Text);
        object obj4 = target4((CallSite) pSite71, obj3, "Unit Price ");
        if (target3((CallSite) pSite70, obj4))
          return "False (F12)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site73 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site73 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target5 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site73.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite73 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site73;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site74 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site74 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target6 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site74.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite74 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site74;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site75 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site75 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj5 = Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site75.Target((CallSite) Sec1.\u003Ccau11\u003Eo__SiteContainer6b.\u003C\u003Ep__Site75, worksheet.get_Range((object) "F30", (object) "F30").Text);
        object obj6 = target6((CallSite) pSite74, obj5, "Total");
        if (target5((CallSite) pSite73, obj6))
          return "False (F30)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau5_5(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site77 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site77 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site77.Target((CallSite) Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site77, d.Worksheets[(object) "Quote"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site78 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site78 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site78.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite78 = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site78;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site79 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site79 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site79.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite79 = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site79;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7a = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7a.Target((CallSite) Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7a, worksheet.get_Range((object) "E12", (object) "E12").Text);
        object obj2 = target2((CallSite) pSite79, obj1, "Quantity");
        if (target1((CallSite) pSite78, obj2))
          return "False (E12)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7b = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target3 = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7b.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite7b = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7b;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7c = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target4 = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7c.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite7c = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7c;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7d = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj3 = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7d.Target((CallSite) Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7d, worksheet.get_Range((object) "F12", (object) "F12").Text);
        object obj4 = target4((CallSite) pSite7c, obj3, "Unit Price ");
        if (target3((CallSite) pSite7b, obj4))
          return "False (F12)";
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7e = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target5 = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite7e = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7e;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7f = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target6 = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7f.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite7f = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site7f;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site80 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site80 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj5 = Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site80.Target((CallSite) Sec1.\u003Ccau5_5\u003Eo__SiteContainer76.\u003C\u003Ep__Site80, worksheet.get_Range((object) "F30", (object) "F30").Text);
        object obj6 = target6((CallSite) pSite7f, obj5, "Total");
        if (target5((CallSite) pSite7e, obj6))
          return "False (F30)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string goto001(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site82 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site82 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site82.Target((CallSite) Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site82, d.Worksheets[(object) "Donor List"]);
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site83 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site83 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site83.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite83 = Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site83;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site84 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site84 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site84.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite84 = Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site84;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site85 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site85 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site85.Target((CallSite) Sec1.\u003Cgoto001\u003Eo__SiteContainer81.\u003C\u003Ep__Site85, worksheet.get_Range((object) "G8", (object) "G8").Text);
        object obj2 = target2((CallSite) pSite84, obj1, "");
        if (target1((CallSite) pSite83, obj2))
          return "False";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string goto002(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site87 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site87 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec1)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site87.Target((CallSite) Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site87, d.Worksheets[(object) "Demographics"]);
      }
      catch (Exception ex)
      {
        return "Fales (Ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site88 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site88 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site88.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite88 = Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site88;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site89 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site89 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site89.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite89 = Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site89;
        // ISSUE: reference to a compiler-generated field
        if (Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site8a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site8a = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec1), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site8a.Target((CallSite) Sec1.\u003Cgoto002\u003Eo__SiteContainer86.\u003C\u003Ep__Site8a, worksheet.get_Range((object) "C5", (object) "C5").get_Value((object) Missing.Value));
        object obj2 = target2((CallSite) pSite89, obj1, "300");
        if (target1((CallSite) pSite88, obj2))
          return "False(C5)";
      }
      catch (Exception ex)
      {
        return "False (something wrong)";
      }
      return "True";
    }
  }
}
