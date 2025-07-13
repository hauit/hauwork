// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Sec9
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
  internal class Sec9
  {
    public static string CheckCau(int causo, Application a, Workbook d)
    {
      switch (causo)
      {
        case 1:
          return Sec9.cau17(a, d);
        case 2:
          return Sec9.cau8(a, d);
        case 3:
          return Sec9.cau16(a, d);
        case 4:
          return Sec9.cau1(a, d);
        case 5:
          return Sec9.cau2(a, d);
        case 6:
          return Sec9.cau7(a, d);
        case 7:
          return Sec9.cau13(a, d);
        case 8:
          return Sec9.cau14(a, d);
        case 9:
          return Sec9.cau3(a, d);
        case 10:
          return Sec9.cau10(a, d);
        case 11:
          return Sec9.cau11(a, d);
        case 12:
          return Sec9.cau12(a, d);
        case 13:
          return Sec9.cau13(a, d);
        case 14:
          return Sec9.cau14(a, d);
        case 15:
          return Sec9.cau15(a, d);
        case 16:
          return Sec9.cau16(a, d);
        case 17:
          return Sec9.cau17(a, d);
        case 18:
          return Sec9.cau18(a, d);
        case 19:
          return Sec9.cau19(a, d);
        case 20:
          return Sec9.cau20(a, d);
        default:
          return "False";
      }
    }

    private static string cau1(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, d.Worksheets[(object) "Materials"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.PageSetup.Orientation != XlPageOrientation.xlLandscape)
          return "False(Landscape)";
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 = CallSite<Func<CallSite, object, string>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string> target = Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string>> pSite2 = Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        object obj = Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3.Target((CallSite) Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3, worksheet.PageSetup.FitToPagesTall);
        string str = target((CallSite) pSite2, obj);
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 = CallSite<Func<CallSite, object, int>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (int), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4.Target((CallSite) Sec9.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4, worksheet.PageSetup.FitToPagesWide) != 1)
          return "False(Wide=1)";
        if (str != "False")
          return "False(Tall=0)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau2(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau2\u003Eo__SiteContainer5.\u003C\u003Ep__Site6 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau2\u003Eo__SiteContainer5.\u003C\u003Ep__Site6 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau2\u003Eo__SiteContainer5.\u003C\u003Ep__Site6.Target((CallSite) Sec9.\u003Ccau2\u003Eo__SiteContainer5.\u003C\u003Ep__Site6, d.Worksheets[(object) "roster"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.PageSetup.PrintTitleRows != "$7:$7")
          return "False(row 7)";
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
        object documentProperties = d.BuiltinDocumentProperties;
        object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder) null, documentProperties, new object[1]
        {
          (object) "Company"
        });
        if (target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder) null, target, new object[0]).ToString() != "Lucerne Publishing")
          return "False (Lucerne Publishing)";
      }
      catch (Exception ex)
      {
        return "False(add company)";
      }
      return "True";
    }

    private static string cau4(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site8 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site8 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site8.Target((CallSite) Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site8, d.Worksheets[(object) "Living Online"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site9 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site9 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site9.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite9 = Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Site9;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Sitea == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Sitea = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Sitea.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSitea = Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Sitea;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Siteb == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Siteb = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Siteb.Target((CallSite) Sec9.\u003Ccau4\u003Eo__SiteContainer7.\u003C\u003Ep__Siteb, worksheet.get_Range((object) "E32", (object) "E32").Text);
        object obj2 = target2((CallSite) pSitea, obj1, "=AVERAGE(E2:E30)");
        if (target1((CallSite) pSite9, obj2))
          return "Fasle(file->option show formular)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
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
        if (Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sited == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sited = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sited.Target((CallSite) Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sited, d.Worksheets[(object) "Q2 Sales"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSitee = Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSitef = Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Site10 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Site10 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Site10.Target((CallSite) Sec9.\u003Ccau5\u003Eo__SiteContainerc.\u003C\u003Ep__Site10, worksheet.get_Range((object) "F6", (object) "F6").Text);
        object obj2 = target2((CallSite) pSitef, obj1, "=AVERAGE(Table2[@[April]:[June]])");
        if (target1((CallSite) pSitee, obj2))
          return "Fasle(file->option show formular)";
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
        object documentProperties = d.BuiltinDocumentProperties;
        object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder) null, documentProperties, new object[1]
        {
          (object) "Title"
        });
        string str = target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder) null, target, new object[0]).ToString();
        if (str != "2015")
          return "False (" + str + ")";
      }
      catch (Exception ex)
      {
        return "False(add title)";
      }
      return "True";
    }

    private static string cau7(Application a, Workbook d)
    {
      try
      {
        foreach (Worksheet worksheet in d.Worksheets)
        {
          try
          {
            // ISSUE: reference to a compiler-generated field
            if (Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site12 == null)
            {
              // ISSUE: reference to a compiler-generated field
              Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site12 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
              {
                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
              }));
            }
            // ISSUE: reference to a compiler-generated field
            Func<CallSite, object, bool> target1 = Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site12.Target;
            // ISSUE: reference to a compiler-generated field
            CallSite<Func<CallSite, object, bool>> pSite12 = Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site12;
            // ISSUE: reference to a compiler-generated field
            if (Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site13 == null)
            {
              // ISSUE: reference to a compiler-generated field
              Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site13 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
              {
                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
              }));
            }
            // ISSUE: reference to a compiler-generated field
            Func<CallSite, object, string, object> target2 = Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site13.Target;
            // ISSUE: reference to a compiler-generated field
            CallSite<Func<CallSite, object, string, object>> pSite13 = Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site13;
            // ISSUE: reference to a compiler-generated field
            if (Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site14 == null)
            {
              // ISSUE: reference to a compiler-generated field
              Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site14 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
              {
                CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
              }));
            }
            // ISSUE: reference to a compiler-generated field
            // ISSUE: reference to a compiler-generated field
            object obj1 = Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site14.Target((CallSite) Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site14, worksheet.PageSetup.Zoom);
            object obj2 = target2((CallSite) pSite13, obj1, "False");
            if (target1((CallSite) pSite12, obj2))
              return "False(" + worksheet.Name + ")";
            // ISSUE: reference to a compiler-generated field
            if (Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site15 == null)
            {
              // ISSUE: reference to a compiler-generated field
              Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site15 = CallSite<Func<CallSite, object, int>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (int), typeof (Sec9)));
            }
            // ISSUE: reference to a compiler-generated field
            // ISSUE: reference to a compiler-generated field
            if (Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site15.Target((CallSite) Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site15, worksheet.PageSetup.FitToPagesWide) != 1)
              return "False(Wide=1 of " + worksheet.Name + ")";
            // ISSUE: reference to a compiler-generated field
            if (Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site16 == null)
            {
              // ISSUE: reference to a compiler-generated field
              Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site16 = CallSite<Func<CallSite, object, int>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (int), typeof (Sec9)));
            }
            // ISSUE: reference to a compiler-generated field
            // ISSUE: reference to a compiler-generated field
            if (Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site16.Target((CallSite) Sec9.\u003Ccau7\u003Eo__SiteContainer11.\u003C\u003Ep__Site16, worksheet.PageSetup.FitToPagesTall) != 1)
              return "False(Tall=1 of " + worksheet.Name + ")";
          }
          catch (Exception ex)
          {
            return "False (Something not finish!)";
          }
        }
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      return "True";
    }

    private static string cau8(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau8\u003Eo__SiteContainer17.\u003C\u003Ep__Site18 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau8\u003Eo__SiteContainer17.\u003C\u003Ep__Site18 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau8\u003Eo__SiteContainer17.\u003C\u003Ep__Site18.Target((CallSite) Sec9.\u003Ccau8\u003Eo__SiteContainer17.\u003C\u003Ep__Site18, d.Worksheets[(object) "Inbound call"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.PageSetup.PrintArea != "$A$1:$C$19")
          return "False($A$1:$C$19)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
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
        if (Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1a = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1a.Target((CallSite) Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1a, d.Worksheets[(object) "Q2 Sales"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1b == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1b = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1b.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite1b = Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1b;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1c = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1c.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite1c = Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1c;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1d = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1d.Target((CallSite) Sec9.\u003Ccau9\u003Eo__SiteContainer19.\u003C\u003Ep__Site1d, worksheet.get_Range((object) "F6", (object) "F6").Text);
        object obj2 = target2((CallSite) pSite1c, obj1, "=AVERAGE(Table2[@[April]:[June]])");
        if (target1((CallSite) pSite1b, obj2))
          return "Fasle(file->option show formular)";
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
        object documentProperties = d.BuiltinDocumentProperties;
        object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder) null, documentProperties, new object[1]
        {
          (object) "Subject"
        });
        if (target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder) null, target, new object[0]).ToString() != "")
          return "False";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
      }
      return "True";
    }

    private static string cau11(Application a, Workbook d)
    {
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau11\u003Eo__SiteContainer1e.\u003C\u003Ep__Site1f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau11\u003Eo__SiteContainer1e.\u003C\u003Ep__Site1f = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: variable of a compiler-generated type
        Worksheet worksheet = Sec9.\u003Ccau11\u003Eo__SiteContainer1e.\u003C\u003Ep__Site1f.Target((CallSite) Sec9.\u003Ccau11\u003Eo__SiteContainer1e.\u003C\u003Ep__Site1f, d.Worksheets[(object) "Expenses"]);
        if (worksheet.PageSetup.PrintArea != "$B$5:$D$52")
          return "False($B$5:$D$52)";
      }
      catch (Exception ex)
      {
        return "Fales (Tên Trang Tính)";
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
        if (Sec9.\u003Ccau12\u003Eo__SiteContainer20.\u003C\u003Ep__Site21 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau12\u003Eo__SiteContainer20.\u003C\u003Ep__Site21 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau12\u003Eo__SiteContainer20.\u003C\u003Ep__Site21.Target((CallSite) Sec9.\u003Ccau12\u003Eo__SiteContainer20.\u003C\u003Ep__Site21, d.Worksheets[(object) "Scholarships"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.PageSetup.PrintTitleColumns != "$A:$A")
          return "False($A:$A)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
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
        if (Sec9.\u003Ccau13\u003Eo__SiteContainer22.\u003C\u003Ep__Site23 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau13\u003Eo__SiteContainer22.\u003C\u003Ep__Site23 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau13\u003Eo__SiteContainer22.\u003C\u003Ep__Site23.Target((CallSite) Sec9.\u003Ccau13\u003Eo__SiteContainer22.\u003C\u003Ep__Site23, d.Worksheets[(object) "Q2 Sales"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.PageSetup.BottomMargin != 54.0)
          return "False(Build In)";
        if (worksheet.PageSetup.LeftMargin != 18.0)
          return "False(Build In)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau14(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau14\u003Eo__SiteContainer24.\u003C\u003Ep__Site25 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau14\u003Eo__SiteContainer24.\u003C\u003Ep__Site25 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau14\u003Eo__SiteContainer24.\u003C\u003Ep__Site25.Target((CallSite) Sec9.\u003Ccau14\u003Eo__SiteContainer24.\u003C\u003Ep__Site25, d.Worksheets[(object) "Games"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.PageSetup.TopMargin != 72.0)
          return "False(Top)";
        if (worksheet.PageSetup.BottomMargin != 72.0)
          return "False(Bottom)";
        if (worksheet.PageSetup.LeftMargin != 108.0)
          return "False(Left)";
        if (worksheet.PageSetup.RightMargin != 108.0)
          return "False(Right)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau15(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet1;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau15\u003Eo__SiteContainer26.\u003C\u003Ep__Site27 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau15\u003Eo__SiteContainer26.\u003C\u003Ep__Site27 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet1 = Sec9.\u003Ccau15\u003Eo__SiteContainer26.\u003C\u003Ep__Site27.Target((CallSite) Sec9.\u003Ccau15\u003Eo__SiteContainer26.\u003C\u003Ep__Site27, d.Worksheets[(object) "Games"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet1.ListObjects[(object) "Table2"].AlternativeText != "data")
          return "False(data)";
        if (worksheet1.ListObjects[(object) "Table3"].AlternativeText != "data")
          return "False(data)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet2;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau15\u003Eo__SiteContainer26.\u003C\u003Ep__Site28 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau15\u003Eo__SiteContainer26.\u003C\u003Ep__Site28 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet2 = Sec9.\u003Ccau15\u003Eo__SiteContainer26.\u003C\u003Ep__Site28.Target((CallSite) Sec9.\u003Ccau15\u003Eo__SiteContainer26.\u003C\u003Ep__Site28, d.Worksheets[(object) "Shareholders Info"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet2.ListObjects[(object) "Table1"].AlternativeText != "data")
          return "False(data)";
        // ISSUE: reference to a compiler-generated method
        if (worksheet2.Shapes.Item((object) "Chart 1").Title != "data")
          return "False(data)";
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
        object documentProperties = d.BuiltinDocumentProperties;
        object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (System.Reflection.Binder) null, documentProperties, new object[1]
        {
          (object) "Subject"
        });
        if (target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (System.Reflection.Binder) null, target, new object[0]).ToString() != "")
          return "False";
      }
      catch (Exception ex)
      {
        return "False(something wrong)";
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
        if (Sec9.\u003Ccau17\u003Eo__SiteContainer29.\u003C\u003Ep__Site2a == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau17\u003Eo__SiteContainer29.\u003C\u003Ep__Site2a = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau17\u003Eo__SiteContainer29.\u003C\u003Ep__Site2a.Target((CallSite) Sec9.\u003Ccau17\u003Eo__SiteContainer29.\u003C\u003Ep__Site2a, d.Worksheets[(object) "January"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        if (worksheet.PageSetup.PrintArea != "$A$4:$F$20")
          return "Fasle(vùng in là $A$4:$F$20)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
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
        if (Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2c == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2c = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2c.Target((CallSite) Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2c, d.Worksheets[(object) "Q2 Sales"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2d == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2d = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2d.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite2d = Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2d;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2e == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2e = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2e.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite2e = Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2e;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2f == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2f = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2f.Target((CallSite) Sec9.\u003Ccau18\u003Eo__SiteContainer2b.\u003C\u003Ep__Site2f, worksheet.get_Range((object) "F6", (object) "F6").Text);
        object obj2 = target2((CallSite) pSite2e, obj1, "=AVERAGE(Table2[@[April]:[June]])");
        if (target1((CallSite) pSite2d, obj2))
          return "Fasle(file->option show formular)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
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
        if (Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site31 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site31 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site31.Target((CallSite) Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site31, d.Worksheets[(object) "Q2 Sales"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site32 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site32 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site32.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite32 = Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site32;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site33 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site33 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site33.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite33 = Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site33;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site34 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site34 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site34.Target((CallSite) Sec9.\u003Ccau19\u003Eo__SiteContainer30.\u003C\u003Ep__Site34, worksheet.get_Range((object) "F6", (object) "F6").Text);
        object obj2 = target2((CallSite) pSite33, obj1, "=AVERAGE(Table2[@[April]:[June]])");
        if (target1((CallSite) pSite32, obj2))
          return "Fasle(file->option show formular)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }

    private static string cau20(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site36 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site36 = CallSite<Func<CallSite, object, Worksheet>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec9)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site36.Target((CallSite) Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site36, d.Worksheets[(object) "Q2 Sales"]);
      }
      catch (Exception ex)
      {
        return "Fales (ten trang tinh)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site37 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site37 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site37.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite37 = Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site37;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site38 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site38 = CallSite<Func<CallSite, object, string, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site38.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite38 = Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site38;
        // ISSUE: reference to a compiler-generated field
        if (Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site39 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site39 = CallSite<Func<CallSite, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec9), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        object obj1 = Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site39.Target((CallSite) Sec9.\u003Ccau20\u003Eo__SiteContainer35.\u003C\u003Ep__Site39, worksheet.get_Range((object) "F6", (object) "F6").Text);
        object obj2 = target2((CallSite) pSite38, obj1, "=AVERAGE(Table2[@[April]:[June]])");
        if (target1((CallSite) pSite37, obj2))
          return "Fasle(file->option show formular)";
      }
      catch (Exception ex)
      {
        return "False (Something not finish!)";
      }
      return "True";
    }
  }
}
