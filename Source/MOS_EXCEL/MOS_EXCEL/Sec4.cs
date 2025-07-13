// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.Sec4
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;

namespace WindowsFormsApplication1
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
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec4)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, d.Worksheets[(object) "Orders"]);
      }
      catch (Exception ex)
      {
        return "Fales (worrksheet)";
      }
      // ISSUE: variable of a compiler-generated type
      FormatConditions formatConditions;
      try
      {
        // ISSUE: reference to a compiler-generated method
        formatConditions = worksheet.get_Range((object) "G2", (object) "G526").FormatConditions;
      }
      catch (Exception ex)
      {
        return "False (Not FormatCondition)";
      }
      if (formatConditions.Count != 1)
        return "False (Number of FormatCondition)";
      // ISSUE: variable of a compiler-generated type
      AboveAverage aboveAverage;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 = CallSite<Func<CallSite, object, AboveAverage>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (AboveAverage), typeof (Sec4)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        aboveAverage = Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2.Target((CallSite) Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site2, formatConditions.Item((object) 1));
        if (aboveAverage.AboveBelow != XlAboveBelow.xlAboveAverage)
          return "False(AboveAverage)";
      }
      catch (Exception ex)
      {
        return "False(AboveAverage)";
      }
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec4), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, bool> target1 = Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, bool>> pSite3 = Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site3;
        // ISSUE: reference to a compiler-generated field
        if (Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec4), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        Func<CallSite, object, string, object> target2 = Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4.Target;
        // ISSUE: reference to a compiler-generated field
        CallSite<Func<CallSite, object, string, object>> pSite4 = Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site4;
        // ISSUE: reference to a compiler-generated field
        if (Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec4), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
          {
            CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
          }));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        object obj1 = Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5.Target((CallSite) Sec4.\u003Ccau1\u003Eo__SiteContainer0.\u003C\u003Ep__Site5, aboveAverage.Font.Color);
        object obj2 = target2((CallSite) pSite4, obj1, "24832");
        if (target1((CallSite) pSite3, obj2))
          return "False(sai format)";
      }
      catch (Exception ex)
      {
        return "False(somthing wrong)";
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
        if (Sec4.\u003Ccau2\u003Eo__SiteContainer6.\u003C\u003Ep__Site7 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec4.\u003Ccau2\u003Eo__SiteContainer6.\u003C\u003Ep__Site7 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec4)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec4.\u003Ccau2\u003Eo__SiteContainer6.\u003C\u003Ep__Site7.Target((CallSite) Sec4.\u003Ccau2\u003Eo__SiteContainer6.\u003C\u003Ep__Site7, d.Worksheets[(object) "London"]);
      }
      catch (Exception ex)
      {
        return "Fales ()";
      }
      // ISSUE: variable of a compiler-generated type
      FormatConditions formatConditions;
      try
      {
        // ISSUE: reference to a compiler-generated method
        formatConditions = worksheet.get_Range((object) "D5", (object) "D21").FormatConditions;
      }
      catch (Exception ex)
      {
        return "False (Not FormatCondition)";
      }
      return formatConditions.Count != 0 ? "False (Chua xoa bo dinh dang co dieu kien!)" : "True";
    }

    private static string cau3(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Site9 == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Site9 = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec4)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Site9.Target((CallSite) Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Site9, d.Worksheets[(object) "Products"]);
      }
      catch (Exception ex)
      {
        return "Fales (Not found Worksheet)";
      }
      // ISSUE: variable of a compiler-generated type
      FormatConditions formatConditions;
      try
      {
        // ISSUE: reference to a compiler-generated method
        formatConditions = worksheet.get_Range((object) "E3", (object) "E54").FormatConditions;
      }
      catch (Exception ex)
      {
        return "False (Not FormatCondition)";
      }
      if (formatConditions.Count != 1)
        return "False (Number of FormatCondition)";
      // ISSUE: variable of a compiler-generated type
      IconSetCondition conSetCondition;
      // ISSUE: variable of a compiler-generated type
      IconSet conSet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea = CallSite<Func<CallSite, object, IconSetCondition>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (IconSetCondition), typeof (Sec4)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        conSetCondition = Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea.Target((CallSite) Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Sitea, formatConditions.Item((object) 1));
        // ISSUE: reference to a compiler-generated field
        if (Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb = CallSite<Func<CallSite, object, IconSet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (IconSet), typeof (Sec4)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        conSet = Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb.Target((CallSite) Sec4.\u003Ccau3\u003Eo__SiteContainer8.\u003C\u003Ep__Siteb, conSetCondition.IconSet);
      }
      catch (Exception ex)
      {
        return "False  (Not Iconset)";
      }
      if (conSet.ID != XlIconSet.xl3TrafficLights1)
        return "False (loại)";
      if (conSet.Count != 3)
        return "False (not 3 Icon)";
      if (conSetCondition.IconCriteria.Count != 3)
        return "False (not 3 Criteria)";
      if (conSetCondition.IconCriteria[(object) 1].Icon != XlIcon.xlIconRedCircleWithBorder)
        return "False (0FilledBoxes)";
      if (conSetCondition.IconCriteria[(object) 2].Icon != XlIcon.xlIconYellowCircle)
        return "False (1FilledBox)";
      if (conSetCondition.IconCriteria[(object) 3].Icon != XlIcon.xlIconGreenCircle)
        return "False (2FilledBoxes)";
      if (conSetCondition.IconCriteria[(object) 1].Type != XlConditionValueTypes.xlConditionValuePercent || conSetCondition.IconCriteria[(object) 2].Type != XlConditionValueTypes.xlConditionValuePercent || conSetCondition.IconCriteria[(object) 3].Type != XlConditionValueTypes.xlConditionValuePercent)
        return "False (type Number)";
      return conSetCondition.IconCriteria[(object) 1].Operator != 7 || conSetCondition.IconCriteria[(object) 2].Operator != 7 || conSetCondition.IconCriteria[(object) 3].Operator != 7 ? "False (Operator)" : "True";
    }

    private static string cau4(Application a, Workbook d)
    {
      // ISSUE: variable of a compiler-generated type
      Worksheet worksheet;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sited == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sited = CallSite<Func<CallSite, object, Worksheet>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Worksheet), typeof (Sec4)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        worksheet = Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sited.Target((CallSite) Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sited, d.Worksheets[(object) "Summary"]);
      }
      catch (Exception ex)
      {
        return "Fales (Not found Worksheet)";
      }
      // ISSUE: variable of a compiler-generated type
      FormatConditions formatConditions;
      try
      {
        // ISSUE: reference to a compiler-generated method
        formatConditions = worksheet.get_Range((object) "F4", (object) "F11").FormatConditions;
      }
      catch (Exception ex)
      {
        return "False (Not FormatCondition)";
      }
      if (formatConditions.Count != 1)
        return "False (Number of FormatCondition)";
      // ISSUE: variable of a compiler-generated type
      FormatCondition formatCondition;
      try
      {
        // ISSUE: reference to a compiler-generated field
        if (Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee == null)
        {
          // ISSUE: reference to a compiler-generated field
          Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee = CallSite<Func<CallSite, object, FormatCondition>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (FormatCondition), typeof (Sec4)));
        }
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated field
        // ISSUE: reference to a compiler-generated method
        formatCondition = Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee.Target((CallSite) Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitee, formatConditions.Item((object) 1));
      }
      catch (Exception ex)
      {
        return "False  (Not CellValue)";
      }
      if (formatCondition.Type != 1)
        return "False (Sai kiểu)";
      if (formatCondition.Operator != 5)
        return "False (chọn sai toán tử)";
      if (formatCondition.Formula1 != "=5000000")
        return "False (=5000000)";
      // ISSUE: reference to a compiler-generated field
      if (Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof (Sec4), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, bool> target1 = Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, bool>> pSitef = Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Sitef;
      // ISSUE: reference to a compiler-generated field
      if (Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site10 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site10 = CallSite<Func<CallSite, object, string, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof (Sec4), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[2]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null),
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.UseCompileTimeType | CSharpArgumentInfoFlags.Constant, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string, object> target2 = Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site10.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string, object>> pSite10 = Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site10;
      // ISSUE: reference to a compiler-generated field
      if (Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site11 == null)
      {
        // ISSUE: reference to a compiler-generated field
        Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site11 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (Sec4), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      object obj1 = Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site11.Target((CallSite) Sec4.\u003Ccau4\u003Eo__SiteContainerc.\u003C\u003Ep__Site11, formatCondition.Font.Color);
      object obj2 = target2((CallSite) pSite10, obj1, "22428");
      return target1((CallSite) pSitef, obj2) ? "False(sai kiểu định dạng)" : "True";
    }
  }
}
