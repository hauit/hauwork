// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S16
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

#nullable disable
namespace WindowsFormsApplication1;

internal class S16
{
  public static string CheckCau(int causo, Application a, Presentation d)
  {
    switch (causo)
    {
      case 1:
        return S16.Cau1(a, d);
      case 2:
        return S16.Cau2(a, d);
      case 3:
        return S16.Cau3(a, d);
      case 4:
        return S16.Cau4(a, d);
      case 5:
        return S16.Cau5(a, d);
      case 6:
        return S16.Cau6(a, d);
      case 7:
        return S16.Cau7(a, d);
      default:
        return "case out index";
    }
  }

  private static string Cau1(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 3].Shapes.Count != 3)
        return "False(insert Table)";
      if (((object) a.ActivePresentation.Slides[(object) 3].Shapes[(object) 3].Type).ToString() != "msoTable")
        return "False(Table)";
      int count = a.ActivePresentation.Slides[(object) 3].Shapes[(object) 3].Table.Columns.Count;
      if (count.ToString() != "3")
        return "False(3 cot)";
      count = a.ActivePresentation.Slides[(object) 3].Shapes[(object) 3].Table.Rows.Count;
      if (count.ToString() != "4")
        return "False(4 hang)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau2(Application a, Presentation d)
  {
    try
    {
      // ISSUE: reference to a compiler-generated field
      if (S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (S16)));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string> target1 = S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string>> pSite1 = S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site1;
      // ISSUE: reference to a compiler-generated field
      if (S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site2 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S16), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      object obj1 = S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site2.Target((CallSite) S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site2, a.ActivePresentation.Slides[(object) 8].Shapes[(object) "Chart 5"].Chart.ChartStyle);
      if (target1((CallSite) pSite1, obj1) != "261")
        return "False (style 11)";
      // ISSUE: reference to a compiler-generated field
      if (S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, string>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof (string), typeof (S16)));
      }
      // ISSUE: reference to a compiler-generated field
      Func<CallSite, object, string> target2 = S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site3.Target;
      // ISSUE: reference to a compiler-generated field
      CallSite<Func<CallSite, object, string>> pSite3 = S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site3;
      // ISSUE: reference to a compiler-generated field
      if (S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site4 = CallSite<Func<CallSite, object, object>>.Create(Binder.InvokeMember(CSharpBinderFlags.None, "ToString", (IEnumerable<Type>) null, typeof (S16), (IEnumerable<CSharpArgumentInfo>) new CSharpArgumentInfo[1]
        {
          CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, (string) null)
        }));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      object obj2 = S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site4.Target((CallSite) S16.\u003CCau2\u003Eo__SiteContainer0.\u003C\u003Ep__Site4, a.ActivePresentation.Slides[(object) 8].Shapes[(object) "Chart 5"].Chart.ChartColor);
      return target2((CallSite) pSite3, obj2) != "13" ? "False (palette 4)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something Wrong)";
    }
  }

  private static string Cau3(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.PrintOptions.NumberOfCopies.ToString() != "4")
        return "False(4 copy)";
      if (a.ActivePresentation.PrintOptions.OutputType.ToString() != "ppPrintOutputThreeSlideHandouts")
        return "False(3 slide/ 1 page)";
      return ((object) a.ActivePresentation.PrintOptions.Collate).ToString() != "msoFalse" ? "False(Un Collate)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Somthing Wrong)";
    }
  }

  private static string Cau4(Application a, Presentation d)
  {
    try
    {
      if (((object) a.ActivePresentation.Slides[(object) 6].BackgroundStyle).ToString() != "msoBackgroundStyleNotAPreset")
        return "False(slide 6)";
      if (a.ActivePresentation.Slides[(object) 6].Background.Fill.GradientAngle.ToString() != "90")
        return "False(default)";
    }
    catch (Exception ex)
    {
      return "False(Gradient)";
    }
    return "True";
  }

  private static string Cau5(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 7].Shapes[(object) "Content Placeholder 4"].Shadow.Blur.ToString() != "20")
        return "False(ShapeStyle)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau6(Application a, Presentation d) => "True";

  private static string Cau7(Application a, Presentation d) => "True";
}
