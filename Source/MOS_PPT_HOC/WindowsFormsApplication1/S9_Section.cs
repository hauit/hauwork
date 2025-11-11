// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S9_Section
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Runtime.CompilerServices;

#nullable disable
namespace WindowsFormsApplication1;

internal class S9_Section
{
  public static string CheckCau(int causo, Application a, Presentation d)
  {
    switch (causo)
    {
      case 1:
        return S9_Section.cau1(a, d);
      case 2:
        return S9_Section.cau2(a, d);
      case 3:
        return S9_Section.cau3(a, d);
      case 4:
        return S9_Section.cau4(a, d);
      case 5:
        return S9_Section.cau5(a, d);
      case 6:
        return S9_Section.cau6(a, d);
      case 7:
        return S9_Section.cau7(a, d);
      case 8:
        return S9_Section.cau8(a, d);
      case 9:
        return S9_Section.cau9(a, d);
      case 10:
        return S9_Section.cau10(a, d);
      default:
        return "case 11";
    }
  }

  private static string cau1(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.SectionProperties.Count != 1)
        return "False(add section)";
      // ISSUE: reference to a compiler-generated method
      return a.ActivePresentation.SectionProperties.Name(1) != "Title" ? "False(Title)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something Wrong)";
    }
  }

  private static string cau2(Application a, Presentation d)
  {
    try
    {
      int sectionIndex1 = -1;
      for (int sectionIndex2 = 1; sectionIndex2 <= a.ActivePresentation.SectionProperties.Count; ++sectionIndex2)
      {
        // ISSUE: reference to a compiler-generated method
        if (a.ActivePresentation.SectionProperties.Name(sectionIndex2) == "Structures")
          sectionIndex1 = sectionIndex2;
      }
      if (sectionIndex1 == -1)
        return "False(không có section Structures )";
      // ISSUE: reference to a compiler-generated method
      int num = a.ActivePresentation.SectionProperties.FirstSlide(sectionIndex1);
      if (num.ToString() != "3")
        return "False(bắt đầu section structures phải ở slide 3)";
      // ISSUE: reference to a compiler-generated method
      num = a.ActivePresentation.SectionProperties.SlidesCount(sectionIndex1);
      return num.ToString() != "2" ? "False(structures chỉ có 2 slide)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something Wrong)";
    }
  }

  private static string cau3(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 2].Shapes.Count <= 1)
        return "False(add Section Zoom)";
      return !a.ActivePresentation.Slides[(object) 2].Shapes[(object) 2].Name.Contains("Section Zoom") ? "False (add Section Zoom chứ không phải đối tượng khác)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something Wrong)";
    }
  }

  private static string cau4(Application a, Presentation d)
  {
    try
    {
      int sectionIndex1 = -1;
      for (int sectionIndex2 = 1; sectionIndex2 <= a.ActivePresentation.SectionProperties.Count; ++sectionIndex2)
      {
        // ISSUE: reference to a compiler-generated method
        if (a.ActivePresentation.SectionProperties.Name(sectionIndex2) == "Student Clubs")
          sectionIndex1 = sectionIndex2;
      }
      if (sectionIndex1 == -1)
        return "False(không có section Student Clubs )";
      // ISSUE: reference to a compiler-generated method
      int num = a.ActivePresentation.SectionProperties.FirstSlide(sectionIndex1);
      if (num.ToString() != "3")
        return "False(bắt đầu Student Clubs phải ở slide 3)";
      // ISSUE: reference to a compiler-generated method
      num = a.ActivePresentation.SectionProperties.SlidesCount(sectionIndex1);
      return num.ToString() != "5" ? "False(structures chỉ có 5 slide)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something Wrong)";
    }
  }

  private static string cau5(Application a, Presentation d)
  {
    try
    {
      int count = a.ActivePresentation.SlideShowSettings.NamedSlideShows.Count;
      if (count.ToString() != "1")
        return "False (number of SlideShow)";
      if (a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object) 1].Name != "Internal")
        return "False (SlideShow name)";
      count = a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object) 1].Count;
      if (count.ToString() != "4")
        return "False (Number of slide in slide show)";
      // ISSUE: reference to a compiler-generated field
      if (S9_Section.\u003Ccau5\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S9_Section.\u003Ccau5\u003Eo__SiteContainer0.\u003C\u003Ep__Site1 = CallSite<Func<CallSite, object, Array>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Array), typeof (S9_Section)));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      return S9_Section.\u003Ccau5\u003Eo__SiteContainer0.\u003C\u003Ep__Site1.Target((CallSite) S9_Section.\u003Ccau5\u003Eo__SiteContainer0.\u003C\u003Ep__Site1, a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object) 1].SlideIDs).GetValue(4).ToString() != "257" ? "False (Strategy 2020 slide to end)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something Wrong)";
    }
  }

  private static string cau6(Application a, Presentation d)
  {
    try
    {
      int slideId = a.ActivePresentation.Slides[(object) 7].SlideID;
      if (slideId.ToString() != "258")
        return "False (Position slide Advantage)";
      slideId = a.ActivePresentation.Slides[(object) 8].SlideID;
      return slideId.ToString() != "263" ? "False (Position slide 7)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something Wrong)";
    }
  }

  private static string cau7(Application a, Presentation d)
  {
    try
    {
      int num = a.ActivePresentation.SectionProperties.Count;
      if (num.ToString() != "3")
        return "False (number of section)";
      // ISSUE: reference to a compiler-generated method
      if (a.ActivePresentation.SectionProperties.Name(1) != "Rolling Ware")
        return "False (Section 1 name)";
      // ISSUE: reference to a compiler-generated method
      if (a.ActivePresentation.SectionProperties.Name(2) != "Critical Points")
        return "False (Section 2 name)";
      num = a.ActivePresentation.Slides[(object) 1].sectionIndex;
      if (num.ToString() != "1")
        return "False (Position section 1)";
      num = a.ActivePresentation.Slides[(object) 2].sectionIndex;
      if (num.ToString() != "1")
        return "False (Position section 1)";
      num = a.ActivePresentation.Slides[(object) 3].sectionIndex;
      return num.ToString() != "2" ? "False (Position section 2)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something Wrong)";
    }
  }

  private static string cau8(Application a, Presentation d)
  {
    try
    {
      int count = a.ActivePresentation.SlideShowSettings.NamedSlideShows.Count;
      if (count.ToString() != "1")
        return "False (number of SlideShow)";
      if (a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object) 1].Name != "Internal")
        return "False (SlideShow name)";
      count = a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object) 1].Count;
      if (count.ToString() != "3")
        return "False (Number of slide in slide show)";
      // ISSUE: reference to a compiler-generated field
      if (S9_Section.\u003Ccau8\u003Eo__SiteContainer2.\u003C\u003Ep__Site3 == null)
      {
        // ISSUE: reference to a compiler-generated field
        S9_Section.\u003Ccau8\u003Eo__SiteContainer2.\u003C\u003Ep__Site3 = CallSite<Func<CallSite, object, Array>>.Create(Binder.Convert(CSharpBinderFlags.ConvertExplicit, typeof (Array), typeof (S9_Section)));
      }
      // ISSUE: reference to a compiler-generated field
      // ISSUE: reference to a compiler-generated field
      Array array = S9_Section.\u003Ccau8\u003Eo__SiteContainer2.\u003C\u003Ep__Site3.Target((CallSite) S9_Section.\u003Ccau8\u003Eo__SiteContainer2.\u003C\u003Ep__Site3, a.ActivePresentation.SlideShowSettings.NamedSlideShows[(object) 1].SlideIDs);
      if (array.GetValue(1).ToString() != "259")
        return "False (3)";
      if (array.GetValue(2).ToString() != "260")
        return "False (4)";
      return array.GetValue(3).ToString() != "261" ? "False (5)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something Wrong)";
    }
  }

  private static string cau9(Application a, Presentation d) => "True";

  private static string cau10(Application a, Presentation d) => "True";
}
