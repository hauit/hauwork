// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S0_SlideMaster
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;

#nullable disable
namespace WindowsFormsApplication1;

internal class S0_SlideMaster
{
  public static string CheckCau(int causo, Application a, Presentation d)
  {
    switch (causo)
    {
      case 1:
        return S0_SlideMaster.cau1(a, d);
      case 2:
        return S0_SlideMaster.cau2(a, d);
      case 3:
        return S0_SlideMaster.cau3(a, d);
      case 4:
        return S0_SlideMaster.cau4(a, d);
      case 5:
        return S0_SlideMaster.cau5(a, d);
      case 6:
        return S0_SlideMaster.cau6(a, d);
      case 7:
        return S0_SlideMaster.cau7(a, d);
      case 8:
        return S0_SlideMaster.cau8(a, d);
      case 9:
        return S0_SlideMaster.cau9(a, d);
      case 10:
        return S0_SlideMaster.cau10(a, d);
      default:
        return "False (case defalt)";
    }
  }

  private static string cau1(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.SlideMaster.CustomLayouts.Count < 12)
        return "False(them 1 layout)";
      for (int Index = 1; Index <= a.ActivePresentation.SlideMaster.CustomLayouts.Count; ++Index)
      {
        if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Name == "Custom1")
        {
          if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Shapes.Count.ToString() != "6")
            return "False(khong them du hoac thieu placeholder)";
          if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Shapes[(object) 5].PlaceholderFormat.Type.ToString() != "ppPlaceholderPicture")
            return "False(them placeholder truoc)";
          return a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Shapes[(object) 6].PlaceholderFormat.Type.ToString() != "ppPlaceholderBody" ? "False(them placeholder text sau)" : "True";
        }
      }
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "False(ten layout)";
  }

  private static string cau2(Application a, Presentation d)
  {
    try
    {
      // ISSUE: reference to a compiler-generated method
      if (a.ActivePresentation.SlideMaster.Shapes[(object) "Text Placeholder 2"].TextFrame.TextRange.Lines(1, 1).ParagraphFormat.Bullet.Type.ToString() != "ppBulletPicture")
        return "False(picture)";
      if (a.ActivePresentation.SlideMaster.Shapes[(object) "Text Placeholder 2"].TextFrame.TextRange.ParagraphFormat.Bullet.Type.ToString() == "ppBulletPicture")
        return "False(chi dong dau)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string cau3(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.SlideMaster.CustomLayouts.Count < 12)
        return "False(them 1 layout)";
      for (int Index = 1; Index <= a.ActivePresentation.SlideMaster.CustomLayouts.Count; ++Index)
      {
        if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Name == "Trevorslayout")
        {
          if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Shapes.Count.ToString() != "6")
            return "False(khong them du hoac thieu placeholder)";
          if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Shapes[(object) 5].PlaceholderFormat.Type.ToString() != "ppPlaceholderPicture")
            return "False(them placeholder truoc)";
          return a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Shapes[(object) 6].PlaceholderFormat.Type.ToString() != "ppPlaceholderBody" ? "False(them placeholder text sau)" : "True";
        }
      }
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "False(ten layout)";
  }

  private static string cau4(Application a, Presentation d)
  {
    try
    {
      string str = "True";
      if (a.ActivePresentation.HandoutMaster.Shapes.Count != 4)
        return "False (khong them xoa shape)";
      try
      {
        // ISSUE: variable of a compiler-generated type
        Microsoft.Office.Interop.PowerPoint.Shape shape = a.ActivePresentation.HandoutMaster.Shapes[(object) 3];
        if (shape.TextFrame.TextRange.Text != "First Copy")
          return "False (First Copy o Handout footer)";
      }
      catch (Exception ex)
      {
        return "False (not picture)";
      }
      return str;
    }
    catch (Exception ex)
    {
      return "False (Something not finish!)";
    }
  }

  private static string cau5(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) 3].Shapes.Count.ToString() != "5")
        return "False(khong them du hoac thieu placeholder)";
      if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) 3].Shapes[(object) 5].PlaceholderFormat.Type.ToString() != "ppPlaceholderMediaClip")
        return "False(Media)";
      if ((double) a.ActivePresentation.SlideMaster.CustomLayouts[(object) 3].Shapes[(object) 5].Left != (double) a.ActivePresentation.SlideMaster.CustomLayouts[(object) 3].Shapes[(object) 1].Left)
        return "False(Align Left)";
      if ((double) a.ActivePresentation.SlideMaster.CustomLayouts[(object) 3].Shapes[(object) 5].Width != (double) a.ActivePresentation.SlideMaster.CustomLayouts[(object) 3].Shapes[(object) 1].Width)
        return "False(Align Right)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string cau6(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.SlideMaster.CustomLayouts.Count != 12)
        return "False(khong them xoa layout)";
      for (int Index = 1; Index <= a.ActivePresentation.SlideMaster.CustomLayouts.Count; ++Index)
      {
        if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Name == "Ingredients")
        {
          if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Shapes.Count.ToString() != "5")
            return "False(khong them du hoac thieu placeholder)";
          // ISSUE: reference to a compiler-generated method
          return a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Shapes[(object) "Content Placeholder 6"].TextFrame.TextRange.Paragraphs(1, 1).ParagraphFormat.Bullet.Type.ToString() == "ppBulletPicture" ? "True" : "False(Bullet bằng Picture)";
        }
      }
    }
    catch (Exception ex)
    {
      return "False (Something not finish!)";
    }
    return "False(tên layout)";
  }

  private static string cau7(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.SlideMaster.CustomLayouts.Count < 12)
        return "False(duplicate layout)";
      for (int Index = 1; Index <= a.ActivePresentation.SlideMaster.CustomLayouts.Count; ++Index)
      {
        if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Name == "Picture with Text")
        {
          if (a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Shapes.Count.ToString() != "5")
            return "False(sai placeholder hoặc sai Layout)";
          string str = a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Shapes[(object) 5].Name + a.ActivePresentation.SlideMaster.CustomLayouts[(object) Index].Shapes[(object) 4].Name;
          if (str.Contains("Text") && str.Contains("Picture"))
            return "True";
        }
      }
    }
    catch (Exception ex)
    {
      return "False (Something not finish!)";
    }
    return "False(tên layout)";
  }

  private static string cau8(Application a, Presentation d)
  {
    try
    {
      string text1 = a.ActivePresentation.HandoutMaster.Shapes[(object) "Header Placeholder 1"].TextFrame.TextRange.Text;
      if (text1 != "First Up Consultants")
        return $"False({text1}: viet sai)";
      string text2 = a.ActivePresentation.HandoutMaster.Shapes[(object) "Footer Placeholder 3"].TextFrame.TextRange.Text;
      if (text2 != "www.firstupconsultants.com")
        return $"False({text2}: viet sai)";
    }
    catch (Exception ex)
    {
      return "False (Something not finish!)";
    }
    return "True";
  }

  private static string cau9(Application a, Presentation d)
  {
    try
    {
      string str = "True";
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.Layout == PpSlideLayout.ppLayoutTitle)
        {
          if (slide.HeadersFooters.SlideNumber.Visible != MsoTriState.msoFalse)
            return "False (Title slide)";
        }
        else if (slide.HeadersFooters.SlideNumber.Visible != MsoTriState.msoTrue)
          return "False (slide number)";
      }
      return str;
    }
    catch (Exception ex)
    {
      return "False (Something not finish!)";
    }
  }

  private static string cau10(Application a, Presentation d)
  {
    try
    {
      string str = "True";
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.Layout == PpSlideLayout.ppLayoutTitle)
        {
          if (slide.HeadersFooters.SlideNumber.Visible != MsoTriState.msoFalse)
            return "False (Title slide)";
        }
        else if (slide.HeadersFooters.SlideNumber.Visible != MsoTriState.msoTrue)
          return "False (slide number)";
      }
      return str;
    }
    catch (Exception ex)
    {
      return "False (Something not finish!)";
    }
  }
}
