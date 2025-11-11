// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S8_Transition
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Interop.PowerPoint;
using System;

#nullable disable
namespace WindowsFormsApplication1;

internal class S8_Transition
{
  public static string CheckCau(int causo, Application a, Presentation d)
  {
    switch (causo)
    {
      case 1:
        return S8_Transition.cau1(a, d);
      case 2:
        return S8_Transition.cau2(a, d);
      case 3:
        return S8_Transition.cau3(a, d);
      case 4:
        return S8_Transition.cau4(a, d);
      case 5:
        return S8_Transition.cau5(a, d);
      case 6:
        return S8_Transition.cau6(a, d);
      case 7:
        return S8_Transition.cau7(a, d);
      case 8:
        return S8_Transition.cau8(a, d);
      case 9:
        return S8_Transition.cau9(a, d);
      case 10:
        return S8_Transition.cau10(a, d);
      default:
        return "case 11";
    }
  }

  private static string cau1(Application a, Presentation d)
  {
    try
    {
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.SlideShowTransition.EntryEffect != PpEntryEffect.ppEffectFadeSmoothly)
          return "False";
      }
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string cau2(Application a, Presentation d)
  {
    try
    {
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.SlideShowTransition.Duration.ToString() != "2")
          return "False(Duration=2)";
      }
      if (a.ActivePresentation.Slides[(object) 1].SlideShowTransition.EntryEffect.ToString() == a.ActivePresentation.Slides[(object) 2].SlideShowTransition.EntryEffect.ToString())
        return "False(khong apply to all)";
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
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.SlideShowTransition.EntryEffect.ToString() != "ppEffectPushRight")
          return "False(from left)";
      }
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string cau4(Application a, Presentation d)
  {
    try
    {
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.SlideShowTransition.Duration.ToString() != "3")
          return "False(Duration=3)";
        if (slide.SlideShowTransition.SoundEffect.Name != "breeze.wav")
          return "False(breeze)";
      }
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string cau5(Application a, Presentation d)
  {
    try
    {
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.SlideShowTransition.EntryEffect.ToString() != "ppEffectWipeRight")
          return "False(Wipe from left)";
      }
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
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.SlideShowTransition.Duration.ToString() != "3")
          return "False(Duration=3)";
      }
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string cau7(Application a, Presentation d)
  {
    try
    {
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.SlideShowTransition.EntryEffect.ToString() != "ppEffectPushRight")
          return "False(form left)";
      }
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string cau8(Application a, Presentation d)
  {
    try
    {
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.SlideShowTransition.EntryEffect.ToString() != "ppEffectWedge")
          return "False(Wedge)";
      }
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string cau9(Application a, Presentation d)
  {
    try
    {
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.SlideShowTransition.Duration.ToString() != "2")
          return "False(Duration=2 for all slide)";
      }
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string cau10(Application a, Presentation d)
  {
    try
    {
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.SlideShowTransition.EntryEffect.ToString() != "ppEffectRotateRight")
          return "False(Rotate form left)";
      }
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }
}
