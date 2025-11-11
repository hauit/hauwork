// Decompiled with JetBrains decompiler
// Type: WindowsFormsApplication1.S13
// Assembly: WindowsFormsApplication1, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2D6280CC-B1DC-46AA-AAB0-60B9FE9957DE
// Assembly location: D:\hau_Work\Git_HauWork\hauwork\PPT_19\HOC_PPT_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.IO;
using System.Reflection;

#nullable disable
namespace WindowsFormsApplication1;

internal class S13
{
  public static string CheckCau(int causo, Application a, Presentation d)
  {
    switch (causo)
    {
      case 1:
        return S13.Cau11(a, d);
      case 2:
        return S13.Cau12(a, d);
      case 3:
        return S13.Cau13(a, d);
      case 4:
        return S13.Cau14(a, d);
      case 5:
        return S13.Cau15(a, d);
      case 6:
        return S13.Cau16(a, d);
      case 7:
        return S13.Cau17(a, d);
      case 8:
        return S13.Cau8(a, d);
      case 9:
        return S13.Cau9(a, d);
      case 10:
        return S13.Cau10(a, d);
      case 11:
        return S13.cau11(a, d);
      case 12:
        return S13.cau12(a, d);
      case 13:
        return S13.cau13(a, d);
      default:
        return "case out index";
    }
  }

  private static string Cau1(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence.Count != 1)
        return "False (khong them xoa hieu ung)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[1].DisplayName != "Luxembourg")
        return "False (Luxembourg)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectPathCircle")
        return "False (circle)";
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
      if (a.ActivePresentation.SnapToGrid != MsoTriState.msoTrue)
        return "False(snap object)";
      if (a.DisplayGridLines != MsoTriState.msoTrue)
        return "False(griline)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau3(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 8].Shapes[(object) "Heart 3"].Width.ToString() != "569.6")
        return "False(Width)";
      if (a.ActivePresentation.Slides[(object) 8].Shapes[(object) "Heart 3"].Height.ToString() != "505.6")
        return "False(Height)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau4(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 4].SlideShowTransition.SoundEffect.Name != "applause.wav")
        return "False(Transition sound applause)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau5(Application a, Presentation d)
  {
    try
    {
      if (!a.ActivePresentation.Slides[(object) 9].Shapes[(object) "Title 1"].TextFrame.TextRange.Text.Contains("Extra"))
        return "False(sai outline)";
    }
    catch (Exception ex)
    {
      return "False(add slide tu outline file)";
    }
    return "True";
  }

  private static string Cau6(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.HandoutMaster.HeadersFooters.Header.Text != "Draft")
        return "False(Draft)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau7(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.PrintOptions.NumberOfCopies.ToString() != "5")
        return "False(5 copys)";
      if (a.ActivePresentation.PrintOptions.OutputType != PpPrintOutputType.ppPrintOutputNotesPages)
        return "False(note page)";
      if (a.ActivePresentation.PrintOptions.Collate == MsoTriState.msoTrue)
        return "False(Collate)";
      if (a.ActivePresentation.PageSetup.NotesOrientation == MsoOrientation.msoOrientationVertical)
        return "False(Orientation)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau8(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 3].Shapes[(object) "Content Placeholder 4"].TextFrame2.Column.Number.ToString() != "2")
        return "False(2 cot)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau9(Application a, Presentation d)
  {
    try
    {
      // ISSUE: reference to a compiler-generated method
      if (a.ActivePresentation.SectionProperties.Name(4) != "Sample")
        return "False(Sample)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau10(Application a, Presentation d)
  {
    try
    {
      foreach (Slide slide in a.ActivePresentation.Slides)
      {
        if (slide.SlideShowTransition.Duration.ToString() != "2")
          return $"False({(object) slide.SlideNumber})";
      }
    }
    catch (Exception ex)
    {
      return "False(add slide tu outline file)";
    }
    return "True";
  }

  private static string Cau11(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 2].SlideShowTransition.EntryEffect != PpEntryEffect.ppEffectCurtains)
        return "False(Curtains on slide 2)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau12(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 8].Shapes[(object) 1].TextFrame.TextRange.Text != "Title Layout")
        return "False(slide 8)";
      if (a.ActivePresentation.Slides[(object) 12].Shapes[(object) 1].TextFrame.TextRange.Text != "Two Content Layout with SmartArt")
        return "False(slide 12)";
    }
    catch (Exception ex)
    {
      return "False(reused slide)";
    }
    return "True";
  }

  private static string Cau13(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.RemovePersonalInformation != MsoTriState.msoTrue)
        return "False";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau14(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.TemplateName != "Ion Boardroom")
        return "False(Ion Boardroom)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau15(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 6].Shapes.Count != 3)
        return "False(group)";
      if (!a.ActivePresentation.Slides[(object) 6].Shapes[(object) 3].Name.Contains("Group"))
        return "False(Group)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau16(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 3].Shapes[(object) "Diagram 2"].SmartArt.Color.Name != "Colorful Range - Accent Colors 4 to 5")
        return "False(color)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau17(Application a, Presentation d)
  {
    string str = "True";
    try
    {
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence.Count != 4)
        return "False (ap dung cho noi dung)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[1].DisplayName != "Restaurant ")
        return "False (Restaurant)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[1].EffectType.ToString() != "msoAnimEffectWipe")
        return "False (1 Wipe)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[1].EffectParameters.Direction != MsoAnimDirection.msoAnimDirectionUp)
        return "False (1 from top)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[1].Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
        return "False (onclick)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[1].Timing.TriggerDelayTime.ToString() != "0")
        return "False (with)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[2].EffectParameters.Direction != MsoAnimDirection.msoAnimDirectionUp)
        return "False (2 from top)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[2].Timing.TriggerDelayTime.ToString() != "1")
        return "False (2 delpay)";
      str = a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[2].Timing.TriggerType.ToString();
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[2].Timing.TriggerType != MsoAnimTriggerType.msoAnimTriggerAfterPrevious)
        return "False (2 After)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[2].EffectType.ToString() != "msoAnimEffectWipe")
        return "False (2 Wipe)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[3].EffectParameters.Direction != MsoAnimDirection.msoAnimDirectionUp)
        return "False (3 from top)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[3].Timing.TriggerDelayTime.ToString() != "1")
        return "False (3 delpay)";
      str = a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[3].Timing.TriggerType.ToString();
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[3].Timing.TriggerType != MsoAnimTriggerType.msoAnimTriggerAfterPrevious)
        return "False (3 After)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[3].EffectType.ToString() != "msoAnimEffectWipe")
        return "False (3 Wipe)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[3].EffectParameters.Direction != MsoAnimDirection.msoAnimDirectionUp)
        return "False (4 from top)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[3].Timing.TriggerDelayTime.ToString() != "1")
        return "False (4 delpay)";
      str = a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[3].Timing.TriggerType.ToString();
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[3].Timing.TriggerType != MsoAnimTriggerType.msoAnimTriggerAfterPrevious)
        return "False (4 After)";
      if (a.ActivePresentation.Slides[(object) 5].TimeLine.MainSequence[3].EffectType.ToString() != "msoAnimEffectWipe")
        return "False (4 Wipe)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau18(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 2].Shapes[(object) "Wildlife"].MediaFormat.StartPoint.ToString() != "1000")
        return "False()";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau19(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 4].Layout.ToString() != "ppLayoutTwoObjects")
        return "False(Layout)";
    }
    catch (Exception ex)
    {
      return "False(reused slide)";
    }
    return "True";
  }

  private static string Cau20(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 6].Shapes[(object) "Heart 1"].Left.ToString() != "374.7")
        return "False()";
      if (a.ActivePresentation.Slides[(object) 6].Shapes[(object) "Smiley Face 2"].Left.ToString() != "435")
        return "False()";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau21(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.HandoutMaster.Shapes.Count.ToString() != "3")
        return "False(bo date placehonder trong HandoutMaster)";
    }
    catch (Exception ex)
    {
      return "False(loi khong xac dinh)";
    }
    return "True";
  }

  private static string Cau22(Application a, Presentation d)
  {
    try
    {
      if (a.ActivePresentation.Slides[(object) 8].Shapes[(object) "TextBox 4"].TextFrame2.TextRange.Font.Spacing.ToString() != "6")
        return "False(6pt)";
      if (((object) a.ActivePresentation.Slides[(object) 8].Shapes[(object) "TextBox 4"].TextFrame2.TextRange.Font.Shadow.Visible).ToString() != "msoTrue")
        return "False(shadow)";
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
      object documentProperties = a.ActivePresentation.BuiltInDocumentProperties;
      object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (Binder) null, documentProperties, new object[1]
      {
        (object) "Title"
      });
      return target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (Binder) null, target, new object[0]).ToString() != "Golden Badge Solutions" ? "False (Golden Badge Solutions)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Something wrong)";
    }
  }

  private static string cau7(Application a, Presentation d)
  {
    try
    {
      string str = "True";
      if (a.ActivePresentation.PrintOptions.OutputType != PpPrintOutputType.ppPrintOutputThreeSlideHandouts)
        return "False (Handout three Slide)";
      return a.ActivePresentation.PrintOptions.FrameSlides != MsoTriState.msoTrue ? "False (FrameSlides)" : str;
    }
    catch (Exception ex)
    {
      return "False (Somthing Wrong)";
    }
  }

  private static string cau8(Application a, Presentation d)
  {
    try
    {
      return new FileInfo(a.ActivePresentation.FullName).Length < 1800000L ? "False (saved)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Somthing Wrong)";
    }
  }

  private static string cau9(Application a, Presentation d) => "True";

  private static string cau10(Application a, Presentation d)
  {
    try
    {
      object documentProperties = a.ActivePresentation.BuiltInDocumentProperties;
      object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (Binder) null, documentProperties, new object[1]
      {
        (object) "Title"
      });
      if (target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (Binder) null, target, new object[0]).ToString() == "")
        return "False (not remove properties)";
    }
    catch (Exception ex)
    {
      return "False ()";
    }
    try
    {
      string str = "True";
      if (a.ActivePresentation.Slides[(object) 2].Shapes.Count == 4)
        return "False (On Slide Content)";
      return a.ActivePresentation.Slides[(object) 7].Shapes.Count == 4 ? "False (Off Slide Content)" : str;
    }
    catch (Exception ex)
    {
      return "Something Wrong";
    }
  }

  private static string cau11(Application a, Presentation d)
  {
    try
    {
      object documentProperties = a.ActivePresentation.BuiltInDocumentProperties;
      object target = documentProperties.GetType().InvokeMember("Item", BindingFlags.GetProperty, (Binder) null, documentProperties, new object[1]
      {
        (object) "Subject"
      });
      return target.GetType().InvokeMember("Value", BindingFlags.GetProperty, (Binder) null, target, new object[0]).ToString() != "Corporate" ? "False (Corporate)" : "True";
    }
    catch (Exception ex)
    {
      return "False ()";
    }
  }

  private static string cau12(Application a, Presentation d)
  {
    try
    {
      string str = "True";
      if (a.ActivePresentation.PrintOptions.OutputType != PpPrintOutputType.ppPrintOutputTwoSlideHandouts)
        return "False (Handout Two Slide)";
      return a.ActivePresentation.PrintOptions.FitToPage != MsoTriState.msoTrue ? "False (FitToPage)" : str;
    }
    catch (Exception ex)
    {
      return "False (Somthing Wrong)";
    }
  }

  private static string cau13(Application a, Presentation d)
  {
    try
    {
      return new FileInfo(a.ActivePresentation.FullName).Length < 1800000L ? "False (saved)" : "True";
    }
    catch (Exception ex)
    {
      return "False (Somthing Wrong)";
    }
  }
}
