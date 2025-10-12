// Decompiled with JetBrains decompiler
// Type: MOS_WORD_LEARN.Illustrator
// Assembly: MOS_WORD_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;

namespace MOS_WORD_LEARN
{
  internal class Illustrator
  {
    public static string CheckCau(int causo, Application a, _Document d)
    {
      switch (causo)
      {
        case 1:
          return Illustrator.cau54(a, d);
        case 2:
          return Illustrator.cau42(a, d);
        case 3:
          return Illustrator.cau1(a, d);
        case 4:
          return Illustrator.cau4(a, d);
        case 5:
          return Illustrator.cau47(a, d);
        case 6:
          return Illustrator.cau46(a, d);
        case 7:
          return Illustrator.cau53(a, d);
        case 8:
          return Illustrator.cau41(a, d);
        case 9:
          return Illustrator.cau51(a, d);
        case 10:
          return Illustrator.cau49(a, d);
        case 11:
          return Illustrator.cau50(a, d);
        case 12:
          return Illustrator.cau43(a, d);
        case 13:
          return Illustrator.cau5(a, d);
        case 14:
          return Illustrator.cau2(a, d);
        case 15:
          return Illustrator.cau7(a, d);
        case 16:
          return Illustrator.cau52(a, d);
        case 17:
          return Illustrator.cau12(a, d);
        case 18:
          return Illustrator.cau9(a, d);
        case 19:
          return Illustrator.cau45(a, d);
        case 20:
          return Illustrator.cau44(a, d);
        case 21:
          return Illustrator.cau39(a, d);
        case 22:
          return Illustrator.cau15(a, d);
        case 23:
          return Illustrator.cau6(a, d);
        case 24:
          return Illustrator.cau24(a, d);
        case 25:
          return Illustrator.cau25(a, d);
        case 26:
          return Illustrator.cau26(a, d);
        case 27:
          return Illustrator.cau27(a, d);
        case 28:
          return Illustrator.cau28(a, d);
        case 29:
          return Illustrator.cau29(a, d);
        case 30:
          return Illustrator.cau30(a, d);
        case 31:
          return Illustrator.cau31(a, d);
        case 32:
          return Illustrator.cau32(a, d);
        case 33:
          return Illustrator.cau33(a, d);
        case 34:
          return Illustrator.cau34(a, d);
        case 35:
          return Illustrator.cau35(a, d);
        case 36:
          return Illustrator.cau36(a, d);
        case 37:
          return Illustrator.cau37(a, d);
        case 38:
          return Illustrator.cau38(a, d);
        case 39:
          return Illustrator.cau39(a, d);
        case 40:
          return Illustrator.cau40(a, d);
        case 41:
          return Illustrator.cau41(a, d);
        case 42:
          return Illustrator.cau42(a, d);
        case 43:
          return Illustrator.cau43(a, d);
        case 44:
          return Illustrator.cau44(a, d);
        case 45:
          return Illustrator.cau45(a, d);
        case 46:
          return Illustrator.cau46(a, d);
        case 47:
          return Illustrator.cau47(a, d);
        case 48:
          return Illustrator.cau48(a, d);
        case 49:
          return Illustrator.cau49(a, d);
        case 50:
          return Illustrator.cau50(a, d);
        case 51:
          return Illustrator.cau51(a, d);
        case 52:
          return Illustrator.cau52(a, d);
        case 53:
          return Illustrator.cau53(a, d);
        case 54:
          return Illustrator.cau54(a, d);
        case 55:
          return Illustrator.cau55(a, d);
        case 56:
          return Illustrator.cau56(a, d);
        default:
          return "default_Illustrator";
      }
    }

    private static string cau1(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Once again,"))
          ++Index;
        if (d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Layout.Name != "Vertical Bullet List")
                    return "False";
                if (d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Nodes[(object) 1].TextFrame2.TextRange.Text != "Frank Miller")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau2(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[6].Range.InlineShapes[1].Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectBackgroundRemoval)
                    return "False";
                if (double.Parse(d.Paragraphs[6].Range.InlineShapes[1].Fill.PictureEffects[1].EffectParameters[(object) 1].Value.ToString()) >= 0.1)
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau3(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Rectangle 4";
        if (!d.Shapes[ref Index].TextFrame.TextRange.Text.ToLower().Contains("good sounds"))
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau4(Application a, _Document d)
    {
      try
      {
        object count1 = (object) d.Shapes.Count;
        if (d.Shapes[ref count1].TextFrame.TextRange.Text.Contains("We Heard you..."))
                    return "False";
                object count2 = (object) d.Shapes.Count;
        if (d.Shapes[ref count2].TextFrame.TextRange.Font.TextColor.RGB.ToString() != "-16777216")
                    return "False";
                object count3 = (object) d.Shapes.Count;
        if (d.Shapes[ref count3].Left.ToString() != "-999995")
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau5(Application a, _Document d)
    {
      try
      {
        if ((double) d.Paragraphs[2].Range.InlineShapes[1].Line.Weight != 15.0)
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau6(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[3].Range.Font.TextShadow.Type != MsoShadowType.msoShadow22)
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau7(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[4].Range.InlineShapes[1].Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectPencilGrayscale)
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau8(Application a, _Document d)
    {
      try
      {
        object count1 = (object) d.Shapes.Count;
        if (!d.Shapes[ref count1].Name.Contains("Sun"))
                    return "False";
                object count2 = (object) d.Shapes.Count;
        if (d.Shapes[ref count2].Left.ToString() != "-999995")
                    return "False";
                object count3 = (object) d.Shapes.Count;
        if (d.Shapes[ref count3].Top.ToString() != "-999997")
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau9(Application a, _Document d)
    {
      try
      {
        if ((double) d.Paragraphs[3].Range.InlineShapes[1].SmartArt.Nodes[(object) 1].Shapes.ThreeD.BevelTopDepth != 4.0)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau10(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Registration Dates"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index - 1].Range.InlineShapes[1].SmartArt.Layout.Name != "Basic Chevron Process")
                    return "False";
                if (d.Paragraphs[Index - 1].Range.InlineShapes[1].SmartArt.Reverse == MsoTriState.msoTrue)
                    return "False";
                if (d.Paragraphs[Index - 1].Range.InlineShapes[1].SmartArt.Nodes[(object) 1].Shapes.TextFrame2.TextRange.Text != "Register Team")
                    return "False";
                if (d.Paragraphs[Index - 1].Range.InlineShapes[1].SmartArt.Nodes[(object) 2].Shapes.TextFrame2.TextRange.Text != "Managers Meeting")
                    return "False";
                if (d.Paragraphs[Index - 1].Range.InlineShapes[1].SmartArt.Nodes[(object) 3].Shapes.TextFrame2.TextRange.Text != "Have Fun!")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau11(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Once again, we would like to offer special thanks to"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
                    return "False";
                string name = d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Layout.Name;
        if (name != "Vertical Bullet List")
                    return "False";
                if (d.Paragraphs[Index + 1].Range.InlineShapes[1].SmartArt.Nodes[(object) 1].TextFrame2.TextRange.Text != "Frank Miller")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau12(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Picture 3";
        string str = d.Shapes[ref Index].ThreeD.BevelTopInset.ToString();
        if (str != "12")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau13(Application a, _Document d)
    {
      try
      {
        int Index = 1;
        while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Fourth Coffee"))
          ++Index;
        if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index + 1].Range.InlineShapes.Count != 1)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau14(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Rectangle 6";
        if (d.Shapes[ref Index].TextFrame.TextRange.Paragraphs[4].Range.InlineShapes.Count != 1)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau15(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Diagram 1";
        if (d.Shapes[ref Index].SmartArt.Nodes[(object) 2].TextFrame2.TextRange.Text != "organically grown")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau16(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[6].Range.InlineShapes[1].Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectBackgroundRemoval)
                    return "False";
                if (double.Parse(d.Paragraphs[6].Range.InlineShapes[1].Fill.PictureEffects[1].EffectParameters[(object) 1].Value.ToString()) >= 0.1)
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau17(Application a, _Document d)
    {
      try
      {
        object Index1 = (object) "Rectangle 4";
        if (!d.Shapes[ref Index1].TextFrame.TextRange.Text.ToLower().Contains("good sounds"))
                    return "False";
                object Index2 = (object) "Rectangle 4";
        if (d.Shapes[ref Index2].Left.ToString() != "-999995")
                    return "False";
                object Index3 = (object) "Rectangle 4";
        if (d.Shapes[ref Index3].Top.ToString() != "-999997")
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau18(Application a, _Document d)
    {
      try
      {
        object Index1 = (object) "Text Box 3";
        if (!d.Shapes[ref Index1].TextFrame.TextRange.Text.ToLower().Contains("we heard you"))
                    return "False";
                object Index2 = (object) "Text Box 3";
        if (d.Shapes[ref Index2].Left.ToString() != "-999995")
                    return "False";
                object Index3 = (object) "Text Box 3";
        if (d.Shapes[ref Index3].TextFrame.TextRange.Font.Fill.ForeColor.RGB.ToString() != "6051667")
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau19(Application a, _Document d)
    {
      try
      {
        if ((double) d.Paragraphs[2].Range.InlineShapes[1].Line.Weight != 15.0)
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau20(Application a, _Document d)
    {
      try
      {
        object Index1 = (object) "Text Box 2";
        if (!d.Shapes[ref Index1].TextFrame.TextRange.Text.ToUpper().Contains("NEW EMPLOYEE TRAINING"))
                    return "False";
                object Index2 = (object) "Text Box 2";
        if (d.Shapes[ref Index2].TextFrame.TextRange.Font.Fill.ForeColor.RGB.ToString() != "6051667")
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau21(Application a, _Document d)
    {
      try
      {
        object Index1 = (object) "Horizontal Scroll 4";
        if (!d.Shapes[ref Index1].TextFrame.TextRange.Text.ToLower().Contains("class starts at 9:00!"))
                    return "False";
                object Index2 = (object) "Horizontal Scroll 4";
        if (d.Shapes[ref Index2].Left.ToString() != "-999995")
                    return "False";
                object Index3 = (object) "Horizontal Scroll 4";
        if (d.Shapes[ref Index3].Top.ToString() != "-999997")
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau22(Application a, _Document d)
    {
      try
      {
        if ((double) d.Paragraphs[3].Range.InlineShapes[1].SmartArt.Nodes[(object) 1].Shapes.ThreeD.BevelTopDepth != 4.0)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau23(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[4].Range.InlineShapes[1].Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectPencilGrayscale)
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau24(Application a, _Document d)
    {
      try
      {
        object count1 = (object) d.Shapes.Count;
        if (!d.Shapes[ref count1].Name.Contains("Sun"))
                    return "False";
                object count2 = (object) d.Shapes.Count;
        if (d.Shapes[ref count2].Left.ToString() != "-999995")
                    return "False";
                object count3 = (object) d.Shapes.Count;
        if (d.Shapes[ref count3].Top.ToString() != "-999997")
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau25(Application a, _Document d)
    {
      try
      {
        if (d.Shapes.Count != 3)
                    return "False";
                object count1 = (object) d.Shapes.Count;
        if (!d.Shapes[ref count1].TextFrame.TextRange.Text.Contains("Fourth Coffee"))
                    return "False";
                object count2 = (object) d.Shapes.Count;
        if (d.Shapes[ref count2].Left.ToString() != "-999995")
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau26(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[2].Range.InlineShapes.Count != 1)
                    return "False";
                if (d.Paragraphs[2].Range.InlineShapes[1].Title != "")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau27(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Diagram 1";
        if (d.Shapes[ref Index].SmartArt.Nodes[(object) 2].TextFrame2.TextRange.Text != "The best beans")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau28(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[5].Range.InlineShapes[1].Borders[WdBorderType.wdBorderTop].Color.ToString() != "-654245889")
                    return "False";
                if (d.Paragraphs[5].Range.InlineShapes[1].Borders[WdBorderType.wdBorderTop].LineWidth.ToString() != "wdLineWidth300pt")
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau29(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[5].Range.InlineShapes[1].Fill.PictureEffects[1].Type.ToString() != "msoEffectPhotocopy")
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau30(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Picture 3";
        if ((double) d.Shapes[ref Index].ThreeD.BevelTopDepth != 4.0)
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau31(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[5].Range.InlineShapes[1].SoftEdge.Radius.ToString() != "8.858268")
                    return "False";
            }
      catch (Exception ex)
            {
                return "False";
            }
      return "True";
    }

    private static string cau32(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Rectangle 6";
        if (d.Shapes[ref Index].TextFrame.TextRange.Paragraphs[4].Range.InlineShapes.Count != 1)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau33(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[4].Range.InlineShapes.Count != 1)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau34(Application a, _Document d)
    {
      try
      {
        object Index1 = (object) "Picture 3";
        if ((double) d.Shapes[ref Index1].Left != -999995.0)
                    return "False";
                object Index2 = (object) "Picture 3";
        if ((double) d.Shapes[ref Index2].Top != -999995.0)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau35(Application a, _Document d)
    {
      try
      {
        object count = (object) d.Shapes.Count;
        string name = d.Shapes[ref count].Name;
        if (!name.Contains("5-Point Star"))
                    return "False";
                object Index1 = (object) name;
        if ((double) d.Shapes[ref Index1].Left != -999995.0)
                    return "False";
                object Index2 = (object) name;
        if ((double) d.Shapes[ref Index2].Top != -999997.0)
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau36(Application a, _Document d)
    {
      try
      {
        object count = (object) d.Shapes.Count;
        string name = d.Shapes[ref count].Name;
        if (!name.Contains("12-Point Star"))
          return "False";
        object Index1 = (object) name;
        if ((double) d.Shapes[ref Index1].Left != -999995.0)
          return "False";
        object Index2 = (object) name;
        if ((double) d.Shapes[ref Index2].Top != -999997.0)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau37(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[2].Range.InlineShapes.Count != 1)
          return "False";
        if (d.Paragraphs[2].Range.InlineShapes[1].SmartArt == null)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau38(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[5].Range.InlineShapes.Count != 1)
          return "False";
        if (d.Paragraphs[5].Range.InlineShapes[1].SmartArt == null)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau39(Application a, _Document d)
    {
      try
      {
        if (d.Paragraphs[7].Range.InlineShapes[1].SmartArt.Color.Name != "Transparent Gradient Range - Accent 1")
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau40(Application a, _Document d)
    {
      try
      {
        object Index1 = (object) "Text Box 2";
        if (!d.Shapes[ref Index1].TextFrame.TextRange.Text.Contains("ABC"))
          return "False";
        object Index2 = (object) "Text Box 2";
        if ((double) d.Shapes[ref Index2].Left != -999995.0)
          return "False";
        object Index3 = (object) "Text Box 2";
        if ((double) d.Shapes[ref Index3].Top != -999997.0)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau41(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Picture 3";
        if (d.Shapes[ref Index].WrapFormat.Type == WdWrapType.wdWrapSquare &&
            d.Shapes[ref Index].WrapFormat.Side == WdWrapSideType.wdWrapLeft)
          return "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "False";
    }

    private static string cau42(Application a, _Document d)
    {
      try
      {
        if (d.Shapes.Count != 2)
          return "False";
        object Index1 = (object) 2;
        if (!d.Shapes[ref Index1].Name.Contains("Horizontal Scroll"))
          return "False";
        object Index2 = (object) 2;
        if (d.Shapes[ref Index2].TextFrame.TextRange.Text.Trim() != "Remember your calculator!")
          return "Fales";
        object Index3 = (object) 2;
        if (d.Shapes[ref Index3].WrapFormat.Type != WdWrapType.wdWrapSquare)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau43(Application a, _Document d)
    {
      try
      {
        if (d.Shapes.Count != 1)
          return "False";
        object Index = (object) "Picture 1";
        if (d.Shapes[ref Index].Line.ForeColor.RGB != 11957550)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau44(Application a, _Document d)
    {
      try
      {
        if (d.InlineShapes[1].SmartArt.Reverse == MsoTriState.msoTrue)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau45(Application a, _Document d)
    {
      try
      {
        if (d.InlineShapes.Count != 2)
          return "False";
        if ((double) d.InlineShapes[1].SmartArt.Nodes[(object) 1].Shapes.ThreeD.BevelTopDepth != 4.0)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau46(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Picture 11";
        if (d.Shapes[ref Index].WrapFormat.Type != WdWrapType.wdWrapSquare)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau47(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Rectangle 4";
        if (d.Shapes[ref Index].TextFrame.TextRange.Text.Trim() != "ANYTIME ACCOUNT ACCESS")
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau48(Application a, _Document d)
    {
      try
      {
        if (d.Tables.Count != 1)
          return "Fasle";
        if (d.Tables[1].Rows[1].Cells.Count != 1)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau49(Application a, _Document d)
    {
      try
      {
        int num = 0;
        for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Description"))
          {
            num = Index;
            break;
          }
        }
        if (num == 0)
          return "False";
        if (d.Paragraphs[num + 3].Range.InlineShapes.Count == 0)
          return "Fasle";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau50(Application a, _Document d)
    {
      try
      {
        if (d.InlineShapes.Count != 1)
          return "False";
        if (d.InlineShapes[1].AlternativeText == "")
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau51(Application a, _Document d)
    {
      try
      {
        int num = 0;
        for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
        {
          if (d.Paragraphs[Index].Range.Text.Contains("Favorite dinosaurs"))
          {
            num = Index;
            break;
          }
        }
        if (num == 0)
          return "False";
        if (d.Paragraphs[num + 3].Range.InlineShapes.Count == 0)
          return "Fasle";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau52(Application a, _Document d)
    {
      try
      {
        object Index1 = (object) "Picture 6";
        if (d.Shapes[ref Index1].Fill.PictureEffects.Count != 1)
          return "False";
        object Index2 = (object) "Picture 6";
        if (d.Shapes[ref Index2].Fill.PictureEffects[1].Type != MsoPictureEffectType.msoEffectPencilSketch)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau53(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Picture 4";
        if (d.Shapes[ref Index].WrapFormat.Type != WdWrapType.wdWrapSquare)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau54(Application a, _Document d)
    {
      int num = 0;
      for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
      {
        if (d.Paragraphs[Index].Range.Text.Contains("Where to find us"))
        {
          num = Index;
          break;
        }
      }
      if (num == 0)
        return "False";
      return d.Paragraphs[num + 1].Range.InlineShapes.Count != 1 ? "False" : "True";
    }

    private static string cau55(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Picture 4";
        if (d.Shapes[ref Index].WrapFormat.Type != WdWrapType.wdWrapSquare)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau56(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Picture 4";
        if (d.Shapes[ref Index].WrapFormat.Type != WdWrapType.wdWrapSquare)
          return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }
  }
}
