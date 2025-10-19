// Decompiled with JetBrains decompiler
// Type: MOS_WORD_LEARN.Links
// Assembly: MOS_WORD_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.Office.Interop.Word;
using System;

namespace MOS_WORD_LEARN
{
  internal class Links
  {
    public static string CheckCau(int causo, Application a, _Document d)
    {
      switch (causo)
      {
        case 1:
          return Links.cau8(a, d);
        case 2:
          return Links.cau7(a, d);
        case 3:
          return Links.cau3(a, d);
        case 4:
          return Links.cau4(a, d);
        case 5:
          return Links.cau5(a, d);
        case 6:
          return Links.cau6(a, d);
        case 7:
          return Links.cau7(a, d);
        case 8:
          return Links.cau8(a, d);
        case 9:
          return Links.cau9(a, d);
        case 10:
          return Links.cau10(a, d);
        case 11:
          return Links.cau11(a, d);
        case 12:
          return Links.cau12(a, d);
        case 13:
          return Links.cau13(a, d);
        case 14:
          return Links.cau14(a, d);
        case 15:
          return Links.cau15(a, d);
        case 16:
          return Links.cau16(a, d);
        case 17:
          return Links.cau17(a, d);
        case 18:
          return Links.cau18(a, d);
        case 19:
          return Links.cau19(a, d);
        case 20:
          return Links.cau20(a, d);
        default:
          return "default links";
      }
    }

    private static string cau1(Application a, _Document d)
    {
      try
      {
        if (d.Hyperlinks.Count == 0)
          return "False (don't have hyperlink)";
        if (d.Hyperlinks.Count > 1)
          return "False (number of hyperlink)";
        object Index1 = (object) 1;
        if (!d.Hyperlinks[ref Index1].Address.Contains("http://www.tailspintoys.com"))
          return "False (http://www.tailspintoys.com)";
        object Index2 = (object) 1;
        if (d.Hyperlinks[ref Index2].TextToDisplay != "tailspintoys.com")
          return "False (tailspintoys.com)";
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
        if (d.Bookmarks.Count != 1)
          return "False (Number of bookmark)";
        object Index1 = (object) 1;
        if (d.Bookmarks[ref Index1].Name != "Internal")
        {
          object Index2 = (object) 1;
          return "False (name=" + d.Bookmarks[ref Index2].Name + ")";
        }
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
        if (d.Hyperlinks.Count == 0)
                    return "False";
                if (d.Hyperlinks.Count > 1)
                    return "False";
                object Index1 = (object) 1;
        if (!d.Hyperlinks[ref Index1].Address.Contains("http://www.tailspintoys.com"))
                    return "False";
                object Index2 = (object) 1;
        if (d.Hyperlinks[ref Index2].TextToDisplay != "tailspintoys.com")
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
        if (d.Bookmarks.Count != 1)
                    return "False";
                object Index1 = (object) 1;
        if (d.Bookmarks[ref Index1].Name != "BeEntertaining")
        {
          object Index2 = (object) 1;
                    return "False";
                }
        object Index3 = (object) 1;
        if (!d.Bookmarks[ref Index3].Range.Text.Contains("Be Entertaining"))
        {
          object Index4 = (object) 1;
                    return "False";
                }
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
        if (d.Bookmarks.Count != 1)
                    return "False";
                object Index1 = (object) 1;
        if (d.Bookmarks[ref Index1].Name != "ABC")
        {
          object Index2 = (object) 1;
                    return "False";
                }
        object Index3 = (object) 1;
        if (d.Bookmarks[ref Index3].Start != 2605)
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
        if (d.Hyperlinks.Count == 0)
                    return "False";
                if (d.Hyperlinks.Count > 1)
                    return "False";
                object Index1 = (object) 1;
        if (!d.Hyperlinks[ref Index1].Address.Contains("Mos360.vn"))
                    return "False";
                object Index2 = (object) 1;
        if (d.Hyperlinks[ref Index2].TextToDisplay != "Cultivated")
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
        if (d.Hyperlinks.Count == 0)
                    return "False";
                if (d.Hyperlinks.Count > 1)
                    return "False";
                object Index1 = (object) 1;
        if (!d.Hyperlinks[ref Index1].SubAddress.Contains("Introduction"))
                    return "False";
                object Index2 = (object) 1;
        if (d.Hyperlinks[ref Index2].TextToDisplay != "Home")
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
        if (d.Bookmarks.Count == 0)
                    return "False";
                if (d.Bookmarks.Count > 1)
                    return "False";
                object Index = (object) 1;
        if (!d.Bookmarks[ref Index].Name.Contains("Cheating"))
                    return "False";
            }
      catch (Exception ex)
      {
        return "False";
      }
      return "True";
    }

    private static string cau9(Application a, _Document d) => "True";

    private static string cau10(Application a, _Document d) => "True";

    private static string cau11(Application a, _Document d) => "True";

    private static string cau12(Application a, _Document d) => "True";

    private static string cau13(Application a, _Document d) => "True";

    private static string cau14(Application a, _Document d) => "True";

    private static string cau15(Application a, _Document d) => "True";

    private static string cau16(Application a, _Document d) => "True";

    private static string cau17(Application a, _Document d) => "True";

    private static string cau18(Application a, _Document d) => "True";

    private static string cau19(Application a, _Document d) => "True";

    private static string cau20(Application a, _Document d) => "True";

    private static string cau4_2010(Application a, _Document d)
    {
      try
      {
        if (d.Hyperlinks.Count == 0)
          return "False";
        foreach (Hyperlink hyperlink in d.Hyperlinks)
        {
          if (hyperlink.Address.Contains("mos.edu.vn") && hyperlink.TextToDisplay.Contains("click here"))
            return "True";
        }
        return "False";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }

    private static string cau9_2010(Application a, _Document d)
    {
      try
      {
        object Index = (object) "Picture 5";
        return d.Shapes[ref Index].Hyperlink.Name != "_Water" ? "False" : "True";
      }
      catch (Exception ex)
      {
        return "False";
      }
    }
  }
}
