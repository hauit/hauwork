// Decompiled with JetBrains decompiler
// Type: MOS_WORD_LEARN.DocText
// Assembly: MOS_WORD_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.Office.Interop.Word;
using System;

namespace MOS_WORD_LEARN
{
    internal class DocText
    {
        public static string CheckCau(int causo, Application a, _Document d)
        {
            switch (causo)
            {
                case 1:
                    return DocText.cau29(a, d);
                case 2:
                    return DocText.cau19(a, d);
                case 3:
                    return DocText.cau38(a, d);
                case 4:
                    return DocText.cau20(a, d);
                case 5:
                    return DocText.cau30(a, d);
                case 6:
                    return DocText.cau5(a, d);
                case 7:
                    return DocText.cau31(a, d);
                case 8:
                    return DocText.cau32(a, d);
                case 9:
                    return DocText.cau35(a, d);
                case 10:
                    return DocText.cau22(a, d);
                case 11:
                    return DocText.cau12(a, d);
                case 12:
                    return DocText.cau33(a, d);
                case 13:
                    return DocText.cau36(a, d);
                case 14:
                    return DocText.cau28(a, d);
                case 15:
                    return DocText.cau1(a, d);
                case 16:
                    return DocText.cau34(a, d);
                case 17:
                    return DocText.cau47Mask(a, d);
                case 18:
                    return DocText.cau18(a, d);
                case 19:
                    return DocText.cau7(a, d);
                case 20:
                    return DocText.cau20(a, d);
                case 21:
                    return DocText.cau21(a, d);
                case 22:
                    return DocText.cau22(a, d);
                case 23:
                    return DocText.cau23(a, d);
                case 24:
                    return DocText.cau24(a, d);
                case 25:
                    return DocText.cau25(a, d);
                case 26:
                    return DocText.cau26(a, d);
                case 27:
                    return DocText.cau27(a, d);
                case 28:
                    return DocText.cau28(a, d);
                case 29:
                    return DocText.cau29(a, d);
                case 30:
                    return DocText.cau30(a, d);
                case 31:
                    return DocText.cau31(a, d);
                case 32:
                    return DocText.cau32(a, d);
                case 33:
                    return DocText.cau33(a, d);
                case 34:
                    return DocText.cau34(a, d);
                case 35:
                    return DocText.cau35(a, d);
                case 36:
                    return DocText.cau36(a, d);
                case 37:
                    return DocText.cau37(a, d);
                case 38:
                    return DocText.cau38(a, d);
                default:
                    return "Default DocText";
            }
        }

        private static string cau1(Application a, _Document d)
        {
            try
            {
                if (d.Tables.Count != 9)
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
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("class"))
                        return "False(class)";
                }
                int Index1 = 1;
                while (Index1 < d.Paragraphs.Count && !d.Paragraphs[Index1].Range.Text.Contains("seminar"))
                    ++Index1;
                if (Index1 >= d.Paragraphs.Count)
                    return "False (seminar)";
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }

        private static string cau47Mask(Application a, _Document d)
        {
            try
            {
                Range FindRange(string t)
                {
                    Range r = d.Content;
                    if (r.Find.Execute(t)) return r.Duplicate;
                    return null;
                }

                Range r1 = FindRange("Contest");
                Range r2 = FindRange("The Dirty Details");
                if (r1 == null || r2 == null)
                    return "False";
                var c = r1.Font;
                var b = r2.Font;
                if ( r1.Font.Name != r2.Font.Name ||
                       r1.Font.Size != r2.Font.Size ||
                       r1.Font.Bold != r2.Font.Bold ||
                       r1.Font.Italic != r2.Font.Italic ||
                       r1.Font.Underline != r2.Font.Underline ||
                       r1.Font.Color != r2.Font.Color ||
                       r1.HighlightColorIndex != r2.HighlightColorIndex)
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
                if (!d.Paragraphs[5].Range.Text.Contains("Blue Yonder Airlines is sponsoring a contest that is open to all Graphic design Institution students."))
                    return "False(copy and pase end of paragraph)";
                if (d.Paragraphs[5].Range.Font.Bold != 0)
                    return "False(Paste value)";
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
                if (d.Tables.Count != 1)
                    return "False(table.count:" + (object)d.Tables.Count + ")";
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
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Learning WareWolf™"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
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
                if (d.Tables.Count != 2)
                    return "False(table.count:" + (object)d.Tables.Count + ")";
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
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Rehearse and Video Your Presentation"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (!d.Paragraphs[Index + 3].Range.Text.Contains("Summarize Main Points"))
                    return "False";
                if (!d.Paragraphs[Index + 5].Range.Text.Contains("You can easily change the formatting"))
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
                if (d.Tables.Count != 9)
                    return "False(table.count:" + (object)d.Tables.Count + ")";
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
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("class"))
                        return "False(class)";
                }
                int Index1 = 1;
                while (Index1 < d.Paragraphs.Count && !d.Paragraphs[Index1].Range.Text.Contains("seminar"))
                    ++Index1;
                if (Index1 >= d.Paragraphs.Count)
                    return "False (seminar)";
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
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("Friday"))
                        return "False(Friday)";
                }
                int Index1 = 1;
                while (Index1 < d.Paragraphs.Count && !d.Paragraphs[Index1].Range.Text.Contains("Tuesday"))
                    ++Index1;
                if (Index1 >= d.Paragraphs.Count)
                    return "False (Tuesday)";
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
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("We Made WareWoft™"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False(We Made WareWoft™)";
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
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("Keep it Simple"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (!d.Paragraphs[Index + 1].Range.Text.Contains("you specify directly."))
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
                if (d.Tables.Count != 1)
                    return "False(table.count:" + (object)d.Tables.Count + ")";
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
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains("As we promised, you can use these galleries and Friends Book Club®"))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False (insert ®)";
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
                if (!d.Paragraphs[4].Range.Text.Contains("Fax:ASA-45"))
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
                if (d.Paragraphs[4].Range.Text.Contains("Game"))
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
                if (!d.Paragraphs[5].Range.Text.Contains("Fourth Coffee"))
                    return "False(paste Fourth Coffee vao cuoi doan 4)";
                if (d.Shapes.Count != 4)
                    return "False(paste value)";
                if ((double)d.Paragraphs[5].Range.Font.Size != 11.0)
                    return "False(paste value)";
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
                if (!d.Paragraphs[6].Range.Text.Contains("Fourth Coffee"))
                    return "False";
                if (d.Paragraphs[6].Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphLeft)
                    return "False";
                if (d.Paragraphs[6].Range.Font.Bold != 9999999)
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
                if ((double)d.Paragraphs[1].Range.Font.Size != 11.0)
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
                if (d.Paragraphs[1].Range.Font.Name != "Algerian")
                    return "False";
                if (d.Paragraphs[1].Range.Font.Underline != WdUnderline.wdUnderlineThick)
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
                if (d.Bookmarks.Count != 1)
                    return "False";
                object Index1 = (object)1;
                if (d.Bookmarks[ref Index1].Range.Text != null)
                    return "False";
                object Index2 = (object)1;
                if (d.Bookmarks[ref Index2].StoryType != WdStoryType.wdTextFrameStory)
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
                if (d.Bookmarks.Count != 0)
                    return "False";
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("NOTE TO REVIEW COMMITTEE"))
                        return "False";
                }
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
                string str = "Special Thanks";
                int Index = 1;
                while (Index < d.Paragraphs.Count && !d.Paragraphs[Index].Range.Text.Contains(str))
                    ++Index;
                if (Index >= d.Paragraphs.Count)
                    return "False";
                if (d.Paragraphs[Index + 1].Range.Text != "\r")
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
                object Index = (object)"Rectangle 5";
                if (d.Shapes[ref Index].TextFrame.TextRange.Text != "\r")
                    return "False";
                if (d.Paragraphs[5].Range.ParagraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphLeft)
                    return "False";
                if (!d.Paragraphs[5].Range.Text.Contains("ABOUT OUR COFFEE"))
                    return "False";
                if (d.Paragraphs.Count != 8)
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
                if (!d.Paragraphs[6].Range.Text.Contains("Wilderness Summary©"))
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
                if (!d.Paragraphs[1].Range.Text.Contains("Barstow™ College"))
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
                if (!d.Paragraphs[1].Range.Text.Contains("Barstow College®"))
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
                if (d.Paragraphs[4].Range.Text.Contains("in the embed code"))
                    return "False";
                if (!d.Paragraphs[6].Range.Text.Contains("new look"))
                    return "False";
                if (d.Paragraphs[8].Range.Text.Contains("and SmartArt"))
                    return "False";
                if (d.Paragraphs[10].Range.Characters[5].Font.Bold == -1)
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
                int Index1 = 0;
                for (int Index2 = 1; Index2 < d.Paragraphs.Count; ++Index2)
                {
                    if (d.Paragraphs[Index2].Range.Text.Contains("Payment may be withheld under the following circumstances"))
                    {
                        Index1 = Index2;
                        break;
                    }
                }
                if (Index1 == 0)
                    return "False";
                if (d.Paragraphs[Index1].Range.Font.Bold == -1)
                    return "False";
                if (d.Paragraphs[Index1].Range.Font.Italic == -1)
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
                int Index1 = 0;
                for (int Index2 = 1; Index2 < d.Paragraphs.Count; ++Index2)
                {
                    if (d.Paragraphs[Index2].Range.Text.Contains("The muffin tray will still be hot"))
                    {
                        Index1 = Index2;
                        break;
                    }
                }
                if (Index1 == 0)
                    return "False";
                if (d.Paragraphs[Index1].Range.Characters[1].Text != "(")
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
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("automatic"))
                        return "False";
                }
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
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.ToUpper().Contains("WOODGROVE PLUS SAVINGS"))
                        return "True";
                }
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "False";
        }

        private static string cau33(Application a, _Document d)
        {
            try
            {
                if (d.Comments.Count != 0)
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
                int Index1 = 0;
                int num = 0;
                int Index2 = 0;
                for (int Index3 = 1; Index3 < d.Paragraphs.Count; ++Index3)
                {
                    if (d.Paragraphs[Index3].Range.Text.Contains("powerful new"))
                        Index1 = Index3;
                    if (d.Paragraphs[Index3].Range.Text.Contains("Themes and styles also help keep your"))
                        num = Index3;
                    if (d.Paragraphs[Index3].Range.Text.Contains("1-2 videographers"))
                        Index2 = Index3;
                }
                if (Index1 == 0)
                    return "False";
                if (num == 0 || Index2 == 0)
                    return "False";
                if (!d.Paragraphs[Index1].Range.Text.Contains("powerful new"))
                    return "False";
                if (d.Paragraphs[49].Range.Text.Contains("Themes and styles also help keep your"))
                    return "False";
                if (d.Paragraphs[Index2].Range.Characters[5].Font.Bold == -1)
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
                int num1 = 0;
                int num2 = 0;
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.ToLower().Contains("city"))
                        ++num1;
                    if (d.Paragraphs[Index].Range.Text.ToLower().Contains("community"))
                        ++num2;
                }
                if (num1 > 0)
                    return "False";
                if (num2 != 6)
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
                if (d.Comments.Count != 1)
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
                if (d.Sections.Count != 3)
                    return "False";
                if (d.Sections[2].PageSetup.TextColumns.Count != 2)
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
                if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Characters[3].Font.TextColor.RGB != -738131969)
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
