// Decompiled with JetBrains decompiler
// Type: MOS_WORD_LEARN.Style
// Assembly: MOS_WORD_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.CompilerServices;

namespace MOS_WORD_LEARN
{
    internal class Style
    {
        public static string CheckCau(int causo, Application a, _Document d)
        {
            switch (causo)
            {
                case 1:
                    return Style.cau1(a, d);
                case 2:
                    return Style.cau2(a, d);
                case 3:
                    return Style.cau3(a, d);
                case 4:
                    return Style.cau4(a, d);
                case 5:
                    return Style.cau5(a, d);
                case 6:
                    return Style.cau6(a, d);
                case 7:
                    return Style.cau7(a, d);
                case 8:
                    return Style.cau8(a, d);
                case 9:
                    return Style.cau9(a, d);
                default:
                    return "False";
            }
        }

        private static string cau1(Application a, _Document d)
        {
            try
            {
                int index = 1;

                // Tìm đoạn có chứa đoạn văn bản chỉ định
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("These distractions could be on the"))
                {
                    index++;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                // Lấy style của đoạn văn bản đó
                Paragraph para = d.Paragraphs[index];
                var styleObj = para.Range.get_Style(); // hoặc .Style nếu không bị lỗi
                var style = styleObj as Microsoft.Office.Interop.Word.Style;

                if (style == null)
                    return "False";

                if (style.NameLocal != "Intense Emphasis")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }


        private static string cau2(Application a, _Document d)
        {
            try
            {
                int index = 1;

                // Tìm đoạn chứa "My Coffee"
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("My Coffee"))
                {
                    index++;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                // Dùng dynamic để truy cập CharacterStyle
                Paragraph para = d.Paragraphs[index];
                dynamic rng = para.Range;
                var style = rng.CharacterStyle as Microsoft.Office.Interop.Word.Style;

                if (style != null)
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau3(Application a, _Document d)
        {
            try
            {
                int index1 = 1;

                // Tìm đoạn chứa "Information"
                while (index1 <= d.Paragraphs.Count &&
                       !d.Paragraphs[index1].Range.Text.Contains("Information"))
                {
                    index1++;
                }

                if (index1 > d.Paragraphs.Count)
                    return "False";

                var para1 = d.Paragraphs[index1];
                var style1 = para1.Range.get_Style() as Microsoft.Office.Interop.Word.Style;
                if (style1 == null || style1.NameLocal != "Heading 1")
                    return "False";

                // Tìm đoạn chứa "Games Times"
                int index2 = 1;
                while (index2 <= d.Paragraphs.Count &&
                       !d.Paragraphs[index2].Range.Text.Contains("Games Times"))
                {
                    index2++;
                }

                if (index2 > d.Paragraphs.Count)
                    return "False";

                var para2 = d.Paragraphs[index2];
                var style2 = para2.Range.get_Style() as Microsoft.Office.Interop.Word.Style;
                if (style2 == null || style2.NameLocal != "Heading 1")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau4(Application a, _Document d)
        {
            try
            {
                int index1 = 1;

                // Tìm đoạn chứa "Online"
                while (index1 <= d.Paragraphs.Count &&
                       !d.Paragraphs[index1].Range.Text.Contains("Online"))
                {
                    index1++;
                }

                if (index1 > d.Paragraphs.Count)
                    return "False";

                var para1 = d.Paragraphs[index1];
                var style1 = para1.Range.get_Style() as Microsoft.Office.Interop.Word.Style;
                if (style1 == null || style1.NameLocal != "Heading 2")
                    return "False";

                int index2 = 1;

                // Tìm đoạn chứa "In Person"
                while (index2 <= d.Paragraphs.Count &&
                       !d.Paragraphs[index2].Range.Text.Contains("In Person"))
                {
                    index2++;
                }

                if (index2 > d.Paragraphs.Count)
                    return "False";

                var para2 = d.Paragraphs[index2];
                var style2 = para2.Range.get_Style() as Microsoft.Office.Interop.Word.Style;
                if (style2 == null || style2.NameLocal != "Heading 2")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau5(Application a, _Document d)
        {
            try
            {
                int num = 0;
                for (int Index = 1; Index < d.Paragraphs.Count; ++Index)
                {
                    if (d.Paragraphs[Index].Range.Text.Contains("Kids love dinosaurs"))
                    {
                        num = Index;
                        break;
                    }
                }
                if (num == 0)
                    return "False";
                if (!d.Paragraphs[num + 2].Range.Text.Contains("Save time in Word with"))
                    return "False";
                if (d.Paragraphs[num + 2].Range.Font.Bold != 0)
                    return "False";
                if (d.Paragraphs[num + 2].Range.Font.TextColor.RGB != -16777216)
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
                int index = 1;

                // Tìm đoạn chứa câu cần kiểm tra
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("This event begins at 7 pm on "))
                {
                    index++;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                // Truy cập dynamic để lấy CharacterStyle
                var para = d.Paragraphs[index];
                dynamic rng = para.Range;
                var style = rng.CharacterStyle as Microsoft.Office.Interop.Word.Style;

                if (style == null || style.NameLocal != "Intense Emphasis")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }


        private static string cau7(Application a, _Document d)
        {
            try
            {
                int index = 1;

                // Tìm đoạn chứa "Genre"
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("Genre"))
                {
                    index++;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                var para1 = d.Paragraphs[index];
                dynamic rng1 = para1.Range;
                var style1 = rng1.CharacterStyle as Microsoft.Office.Interop.Word.Style;

                if (style1 == null || style1.NameLocal != "Subtle Emphasis")
                    return "False";

                // Tìm đoạn chứa "External Parther"
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("External Parther"))
                {
                    index++;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                var para2 = d.Paragraphs[index];
                dynamic rng2 = para2.Range;
                var style2 = rng2.CharacterStyle as Microsoft.Office.Interop.Word.Style;

                if (style2 == null || style2.NameLocal != "Subtle Emphasis")
                    return "False";
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }


        private static string cau8(Application a, _Document d)
        {
            try
            {
                int index = 1;
                // Tìm đoạn văn có chứa chuỗi "Microsoft Office Specialist (MOS)"
                while (index <= d.Paragraphs.Count &&
                       !d.Paragraphs[index].Range.Text.Contains("Microsoft Office Specialist (MOS)"))
                {
                    ++index;
                }

                if (index > d.Paragraphs.Count)
                    return "False";

                // Lấy style đoạn văn này
                var paraStyle = d.Paragraphs[index].Range.get_Style() as Microsoft.Office.Interop.Word.Style;
                if (paraStyle == null)
                    return "False";

                if (paraStyle.NameLocal != "Title")
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
                int index = 1;

                string[] requiredTexts = new string[]
                {
            "Organize",
            "Cooking",
            "Set up a time line",
            "Cook ahead",
            "Minimum of ingredients to prepare recipes"
                };

                foreach (string text in requiredTexts)
                {
                    // Tìm đoạn văn chứa nội dung yêu cầu
                    while (index <= d.Paragraphs.Count && !d.Paragraphs[index].Range.Text.Contains(text))
                        ++index;

                    if (index > d.Paragraphs.Count)
                        return "False";

                    // Kiểm tra style đoạn văn
                    var paraStyle = d.Paragraphs[index].Range.get_Style() as Microsoft.Office.Interop.Word.Style;
                    if (paraStyle == null)
                        return "False";

                    if (paraStyle.NameLocal != "Heading 1")
                        return "False";
                }
            }
            catch (Exception)
            {
                return "False";
            }

            return "True";
        }

        private static string cau10(Application a, _Document d)
        {
            try
            {
            }
            catch (Exception ex)
            {
                return "False";
            }
            return "True";
        }
    }
}
