// Decompiled with JetBrains decompiler
// Type: MOS_WORD_LEARN.Final_Steps
// Assembly: MOS_WORD_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: E70F0E30-D79F-439D-990F-4092B40A7B30
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_WORD_19\Chay File Nay.exe

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using Shape = Microsoft.Office.Interop.Word.Shape;

namespace MOS_WORD_LEARN
{
    internal class Final_Steps
    {
        public static string CheckCau(int causo, Application a, _Document d)
        {
            switch (causo)
            {
                case 1:
                    return Final_Steps.cau21(a, d);
                case 2:
                    return Final_Steps.cau25(a, d);
                case 3:
                    return Final_Steps.cau20(a, d);
                case 4:
                    return Final_Steps.cau23(a, d);
                case 5:
                    return Final_Steps.cau26(a, d);
                case 6:
                    return Final_Steps.cau14(a, d);
                case 7:
                    return Final_Steps.cau16(a, d);
                case 8:
                    return Final_Steps.cau6(a, d);
                case 9:
                    return Final_Steps.cau24(a, d);
                case 10:
                    return Final_Steps.cau19(a, d);
                case 11:
                    return Final_Steps.cau22(a, d);
                case 12:
                    return Final_Steps.cau12(a, d);
                case 13:
                    return Final_Steps.cau13(a, d);
                case 14:
                    return Final_Steps.cau15(a, d);
                case 15:
                    return Final_Steps.cau15(a, d);
                case 16:
                    return Final_Steps.cau16(a, d);
                case 17:
                    return Final_Steps.cau17(a, d);
                case 18:
                    return Final_Steps.cau18(a, d);
                case 19:
                    return Final_Steps.cau19(a, d);
                case 20:
                    return Final_Steps.cau20(a, d);
                case 21:
                    return Final_Steps.cau21(a, d);
                case 22:
                    return Final_Steps.cau22(a, d);
                case 23:
                    return Final_Steps.cau23(a, d);
                case 24:
                    return Final_Steps.cau24(a, d);
                case 25:
                    return Final_Steps.cau25(a, d);
                case 26:
                    return Final_Steps.cau26(a, d);
                case 27:
                    return Final_Steps.cau27(a, d);
                default:
                    return "default final steps";
            }
        }

        private static string cau1(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Text Box 2";
                return d.Shapes[ref Index].TextFrame.TextRange.Text != "Requires committee review\r" ? "False(Requires committee review)" : "True";
            }
            catch (Exception ex)
            {
                return "False(something wrong)";
            }
        }

        private static string cau2(Application a, _Document d)
        {
            try
            {
                if (!a.ActiveWindow.View.ShowTabs)
                    return "False(Tab)";
                if (!a.ActiveWindow.View.ShowSpaces)
                    return "False(Spaces)";
                if (a.ActiveWindow.View.ShowAll)
                    return "False(turn off show All)";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau3(Application a, _Document d)
        {
            try
            {
                if (d.Paragraphs[1].Range.Font.Name != "Century Gothic")
                    return "False";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau4(Application a, _Document d)
        {
            try
            {
                if (!d.RemovePersonalInformation)
                    return "False(RemovePersonal)";
                object Index = (object)"Text Box 5";
                string text = d.Shapes[ref Index].TextFrame.TextRange.Text;
                return "True";
            }
            catch (Exception ex)
            {
                return "False(only RemovePersonal)";
            }
        }

        private static string cau5(Application a, _Document d)
        {
            try
            {
                if (d.Paragraphs[10].Alignment != WdParagraphAlignment.wdAlignParagraphJustify)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau6(Application a, _Document d)
        {
            try
            {
                if (!a.ActiveWindow.View.ShowTabs)
                    return "False(show Tab)";
                if (a.ActiveWindow.View.ShowSpaces)
                    return "False(don't show Spaces)";
                if (a.ActiveWindow.View.ShowHiddenText)
                    return "False(don't show HiddenText)";
                if (a.ActiveWindow.View.ShowAll)
                    return "False(turn off show All)";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau7(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Text Box 3";
                return d.Shapes[ref Index].TextFrame.TextRange.Text != "Draft\r" ? "False(Draft)" : "True";
            }
            catch (Exception ex)
            {
                return "False(something wrong)";
            }
        }

        private static string cau8(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Text Box 2";
                return d.Shapes[ref Index].TextFrame.TextRange.Text != "Requires committee review\r" ? "False(Requires committee review)" : "True";
            }
            catch (Exception ex)
            {
                return "False(something wrong)";
            }
        }

        private static string cau9(Application a, _Document d)
        {
            try
            {
                if (!a.ActiveWindow.View.ShowTabs)
                    return "False(Tab)";
                if (!a.ActiveWindow.View.ShowSpaces)
                    return "False(Spaces)";
                if (a.ActiveWindow.View.ShowAll)
                    return "False(turn off show All)";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau10(Application a, _Document d)
        {
            try
            {
                return !d.RemovePersonalInformation ? "False(RemovePersonal)" : "True";
            }
            catch (Exception ex)
            {
                return "False(only RemovePersonal)";
            }
        }

        private static string cau11(Application a, _Document d)
        {
            try
            {
                return !d.RemovePersonalInformation ? "False(RemovePersonal)" : "True";
            }
            catch (Exception ex)
            {
                return "False(only RemovePersonal)";
            }
        }

        private static string cau12(Application a, _Document d)
        {
            try
            {
                string subject = d.BuiltInDocumentProperties["Subject"].Value as string;

                if (!string.IsNullOrEmpty(subject) &&
                    subject.IndexOf("program", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "True";
                }
                else
                {
                    return "False";
                }
            }
            catch (Exception)
            {
                return "False(only RemovePersonal)";
            }
        }


        private static string cau13(Application a, _Document d)
        {
            try
            {
                Range content = d.Content;
                Find findHeading = content.Find;
                findHeading.Text = "Summarize main points";
                if (!findHeading.Execute()) return "False(heading not found)";

                Range afterHeading = d.Range(content.End, d.Content.End);
                Paragraph firstPara = afterHeading.Paragraphs.First;
                if (firstPara == null) return "False(no paragraph found)";

                Range paraRange = firstPara.Range;

                Find findWord = paraRange.Find;
                findWord.Text = "look";
                if (!findWord.Execute()) return "False(word 'look' not found)";

                Range wordRange = paraRange.Duplicate;
                wordRange.Start = paraRange.Start;
                wordRange.End = paraRange.End;

                d.Footnotes.Add(paraRange, Text: "resource");

                return "True";
            }
            catch (Exception)
            {
                return "False(error)";
            }
        }


        private static string cau14(Application a, _Document d)
        {
            try
            {
                object Index = (object)"ui";
                if (a.AutoCorrect.Entries[ref Index].Value != "aut")
                    return "False(aut)";
            }
            catch (Exception ex)
            {
                return "False(uil)";
            }
            return "True";
        }

        private static string cau15(Application a, _Document d)
        {
            try
            {
                // 1) Tìm heading "A primer on Improving profits"
                Range rng = d.Content;
                bool found = rng.Find.Execute("A primer on Improving profits",
                                              MatchCase: false,
                                              MatchWholeWord: true,
                                              Forward: true,
                                              Wrap: WdFindWrap.wdFindStop);
                if (!found)
                    return "False(heading not found)";

                Range headingRange = rng.Duplicate; // range chứa heading

                // 2) Tìm SmartArt có anchor nằm sau heading (nếu có). Nếu không, fallback sang SmartArt đầu tiên.
                Shape smartArtShape = null;
                foreach (Shape s in d.Shapes)
                {
                    try
                    {
                        if (s.Type == MsoShapeType.msoSmartArt)
                        {
                            Range anchor = s.Anchor as Range;
                            if (anchor != null && anchor.Start >= headingRange.End)
                            {
                                smartArtShape = s;
                                break;
                            }
                        }
                    }
                    catch
                    {
                        // ignore shapes we can't read anchor of
                    }
                }

                if (smartArtShape == null)
                {
                    // fallback: lấy SmartArt đầu tiên trong document
                    foreach (Shape s in d.Shapes)
                    {
                        if (s.Type == MsoShapeType.msoSmartArt)
                        {
                            smartArtShape = s;
                            break;
                        }
                    }
                }

                if (smartArtShape == null)
                    return "False(smartart not found)";

                // 3) Lấy tất cả nodes vào danh sách để xử lý theo thứ tự
                var allNodes = smartArtShape.SmartArt.AllNodes;
                var nodeList = new List<Microsoft.Office.Core.SmartArtNode>();
                foreach (Microsoft.Office.Core.SmartArtNode n in allNodes)
                    nodeList.Add(n);

                if (nodeList.Count == 0)
                    return "False(no nodes)";

                // 4) Lấy text từng node
                var texts = nodeList
                    .Select(n => (n.TextFrame2 != null && n.TextFrame2.TextRange != null)
                                 ? (n.TextFrame2.TextRange.Text ?? string.Empty).Trim()
                                 : string.Empty)
                    .ToList();

                // 5) Tìm vị trí "Recycle" và "Rework"
                int idxRecycle = texts.FindIndex(t => string.Equals(t, "Recycle", StringComparison.OrdinalIgnoreCase));
                int idxRework = texts.FindIndex(t => string.Equals(t, "Rework", StringComparison.OrdinalIgnoreCase));

                if (idxRecycle == -1 || idxRework == -1)
                    return "False(nodes not found)";

                // 6) Di chuyển text Recycle để đứng ngay trước Rework
                string recycleText = texts[idxRecycle];
                texts.RemoveAt(idxRecycle);

                // Sau khi remove, nếu recycle ở trước rework ban đầu thì index rework giảm 1
                int idxReworkAfterRemoval = idxRework;
                if (idxRecycle < idxRework) idxReworkAfterRemoval = idxRework - 1;

                // chèn recycle trước rework
                texts.Insert(idxReworkAfterRemoval, recycleText);

                // 7) Gán lại text mới cho các node theo thứ tự nodeList
                for (int i = 0; i < nodeList.Count && i < texts.Count; i++)
                {
                    try
                    {
                        nodeList[i].TextFrame2.TextRange.Text = texts[i];
                    }
                    catch
                    {
                        // tiếp tục nếu một node không gán được
                    }
                }

                return "True";
            }
            catch (Exception ex)
            {
                // trả về thông báo lỗi để debug nếu cần
                return "False(error: " + ex.Message + ")";
            }
        }

        private static string cau16(Application a, _Document d)
        {
            try
            {
                if (a.ActiveWindow.View.ShowTabs)
                    return "False(not Tab)";
                if (a.ActiveWindow.View.ShowSpaces)
                    return "False(not Spaces)";
                if (!a.ActiveWindow.View.ShowHiddenText)
                    return "False(Hidden text)";
                if (a.ActiveWindow.View.ShowAll)
                    return "False(turn off show All)";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau17(Application a, _Document d)
        {
            try
            {
                try
                {
                    if (a.ActiveWindow.View.ShowTabs)
                        return "False(not Tab)";
                    if (a.ActiveWindow.View.ShowSpaces)
                        return "False(not Spaces)";
                    if (a.ActiveWindow.View.ShowHiddenText)
                        return "False(not Hidden text)";
                    if (!a.ActiveWindow.View.ShowParagraphs)
                        return "False(paragraph mark)";
                    if (a.ActiveWindow.View.ShowAll)
                        return "False(turn off show All)";
                }
                catch (Exception ex)
                {
                    return "False (Something not finish!)";
                }
                return "True";
            }
            catch (Exception ex)
            {
                return "False(Something not finish!)";
            }
        }

        private static string cau18(Application a, _Document d)
        {
            try
            {
                if (!a.ActiveWindow.View.ShowTabs)
                    return "False(Tab)";
                if (!a.ActiveWindow.View.ShowSpaces)
                    return "False(Spaces)";
                if (a.ActiveWindow.View.ShowHiddenText)
                    return "False(not Hidden text)";
                if (a.ActiveWindow.View.ShowParagraphs)
                    return "False(not paragraph mark)";
                if (a.ActiveWindow.View.ShowAll)
                    return "False(turn off show All)";
            }
            catch (Exception ex)
            {
                return "False(Something not finish!)";
            }
            return "True";
        }

        private static string cau19(Application a, _Document d) => "True";

        private static string cau20(Application a, _Document d)
        {
            try
            {
                if (d.CompatibilityMode == 11)
                    return "False(File==>Convert)";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau21(Application a, _Document d)
        {
            try
            {
                if (!d.Name.StartsWith("Notes."))
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
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau23(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Text Box 3";
                if (!d.Shapes[ref Index].TextFrame.TextRange.Text.Contains("animals"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau24(Application a, _Document d)
        {
            try
            {
                if (d.RemovePersonalInformation)
                    return "False(không xóa PersonalInformation)";
                if (d.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.Count != 0)
                    return "False";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau25(Application a, _Document d)
        {
            try
            {
                if (d.Name != "Memo.txt")
                    return "False(luu dang text)";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau26(Application a, _Document d)
        {
            try
            {
                object Index = (object)"Text Box 2";
                if (!d.Shapes[ref Index].TextFrame.TextRange.Text.Contains("bicycles"))
                    return "False";
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau27(Application a, _Document d)
        {
            try
            {
            }
            catch (Exception ex)
            {
                return "False (Something not finish!)";
            }
            return "True";
        }

        private static string cau6_2010(Application a, _Document d)
        {
            try
            {
                foreach (OtherCorrectionsException correctionsException in a.AutoCorrect.OtherCorrectionsExceptions)
                {
                    if (correctionsException.Name == "Contoso" || correctionsException.Name == "contoso")
                        return "True";
                }
                return "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau12_2010(Application a, _Document d)
        {
            try
            {
                return a.AutoCorrect.ReplaceText ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau26_2010(Application a, _Document d)
        {
            try
            {
                return !d.Windows.SyncScrollingSideBySide ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau28_2010(Application a, _Document d)
        {
            try
            {
                int verticalPercentScrolled = d.ActiveWindow.VerticalPercentScrolled;
                int splitVertical = d.ActiveWindow.SplitVertical;
                return verticalPercentScrolled > 0 && splitVertical > 0 ? d.ActiveWindow.Split.ToString() : "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau29_2010(Application a, _Document d)
        {
            try
            {
                return d.ProtectionType == WdProtectionType.wdAllowOnlyComments ? "True" : "False";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }

        private static string cau33_2010(Application a, _Document d)
        {
            try
            {
                object Index1 = (object)"IIG";
                // ISSUE: variable of a compiler-generated type
                Reviewer reviewer1 = a.ActiveWindow.View.Reviewers[ref Index1];
                if (reviewer1.Visible)
                    return "False";
                object Index2 = (object)"Admin";
                // ISSUE: variable of a compiler-generated type
                Reviewer reviewer2 = a.ActiveWindow.View.Reviewers[ref Index2];
                return !reviewer2.Visible ? "False" : "True";
            }
            catch (Exception ex)
            {
                return "False";
            }
        }
    }
}
