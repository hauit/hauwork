using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1.Base
{
    public class ClsQuestion
    {
        public int Index { get; private set; }
        public string EngQuestion { get; private set; }
        public string VnQuestion { get; private set; }
        public ClsQuestion(int index, string vn, string eng)
        {
            this.Index = index;
            this.EngQuestion = eng;
            this.VnQuestion = vn;
        }
    }

    public static class ClsListQuestion
    {
        private static List<ClsQuestion> listQuestion = new List<ClsQuestion>
        {
            new ClsQuestion(1, "Hiển thị tất cả các ký tự đặc biệt trong Word", "Show all special symbols in word"),
            new ClsQuestion(2, "Chèn một ngắt phần ( Section break ) tên Continuous tại vị trí bắt đầu của tiêu đề : \"Where to find us\"", "Insert a Continuous section break at the beginning of the \"Where to find us\" heading."),
            new ClsQuestion(3, "Thêm một ngắt phần ( Section break ) tên Next page ngay trước \"Vegetables\" ở phía dưới của trang 1.", "Add a Next page section break immediately before \"Vegetables\" at the bottom of page 1."),
            new ClsQuestion(4, "Chèn ngắt trang trước tiêu đề \" Tuning\"", "Insert page break before the \"Tuning\" heading"),
            new ClsQuestion(5, "Chèn ngắt dòng văn bản bên trái từ \"Observation\" của tiêu đề.", "Insert text wrapping break to the left of the word 'Observation' of the heading."),
            new ClsQuestion(6, "Trong danh sách 2 cột, thêm một ngắt cột (Column break) ngay trước văn bản \"Four String\".", "In the two - column list, add a Column break immediately before the text \"Four - String\"."),
            new ClsQuestion(7, "Chia 4 đoạn văn trước cái hình thành 2 cột với độ giãn của cột là \"0.3\" (0.8cm)", "Split the four paragraphs before the picture into two columns with column spacing of \"0.3\" (0.8cm)"),
            new ClsQuestion(8, "Thay đổi hướng giấy của chỉ trang 3 thành nằm ngang (Landscape)", "Change the orientation of only page 3 to Landscape."),
            new ClsQuestion(9, "Sau ngắt phần (Section break) thay đổi hướng giấy thành Landscape.", "After the section break, change the orientation to Landscape."),
            new ClsQuestion(10,"Áp dụng lề giấy tên Moderate cho tài liệu.","Apply Moderate Margin to document."),
            new ClsQuestion(11,"Thay đổi các lề giấy của tài liệu, thiết lập lề trên và dưới là \"0.75\" và lề trái và lề phải là \"0.5\".","You work for Humongous Insurance. You are creating a newletter insert that provides information about insurance costs. Change the page margins of the document set the top and bottom margins to \"0.75\" (\"1.9cm\") and the left and right margins to \"0.5\" (\"1.27cm\")"),
            new ClsQuestion(12,"Trong phần nội dung \"Punlishing processes\" chia 3 đoạn văn sau tiêu đề thành 2 cột.","In the \"Punlishing processes\" section, split the three paragraphs after the heading into two columns."),
            new ClsQuestion(13,"",""),
            new ClsQuestion(14,"",""),
            new ClsQuestion(15,"",""),
            new ClsQuestion(16,"",""),
            new ClsQuestion(17,"",""),
            new ClsQuestion(18,"",""),
            new ClsQuestion(19,"",""),
            new ClsQuestion(20,"",""),
            new ClsQuestion(21,"",""),
            new ClsQuestion(22,"",""),
            new ClsQuestion(23,"",""),
            new ClsQuestion(24,"",""),
            new ClsQuestion(25,"",""),
            new ClsQuestion(26,"",""),
            new ClsQuestion(27,"",""),
            new ClsQuestion(28,"",""),
            new ClsQuestion(29,"",""),
        };

        public static string GetEngQuestion(int index)
        {
            var obj = listQuestion.Where(x => x.Index == index).FirstOrDefault();
            if(obj == null)
            {
                return string.Empty;
            }
            return obj.EngQuestion;
        }

        public static string GetVNQuestion(int index)
        {
            var obj = listQuestion.Where(x => x.Index == index).FirstOrDefault();
            if (obj == null)
            {
                return string.Empty;
            }
            return obj.VnQuestion;
        }

    }

}
