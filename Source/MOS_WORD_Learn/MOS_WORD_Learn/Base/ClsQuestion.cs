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
            new ClsQuestion(13,"Áp dụng tập hợp các kiểu (Style) tên Centered cho tài liệu.","You are preparing a brochure for Southridge Video. You plan to distribute the brochure electronically and in print. Apply the Centered style set to the document."),
            new ClsQuestion(14,"Áp dụng bộ các kiểu (style) tên Lines (Stylish) cho tài liệu.","You work for a publishing company. You are preparing a document that will be sent to potential clients. Apply the Lines (Stylish) style set to the document."),
            new ClsQuestion(15,"Áp dụng chủ đề Ion, thay đổi định dạng của tài liệu thành Shaded.","Apply theme Ion, modify Formatting to Shaded."),
            new ClsQuestion(16,"Thay đổi bộ phông cho toàn bộ tài liệu là Candara.","Change the font set for entry document to \"Candara\""),
            new ClsQuestion(17,"Thêm đường viền trang dạng Box màu Dark Blue, Accent 1 độ rộng là 3pt cho toàn bộ tài liệu.","The owner of Margie's Travel has asked you to finish formatting an event flyer. Add a 3pt Dark Blue, Accent 1 Box page border to the whole document."),
            new ClsQuestion(18,"Thêm đường viền trang dạng Box với độ rộng là 3pt cho toàn bộ tài liệu.","Apply box page border, width 3pt to entry document."),
            new ClsQuestion(19,"Áp dụng màu Green, Accent 6, Lighter 80% làm màu nền của trang.","Apply Green, Accent 6, Lighter 80% as the page background color."),
            new ClsQuestion(20,"Áp dụng hình mờ cho tất cả các trang, sử dụng kiểu DO NOT COPY 1.","Apply watermark to all of pages, use \"DO NOT COPY 1\" style."),
            new ClsQuestion(21,"Thêm trang bìa tên Banded. Xóa trình giữ chỗ \"[Company address]\"","Add a Banded cover page. Delete the \"[Company address]\" placeholder."),
            new ClsQuestion(22," Chèn vào đầu trang kiểu Integral cho tất cả các trang của tài liệu ngoại trừ trang số 1.","You work for Fourth Coffee. You are finalizing a training manual for employees who will bake muffins for the coffee shop. Display the Integral header on all pages of the document except page 1."),
            new ClsQuestion(23,"Chèn vào đầu trang kiểu Banded cho tất cả các trang ngoại trừ trang đầu tiên.","Insert Banded header all page except the first page."),
            new ClsQuestion(24,"Chèn vào chân trang kiểu Whisp chỉ cho trang lẻ.","Insert Whisp footer only the Odd pages."),
            new ClsQuestion(25,"Chèn vào đầu trang kiểu Austin và chèn vào tiêu đề văn bản \"MOS\".","Insert Austin header and insert the title with the text \"MOS\""),
            new ClsQuestion(26,"Thêm đánh số trang kiểu Accent Bar 2 ở phía dưới của mỗi trang.","Add Accent Bar 2 page numbering at the bottom of each page."),
            new ClsQuestion(27,"Ở phía trên của tất cả các trang chèn đánh số trang với mẫu Plain Number 3.","At the top of all pages insert the plain number 3 page number"),
            new ClsQuestion(28,"Thay đổi định dạng của văn bản bắt đầu \"These distractions could be … \" và kết thúc \" …. Poor presentation skills.\" Kiểu Intense Emphasis.","Change the format of the text beginning with \"These distractions could be … \" and ending with \" …. Poor presentation skills.\" to Intense Emphasis."),
            new ClsQuestion(29,"Áp dụng kiểu Intense Emphasis cho văn bản \"My coffee\" bên dưới hình tách cà phê.","Appy style intense emphasis to the text \"My coffee\" below the picture cup of coffee."),
            new ClsQuestion(30,"Áp dụng kiểu Heading 1 cho văn bản \"Information\", \"Game Times\", \"Managers Meeting\", and \"Registration Dates\".","Apply style Heading 1 to text \"Information\", \"Game Times\", \"Managers Meeting\", and \"Registration Dates\"."),
            new ClsQuestion(31,"Áp dụng kiểu Heading 2 cho văn bản \"Online\", \"In person\", and \"By email\".","Appy style Heading 2 to text \"Online\", \"In person\", and \"By email\"."),
            new ClsQuestion(32,"Trong phần nội dung \"Kid love dinosaurs\" sao chép định dạng của đoạn đầu tiên và áp dụng nó cho đoạn văn thứ 2.","In the \"Kid love dinosaurs\" section, copy the formatting of the first paragraph and apply it to the second paragraph."),
            new ClsQuestion(33,"Áp dụng kiểu Intense Emphasis cho đoạn văn sau cái hình.","Apply the Intense Emphasis style to the paragraph after the picture."),
            new ClsQuestion(34,"Trong dòng đầu tiên của bảng, áp dụng kiểu Subtle Emphasis cho văn bản.","In the first table row, apply the Subtle Emphasis style to the text."),
            new ClsQuestion(35,"Trong phần nội dung \" Making moments last forever!\" Chuyển đổi 5 đoạn văn bắt đầu với \"Corporate events\" thành danh sách đánh dấu đầu dòng.","In the \"Making moments last forever!\" section, convert the five paragraphs starting with \"Corporate events\" to a bulleted list."),
            new ClsQuestion(36,"Trong phần nội dung \"Tuning\" gán đánh dấu đầu dòng văn bản \" Standard …. Ukulele Style\" sử dụng hình tên gold.jpg trong thư mục Document.","Into section \"Tuning\" bullet to six line \"Standard …. Ukulele Style\" use the picture gold.jpg in document folder."),
            new ClsQuestion(37,"Trong phần nội dung \"Our Video Collection\", thay đổi danh sách thành ký tự đánh dấu đầu dòng. Sử dụng phông Segoe UI Emoji và mã ký tự \"25B8\".","In the \"Our Video Collection\" section, change the list to use a custom bullet character. Use the Segoe UI Emoji font and character code \"25B8\" ( the black right - pointing small triangle symbol)"),
            new ClsQuestion(38,"Xác định danh sách định nghĩa các loại thực phẩm hữu cơ và thay đổi đánh số bằng cách sử dụng các ký tự A, B, C thay cho các số.","Locate the list that defines organic food categories and change the numbering to use uppercase letters (A, B, C) instead of numbers."),
            new ClsQuestion(39,"Trong phần nội dung \"More dinosaur facts\", thay đổi cấp danh sách của \"Velociraptor\" thành cấp 3.","In the \"More dinosaur facts\" section, change the list level for \"Velociraptor\" to Level 3 ."),
            new ClsQuestion(40,"Ở cuối tài liệu, thay đổi độ giãn dòng của 2 đoạn văn cuối chính xác thành 14pt.","At the end of the document, change the line spacing of the last two paragraphs to exactly 14pt."),
            new ClsQuestion(41,"Thiết lập độ giãn dòng 1.4 cho toàn bộ tài liệu.","Set the line spacing to 1.4 lines for the entire document."),
            new ClsQuestion(42,"Áp dụng độ giãn dòng cỡ 18pt kiểu At least cho 3 đoạn văn ở cuối trang 2.","Apply At least 18pt line spacing to three paragraph at the end of second page."),
            new ClsQuestion(43,"Trong phần nội dung \"Prerequisites\", bắt đầu danh sách bằng \"111\".","In the \"Prerequisites\" section, start the numbered list at \"111\"."),
            new ClsQuestion(44,"Trong phần nội dung \"Top sellers\", tiếp tục đánh số ở trên đầu cột thứ 2 của danh sách, để các hạng mục của danh sách được đánh số là từ 1 đến 6.","In the \"Top sellers\" section, continue the numbering of the list at the top of the second column, so the list items are numbered from 1 through 6."),
            new ClsQuestion(45,"Áp dụng định dạng cho các đoạn văn sau danh sách đánh số để nó tiếp tục đánh số.","Apply format to the paragraphs after the numbering so that it continue numbering."),
            new ClsQuestion(46,"Sao chép văn bản \" Blue Yonder Airlines is …. Institute students\" và dán chỉ văn bản ở cuối đoạn văn nằm gần cái hình.","Copy the text \"Blue Yonder Airlines is …. Institute students\" and paste only the text at the end of the paragraph beneath the graphic."),
            new ClsQuestion(47,"Tô sáng văn bản trong dòng bắt đầu bằng \"Thorp\" thành Turquoise.","Highlight the text in the row starting with \"Thorp\" in Turquoise."),
            new ClsQuestion(48,"Trong phần nội dung \"Where to find us\", trong đoạn văn rỗng chèn hình tên Map từ thư mục Picture.","In the \"Where to find us\" section, in the blank paragraph, insert the Map image from the Pictures folder."),
            new ClsQuestion(49,"Trên khoảng trắng ở phía dưới của trang 2, chèn một hình dạng (shape) tên Scroll: Horizontal chứa văn bản \"Remember your calculator!\". Vị trí của hình ở phía dưới chính giữa của trang, với văn bản bao quanh thành hình vuông (square text wrapping). Kích thước chính xác của hình thì không quan trọng.","In the blank space at the bottom of page 2, insert a shape named Scroll: Horizontal contains the text \"Remember your calculator!\". Position the shape at the bottom center of the page, with the text wrapped in a square. The exact size of the shape doesn't matter."),
            new ClsQuestion(50,"Thêm sơ đồ SmartArt Vertical Bullet List bên dưới văn bản \"Once again, we would like …. \". Thêm văn bản \"Frank Miller\" cho hình màu xanh ở phía trên.","Add a SmartArt Vertical Bullet List diagram below the text \"Once again, we would like …. \". Add the text \"Frank Miller\" to the top green shape."),
            new ClsQuestion(51,"Áp dụng kiểu chữ nghệ thuật (WordArt) tên Gradient Fill Gray cho văn bản \"We heart you ….\" ở phía trên cùng của tài liệu. Thẳng hàng WordArt thành chính giữa của trang.","Apply a WordArt style named Gradient Fill Gray to the text \"We heart you ….\" at the top of the document. Align the WordArt to the center of the page."),
            new ClsQuestion(52,"Trong phần nội dung \"Checking Accounts\", trong hộp văn bản màu xanh đen hãy chèn văn bản \"Anytime Account Access\".","In the \"Checking Accounts\" section, in the dark blue text box, insert the text \"Anytime Account Access\"."),
            new ClsQuestion(53,"Trong phần nội dung \"Serving, thay đổi ngắt dòng văn bản (text wrapping) cho cái hình thành hình vuông. ","In the \"Serving\" section, change the text wrapping for the shape to a square."),
            new ClsQuestion(54,"Trong phần nội dung \"Maintain bicycles\", thay đổi ngắt dòng văn bản (Text wrapping) cho cái hình thành hình vuông.","In the \"Maintain bicycles\" content section, change the Text wrapping for the shape to a square."),
            new ClsQuestion(55,"Tùy chỉnh dòng văn bản nằm bên trái cái hình.","Configure the line text place at the left of the image."),
            new ClsQuestion(56,"Trong phần nội dung \"Favorite dinosaurs\", trong đoạn văn rỗng ở cuối trang, sử dụng tính năng mô hình 3D (3D Models) để chèn mô hình Triceratops từ thư mục 3D Objects. Vị trí mô hình là Inline with Text.","In the \"Favorite dinosaurs\" content section, in the blank paragraph at the end of the page, use the 3D Models feature to insert the Triceratops model from the 3D Objects folder. Position the model is Inline with Text."),
            new ClsQuestion(57,"Trong phần nội dung \"Description\", sử dụng tính năng mô hình 3D (3D Models) để chèn mô hình PillPack từ thư mục 3D Objects vào đoạn văn rỗng. Vị trí mô hình là Inline with Text.","In the \"Description\" section, use the 3D Models feature to insert the PillPack model from the 3D Objects folder into the blank paragraph. The model position is Inline with Text."),
            new ClsQuestion(58,"Trong phần nội dung \"Manufacturing Process\", gán văn bản thay thế \"Process flow\" cho sơ đồ. (Hãy chắc rằng chọn toàn bộ sơ đồ).","In the \"Manufacturing Process\" section, assign the alternative text description \"Process flow\" to the SmartArt graphic. (Be sure to select the entire SmartArt graphic.)"),
            new ClsQuestion(59,"Trên trang 1, thay đổi màu đường viền của hình thành Blue, Accent 1, Darker 25%.","On page 1, change the shape's border color to Blue, Accent 1, Darker 25%."),
            new ClsQuestion(60,"Áp dụng kiểu ảnh (picture style) tên Metal Frame cho hình bánh kem.","Apply the Metal Frame picture style to the picture of pastries."),
            new ClsQuestion(61,"Gỡ bỏ nền của hình ghita, cẩn thận không cắt phần đàn ghita.","Remove the background of the guitar image, being careful not to cut off the guitar."),
            new ClsQuestion(62,"Áp dụng hiệu ứng nghệ thuật (Art Effect) tên Pencil Grayscale cho hình gần trên cùng của trang.","Apply an Art Effect named Pencil Grayscale to the picture near the top of the page."),
            new ClsQuestion(63,"Trong phần nội dung \"Basic dinosaur facts\", áp dụng hiệu ứng nghệ thuật Pencil Sketch cho hình hóa thạch.","In the \"Basic dinosaur facts\" section, apply the Pencil Sketch artistic effect to the fossil image."),
            new ClsQuestion(64,"Áp dụng hiệu ứng Soft Round Bevel cho hình tách café trên chiếc đĩa.","Apply a Soft Round Bevel effect to the coffee cup image on the saucer."),
            new ClsQuestion(65,"Áp dụng hiệu ứng Soft Round Bevel cho sơ đồ gần phía trên của trang 1.","Apply the Soft Round Bevel shape effect to the SmartArt near the top of page 1."),
            new ClsQuestion(66,"Trong phần nội dung \"Overview\", áp dụng hiệu ứng Soft Round Bevel cho sơ đồ.( Hãy chắc chắn là chọn toàn bộ sơ đồ).","In the \"Overview\" section, apply the Soft Round Bevel shape effect to the SmartArt graphic. (Make sure to select the entire SmartArt graphic)."),
            new ClsQuestion(67,"Thay đổi cách hiển thị nội dung của sơ đồ để hiển thị từ 1 đến 8 từ trái sang phải. Không thay đổi thứ tự các mục trong Text Pane.","Change the display of the SmartArt graphic content are listed from 1 to 8 from left to right. Do not change the order of items in the Text Pane."),
            new ClsQuestion(68,"Với sơ đồ trong phần nội dung \"My coffee\" áp dụng màu Transparent Gradient Range - Accent 1.","With the SmartArt into the section \"My coffee\" apply color Transparent Gradient Range - Accent 1."),
            new ClsQuestion(69,"Sắp xếp văn bản trong sơ đồ để \"organically grown\" nằm trên \"carefully roasted\".","Arrange the text in the SmartArt so that \"organically grown\" is above \"carefully roasted\"."),
            new ClsQuestion(70,"Áp dụng hiệu ứng đổ bóng Offset : Bottom (Offset Bottom) cho văn bản \"Congratulations!\"","Apply the Offset : Bottom (Offset Bottom) shadow effect to the \"Congratulations!\" text."),
            new ClsQuestion(71,"Trong phần nội dung \"Grades\", trong đoạn văn rỗng chèn bảng có 6 dòng và 2 cột. Trong dòng đầu tiên của bảng chèn \"Grade\" vào ô bên trái và \"Score Range\" vào ô bên phải. Sau đó điều chỉnh bảng phù hợp với nội dung.","In the \"Grades\" section, in the blank paragraph insert a table with 6 rows and 2 columns. In the first row of the table insert \"Grade\" in the left cell and \"Score Range\" in the right cell. Then adjust the table to fit the content."),
            new ClsQuestion(72,"Trong phần nội dung \"Banking Fees\", chuyển đổi văn bản được phân tách bởi dấu Tab thành bảng có 2 cột. Chấp nhận các tự động điều chỉnh là mặc định.","In the \"Banking Fees\" content section, converts tab-separated text into a 2-column table. Accepts auto-fit as default."),
            new ClsQuestion(73,"Trong phần nội dung \"Fill Material\" , điều chỉnh kích thước của bảng để mỗi cột có chiều rộng là 2.2\" (5.59cm)","In the \"Fill Material\" section, adjust the size of the table so that each column is 2.2\" (5.59cm) wide."),
            new ClsQuestion(74,"Trong bảng bên dưới \" Wilderness Summary\", điều chỉnh để chiều rộng mỗi cột là như nhau.","In the table below \"Wilderness Summary\", adjust so that each column is the same width."),
            new ClsQuestion(75,"Trong phần nội dung \"Contact Us\" , gộp tất cả các ô trong dòng đầu tiên của bảng.","In the \"Contact Us\" content section, merge all the cells in the first row of the table."),
            new ClsQuestion(76,"Trong phần nội dung \"Programs at other universities\", gộp tất cả các ô trong dòng đầu tiên của bảng.","In the \"Programs at other universities\" section, merge all the cells in the first row of the table."),
            new ClsQuestion(77,"Tùy chỉnh bảng để dòng đầu tiên tự động lặp lại như dòng tiêu đề ở đầu của mỗi trang.","Configure the table so that the first row automatically repeats as a header row at the top of each page."),
            new ClsQuestion(78,"Trong phần nội dung \"Geologic eras\" sắp xếp dữ liệu bảng tăng dần bởi \"Geologic period\" và sau đó tăng dần bởi \"Dinosaur\".","In the \"Geologic eras\" content section, sort the table data ascending by \"Geologic period\" and then ascending by \"Dinosaur\"."),
            new ClsQuestion(79,"Trong dòng đầu tiên của bảng, thiết lập độ giãn ô (cell spacing) là 0.02\" (0.05cm).","In the first row of the table, set the cell spacing to 0.02\" (0.05cm)."),
            new ClsQuestion(80,"Áp dụng kiểu Grid Table 1 Light - Accent 2 cho bảng.","Apply the Grid Table 1 Light - Accent 2 style to the table."),
            new ClsQuestion(81,"Trong phần nội dung \"Contact Us\" gộp các ô trong dòng đầu tiên của bảng.","In the \"Contact Us\" content section, merge the cells in the first row of the table."),
            new ClsQuestion(82,"Trong phần nội dung \" Exclusions\", xóa tất cả định dạng từ đoạn văn bắt đầu \"Payment may be withelp\".","In the \"Exclusions\" content section, clear all formatting from the paragraph starting with \"Payment may be withelp\"."),
            new ClsQuestion(83,"Xóa định dạng văn bản \"Fourth Coffee\" phía trên đầu của tài liệu.","Clear formatting from the text \"Fourth Coffee\"  on the top of the document."),
            new ClsQuestion(84,"Trên đầu các trang (Header) của tài liệu, áp dụng hiệu ứng Fill : Blue, Accent color 1; Shadow cho văn bản.","\"You are reviewing a document that will be distributed to potential clients."),
            new ClsQuestion(85,"Gạch chân văn bản \"Fourth Coffee\" với kiểu gạch chân dày (Thick Underline) và thay đổi phông chữ thành Algerian. ","At the top of the document's pages (Header), apply the Fill effect: Blue, Accent color 1; Shadow to the text.\""),
            new ClsQuestion(86,"Trong phần nội dung \"Depanning\", chèn ký tự hình nhiệt kế trước cụm từ \"The muffin tray will still be hot!\". Sử dụng phông chữ Webdings và mã ký tự \"225\" (ký tự hình nhiệt kế).","Underline the \"Fourth Coffee\" text with a Thick Underline style and change the font to Algerian."),
            new ClsQuestion(87,"Thêm ký tự Trade Mark ngay sau văn bản \"WareWolf\" trên trang số 4. ","In the \"Depanning\" text, insert a thermometer symbol before the phrase \"The muffin tray will still be hot!\". Use the Webdings font and the character code \"225\" (the thermometer character)."),
            new ClsQuestion(88,"Tìm từ \"automatic\" và xóa nó từ tài liệu.","Add the Trade Mark immediately after the \"WareWolf\" text on page 4."),
            new ClsQuestion(89,"Sử dụng tính năng của Word để thay thế tất cả các trường hợp \"Woodgrove Basic\" thành \"Woodgrove Plus\".","\"You work for Woodgrove Bank. You are preparing a brochure that explains U.S bank account options for international students. "),
            new ClsQuestion(90,"Sử dụng tính năng của Word để thay thế tất cả các trường hợp \"city\" thành \"community\".","Find the word \"\"automatic\"\" and delete it from the document.\""),
            new ClsQuestion(91,"Sử dụng tính năng \"Go to\" để di chuyển đến dấu trang (bookmark) tên \"Code\". Xóa đoạn văn tại vị trí đó.","Use Word's feature to replace all instances of \"Woodgrove Basic\" with \"Woodgrove Plus\"."),
            new ClsQuestion(92,"Sử dụng tính năng Go to điều hướng đến tiêu đề (heading) thứ 5 và xóa đoạn văn đầu tiên dưới nó.","Use Word's Go to feature to replace all instances of \"city\" with \"community\"."),
            new ClsQuestion(93,"Trong phần nội dung \"Savings Accounts\", xóa bình luận được đính kèm cho văn bản \"$3,000\"","Use the \"Go to\" feature to navigate to the bookmark named \"Code\". Delete the paragraph at that location."),
            new ClsQuestion(94,"Trong phần nội dung \"Evaluate and monitor the program\", giải quyết (resolve) bình luận.","Use the Go to feature to navigate to the 5th heading and delete the first paragraph below it."),
            new ClsQuestion(95,"Chấp nhận tất cả các từ được chèn vào và xóa bỏ trong tài liệu, và từ chối tất cả các thay đổi định dạng. ","In the \"Savings Accounts\" content section, delete the comment attached to the text \"$3,000\""),
            new ClsQuestion(96,"Thêm nội dung của tập tin Farmers.docx trong thư mục Documents vào cuối tài liệu.","In the \"Evaluate and monitor the program\" content section, resolve the comment."),
            new ClsQuestion(97,"Chấp nhận tất cả các dấu được chèn vào và xóa bỏ. Từ chối tất cả các thay đổi.","Accept all inserted and deleted words in the document, and reject all formatting changes."),
            new ClsQuestion(98,"Sao chép văn bản \" Blue Yonder Airlines is ….\" và dán chỉ văn bản ở cuối đoạn văn gần cái hình.","Add the contents of the Farmers.docx file in the Documents folder to the end of the document."),
            new ClsQuestion(99,"Sao chép văn bản \"Fourth Coffee\" và dán sau văn bản \"My Coffee\" để nó được in đậm và canh lề trái.","Accept all tracked insertions and deletions. Reject all formatting changes."),
            new ClsQuestion(100,"Cắt đoạn văn thứ 2 bên dưới tiêu đề \"Rehearse … \" và dán nó giữa 2 đoạn văn dưới tiêu đề \"Summarize Main Points\".","Cut the second paragraph below the \"Rehearse … \" heading and paste it between the two paragraphs below the \"Summarize Main Points\" heading."),
            new ClsQuestion(101,"Ở cuối tài liệu, xác định đoạn văn bản bắt đầu với \"If you cheat on an assignment\". Thêm một dấu trang (bookmark) tên \"Cheating\" tại đầu đoạn văn.","\"You work at Bellows College. You are helping a professor create a course syllabus."),
            new ClsQuestion(102,"Chèn liên kết (hyperlink) cho từ \"Home\" ở cuối tài liệu để liên kết đến tiêu đề \"Introduction\".","At the end of the document, locate the paragraph that begins with \"\"If you cheat on an assignment\"\". Add a bookmark called \"\"Cheating\"\" at the beginning of the paragraph.\""),
            new ClsQuestion(103,"Thêm liên kết cho văn bản \"tailspintoys.com\" ở đoạn cuối cùng để liên kết đến địa chỉ trang web \"http://www.tailspintoys.com\"","Insert a hyperlink for the word \"Home\" at the end of the document to link to the \"Introduction\" heading."),
            new ClsQuestion(104,"Trong đoạn văn trống sau tiêu đề của tài liệu, chèn một mục lục. Sử dụng kiểu Automatic Table 1 style.","Add a hyperlink for the text \"tailspintoys.com\" at the end of the paragraph to link to the website address \"http://www.tailspintoys.com\""),
            new ClsQuestion(105,"Tạo lại mục lục để hiển thị chỉ tiêu đề cấp 1.","In the blank paragraph after the document title, insert a table of contents. Use the Automatic Table 1 style."),
            new ClsQuestion(106,"Cập nhật mục lục dưới tiêu đề \"Contents\".","Recreate the table of contents to display only the level 1 headings."),
            new ClsQuestion(107,"Trong đoạn văn đầu sau tiêu đề \"Our video Collection\", chèn chú thích (footnote) \"Free to join\" sau từ \"Member\"","Update the table of contents under the \"Contents\" heading."),
            new ClsQuestion(108,"Trong phần nội dung \"Event Packages\", chèn chú thích (footnote) bên phải của tiêu đề. Nhập vào chú thích văn bản \"Includes digital files.\"","In the first paragraph after the \"Our video Collection\" heading, insert the footnote \"Free to join\" after the word \"Member\""),
            new ClsQuestion(109,"Chuyển đổi tất cả các chú thích cuối trang (endnote) thành chú thích (footnote)","In the \"Event Packages\" body, insert a footnote to the right of the heading. Enter the text \"Includes digital files.\""),
            new ClsQuestion(110,"Trong phần nội dung \"Description\" , chèn một trình giữ chỗ trích dẫn mới ( New Placeholder Citation) tên \"Manufacturing1\" ở cuối đoạn văn thứ 2 sau tiêu đề.","Convert all endnotes to footnotes"),
            new ClsQuestion(111,"Sửa đổi nguồn trích dẫn để thay đổi năm thành 2001.","In the \"Description\" body, insert a New Placeholder Citation named \"Manufacturing1\" at the end of the second paragraph after the heading."),
            new ClsQuestion(112,"Lưu một bản sao của tài liệu như mẫu Word 2019 tên \"Notes\" để tương thích với các tính năng Word mới nhất và không hỗ trợ macros. Lưu mẫu tập tin vào vị trí mặc định.","Edit the citation source to change the year to 2001."),
            new ClsQuestion(113,"Lưu một bản sao của tài liệu trong thư mục Documents kiểu Plain - Text tên \"Memo\"","\"You have been assigned the task of keeping meeting notes for your company's Board of Directors."),
            new ClsQuestion(114,"Sửa đổi tài liệu để xóa nó khỏi chế độ tương thích (Compatibility Mode)","Save a copy of the document as a Word 2019 template named \"\"Notes\"\" to be compatible with the latest Word features and not support macros. Save the template file to the default location.\""),
            new ClsQuestion(115,"Trong thuộc tính của tập tin, thêm \"animals\" vào danh mục (category)","\"This project has only one task. You create a memo in Word, and then decide to send the information as a text message."),
            new ClsQuestion(116,"Trong thuộc tính của tập tin, thêm \"bicycles\" vào danh mục (category)","Save a copy of the document in a Plain-Text Documents folder named \"\"Memo\"\"\""),
            new ClsQuestion(117,"Cấu hình word để nó tự động thay thế \"ui\" bằng \"aut\"","\"You are reviewing an event flyer for the owner of Margie's Travel."),
            new ClsQuestion(118,"Cho hiện văn bản ẩn","Modify the document to remove it from Compatibility Mode\""),
            new ClsQuestion(119,"Hiển thị các ký hiệu định dạng tab trong tài liệu. Không hiển thị các ký hiệu định dạng khác.","\"You work for Tailspin Toys. You are creating an internal product announcement and training document. "),
            new ClsQuestion(120,"Kiểm tra tài liệu và gỡ bỏ các đầu trang, chân trang và hình mờ được tìm thấy (Headers, Footers và Watermarks). Không gỡ bỏ các thông tin khác.","In the file properties, add \"\"animals\"\" to the category\""),
            new ClsQuestion(121,"Kiểm tra các vấn đề về khả năng truy cập của tài liệu. Sửa các vấn đề được báo cáo trong kết quả kiểm tra bằng cách sử dụng hành động đầu tiên được gợi ý.","\"You work for Trey Research. You are preparing a white paper summarizing research about cycling on college campuses."),
            new ClsQuestion(122,"Kiểm tra các vấn đề về khả năng truy cập của tài liệu. Sửa các vấn đề được báo cáo có liên quan đến bảng trong kết quả kiểm tra bằng cách sử dụng hành động đầu tiên được gợi ý. Không sửa các vấn đề báo cáo khác.","In the file properties, add \"\"bicycles\"\" to the category\""),
            ////TODO: Below question not yet be implemented, need to comment source code
            //new ClsQuestion(123,"Trong thuộc tính của tập tin, thêm \"program\" vào danh mục (Subject)","In the file properties, add \"program\" to the Subject"),
        };

        public static string GetEngQuestion(int index)
        {
            var obj = listQuestion.Where(x => x.Index == index).FirstOrDefault();
            if (obj == null)
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
