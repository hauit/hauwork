using Microsoft.Office.Interop.Excel;
using MOS_PPT_LEARN;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;

namespace MOS_PPT_LEARN.Base
{
    public class ClsQuestion
    {
        public bool Status { get; private set; }
        public int MaskIndex { get; private set; } //Index for display on GUI
        public int CorrectIndex { get; private set; } //Index for get file and check result
        public string EngQuestion { get; private set; }
        public string VnQuestion { get; private set; }
        public string Url { get; private set; }
        public ClsQuestion(bool status,int maskIndex, int correctIndex, string url, string vn, string eng)
        {
            this.Status = status;
            this.MaskIndex = maskIndex;
            this.CorrectIndex = correctIndex;
            this.EngQuestion = eng;
            this.VnQuestion = vn;
            this.Url = url;
        }
    }

    public static class ClsListQuestion
    {
        private static List<ClsQuestion> listQuestion = new List<ClsQuestion>
        {
            new ClsQuestion(true,23,1,"https://go.mos360.vn/mosexcel048","Trên slide 2, trong phần giữ chỗ nội dung, chèn SmarArt tên Vertical Curved List. Nhãn của hình đầu tiên là \"Structures\" và hình thứ 2 là \"Perennials\". Xóa bất cứ hình nào không sử dụng.","On slide 2, in the content placeholder, insert a Vertical Curved List SmartArt graphic. Label the first shape \"Structures\" and the second shape \"Perennials\". Delete any unused shapes."),
            new ClsQuestion(true,24,2,"https://go.mos360.vn/mosexcel048","Trên slide 3, sử dụng tính năng mô hình 3D (3D Models) chèn mô hình Fountain từ thư mục 3D Object. Thay đổi kích thước chiều cao của mô hình thành 4\" (10.16cm). Đặt mô hình nằm bên trái danh sách đánh dấu đầu dòng. Vị trí chính xác của mô hình thì không quan trọng.","On slide 3, use the 3D Models feature to insert the Fountain model from the 3D Objects folder. Resize the model to a height of 4\" (10.16 cm). Place the model to the left of the bulleted list. The exact position of the model does not matter."),
            new ClsQuestion(true,68,3,"https://go.mos360.vn/mosexcel048","Trên slide 5, sử dụng công cụ trên thẻ Draw để tô sáng văn bản “???\" với màu Yellow, 6 mm Highlighter hiển thị gần như là ???","On slide 5, use the tools on the Draw tab to highlight the text \"???\" with the Yellow, 6 mm highlighter, approximately as shown:???"),
            new ClsQuestion(true,83,4,"https://go.mos360.vn/mosexcel048","Trên slide 4, thêm chuyển động Swivel cho hình vọng lâu.","On slide 4, add the Swivel animation to the gazebo image."),
            new ClsQuestion(true,93,5,"https://go.mos360.vn/mosexcel048","Áp dụng chuyển tiếp Wipe với hiệu ứng From Left cho tất cả các slide.","Apply the Wipe transition with the From Left effect option to all slides."),
            new ClsQuestion(true,100,6,"https://go.mos360.vn/mosexcel048","Tạo section tên “Structures\" để bao gồm chỉ slide 3 và 4","Create a section named \"Structures\" that includes only slides 3 and 4."),
            new ClsQuestion(true,25,7,"https://go.mos360.vn/mosexcel048","Trên slide \"Travel Destinations\" thêm văn bản thay thế “ocean\" cho hình ảnh.","On the \"Travel Destinations\" slide, add the alt text description 'ocean\" to the image."),
            new ClsQuestion(true,43,8,"https://go.mos360.vn/mosexcel048","Trên slide \"Price Compasion\" ở nơi giữ chỗ nội dung, tạo biểu đồ 3-D Clustered Column để hiển thị chỉ nội dung của bảng. Bạn có sao chép và dán hoặc nhập thủ công dữ liệu bảng vào biểu đồ.","On the \"Price Compasion\" slide, in the content placeholder, create a 3-D Clustered Column chart that displays only the table content. You may copy and paste or manually enter the table data in the chart worksheet."),
            new ClsQuestion(true,53,9,"https://go.mos360.vn/mosexcel048","Sau slide \"Extra Activities\", tạo các slide băng cách nhập tập tin Destinations từ bên ngoài trong thư mục Documents","After the \"Extra Activities\" slide, create slides by importing the outline from the Destinations document in the Documents folder."),
            new ClsQuestion(true,118,10,"https://go.mos360.vn/mosexcel048","Trong thuộc tính của tập tin, thêm “Travel\" như là Category","You are developing a sales presentation for Margies Travel. In the file properties, add \"Travel\" as a category."),
            new ClsQuestion(true,126,11,"https://go.mos360.vn/mosexcel048","Trên slide \"Extra Activities\", chèn một dòng ở cuối bảng. Trong dòng, nhập “Sporting Event\" trong cột “Event\" và “$175\" trong cột “Price\"","On the \"Extra Activities\" slide, insert a row at the end of the table. In the row, enter \"Sporting Event\" in the \"Event\" column and \"$175\" in the \"Price\" column."),
            new ClsQuestion(true,37,12,"https://go.mos360.vn/mosexcel048","Trên slide 1, cắt (Crop) hình người đàn ông đang chạy bộ để thẳng hàng cạnh bên phải của cái hình với cạnh bên phải của slide. Không thay đổi tỷ lệ ảnh.","On slide 1, crop the image of the runner so that its right edge aligns with the right edge of the slide. Do not change the image scale."),
            new ClsQuestion(true,38,13,"https://go.mos360.vn/mosexcel048","Trên slide 5, thay đổi chế xem của mô hình 3D thành Above Front Right. Sau đó, điều chỉnh chiều cao của mô hình thành 4.5\" (11.43cm).","On slide 5, change the view of the 3D model to Above Front Right. Then, resize the model to a height of 4.5\" (11.43cm)."),
            new ClsQuestion(true,46,14,"https://go.mos360.vn/mosexcel048","Trên slide 6, ở nơi giữ chỗ nội dung, tạo biểu đồ Line with Markers để hiển thị nội dung của bảng.","On slide 6, in the content placeholder, create a Line with Markers chart that displays the table content. You made copy and paste or manually enter the table data in the chart worksheet."),
            new ClsQuestion(true,98,15,"https://go.mos360.vn/mosexcel048","Áp dụng chuyển tiếp Rotate cho slide với hiệu ứng From Left tùy chỉnh cho tất cả các slide.","Apply the Rotate slide transition with the From Left effect option to all slides."),
            new ClsQuestion(true,137,16,"https://go.mos360.vn/mosexcel048","Chèn chân trang cho slide để hiển thị số slide (Slide Number) và từ “Draft\". Áp dụng chân trang cho tất cả các slide ngoại trừ slide tiêu đề.","You are revising a fitness training presentation for potential clients of VanAsdel, Ltd. Insert a slide footer that displays the slide number and the word \"Draft\". Apply the footer to all slides except the title slide."),
            new ClsQuestion(true,39,17,"https://go.mos360.vn/mosexcel048","Trên slide \"Become a Frequent Flyer\", phóng to hình dạng đám mây nhỏ nhất chính xác với kích thước của những đám mây khác. Vị trí của đám mây thì không quan trọng.","On the \"Become a Frequent Flyer\" slide, enlarge the smallest cloud shape to exactly match the size of the other clouds. The exact location of the cloud does not matter."),
            new ClsQuestion(true,55,18,"https://go.mos360.vn/mosexcel048","Sau slide \"Frequent Flyer Points\", tạo các slide bằng cách nhập tập tin BlueCard từ bên ngoài trong thư mục Documents.","After the \"Frequent Flyer Points\" slide, create slides by importing the outline from the BlueCard document in the Documents folder."),
            new ClsQuestion(true,69,19,"https://go.mos360.vn/mosexcel048","Trên slide \"Become a Frequent Flyer, đối với văn bản “Registration is easy and fast\" hãy thay đổi mau văn bản thành màu Blue-Gray, Accent 1","On the \"Become a Frequent Flyer\" slide, for the text \"Registration is easy and fast\", change the text fill color to Blue-Gray, Accent 1"),
            new ClsQuestion(true,88,20,"https://go.mos360.vn/mosexcel048","Trên slide \"Frequent Flyer Rewards\", làm chuyển động hình máy bay để nó bay vào (Fly In) từ phía trên góc trái slide. Thiết lập thời gian chuyển động là 2 giây.","On the \"Frequent Flyer Rewards\" slide, animate the plane image to fly in from the top-left corner of the slide. Set the animation duration to 2 seconds."),
            new ClsQuestion(true,120,21,"https://go.mos360.vn/mosexcel048","Trong thuộc tính của tập tin, thiết lập tiêu đề (Title) thành \"Preferred Customer Program\"","You are creating a presentation for preferred customers of Blue Yonder Airlines. In the file properties, set the title to \"Preferred Customer Program\""),
            new ClsQuestion(true,138,22,"https://go.mos360.vn/mosexcel048","Trên slide \"Frequent Flyer Points\", xóa cột \"Average Percent Earned\". Không bắt buộc phải canh giữa bảng trên slide.","On the \"Frequent Flyer Points\" slide, delete the \"Average Percent Earned\" column. Centering the table on the slide is optional."),
            new ClsQuestion(true,28,23,"https://go.mos360.vn/mosexcel048","Trên slide 2, đảo ngược thứ tự xếp chồng của các ảnh màn hình để điện thoại thông minh ra phía trước, máy tính bảng (tablet) ở giữa và màn hình máy tính ra phía sau.","On slide 2, reverse the stacking order of the screen images so the smartphone in front, the tablet in the middle, and the monitor in back."),
            new ClsQuestion(true,29,24,"https://go.mos360.vn/mosexcel048","Trên slide 3, chuyển đổi danh sách đánh dấu đầu dòng thành sơ đồ Basic Block List.","On slide 3, convert the bulleted list to a Basic Block List SmartArt graphic."),
            new ClsQuestion(true,44,25,"https://go.mos360.vn/mosexcel048","Trên slide 4, thay đổi loại biểu đồ thành Clustered Bar.","On slide 4, change the chart type to a Clustered Bar chart."),
            new ClsQuestion(true,54,26,"https://go.mos360.vn/mosexcel048","Bạn đang chuẩn bị một bài thuyết trình về thời gian sử dụng màn hình cho Trey Research. Ẩn slide 5.","You are preparing a presentation about screen time for Trey Research. Hide slide 5."),
            new ClsQuestion(true,94,27,"https://go.mos360.vn/mosexcel048","Đối với tất cả các slide, thiết lập thời lượng chuyển tiếp là 3 giây.","For all slides, set the transition duration to 3 seconds."),
            new ClsQuestion(true,128,28,"https://go.mos360.vn/mosexcel048","Trên slide 1, chèn một liên kết từ văn bản \"Trey Research\" đến \"http://www.treyresearch.net\"","On slide 1, insert a hyperlink from the text \"Trey Research\" to \"http://www.treyresearch.net\"."),
            new ClsQuestion(true,33,29,"https://go.mos360.vn/mosexcel048","Trên slide 6, thêm văn bản thay thế \"Soccer players\" cho hình ở phía trên góc phải.","On slide 6, add the alt text description \"Soccer players\" to the upper-right image."),
            new ClsQuestion(true,34,30,"https://go.mos360.vn/mosexcel048","Trên slide 2, trong nơi giữ chỗ nội dung, chèn sơ đồ Vertical Box List chứa văn bản “Learn\", “Play\" và “Sociallize\" từ trên xuống dưới.","On slide 2, in the content placeholder, insert a Vertical Box List SmartArt graphic that contains the text \"Learn\", \"Play\", and \"Sociallize\" from top to bottom."),
            new ClsQuestion(true,35,31,"https://go.mos360.vn/mosexcel048","Trên slide 5, sử dụng tính năng mô hình 3D, chèn mô hình Piano từ thư mục 3D Object. Điều chỉnh chiều cao mô hình thành 2\" (5.08cm) và chiều rộng 1.78\" (4.15cm). Đặt mô hình vào hình chữ nhật trồng ở bên trái. Vị trí chính xác của mô hình thì không bắt buộc.","On slide 5, use the 3D Models feature to insert the Piano model from the 3D Objects folder. Resize the model to a height of 2\" (5.08 cm) and a width of 1.78\" (4.15 cm). Place the model in the empty rectangle on the left. The exact position of the model does not matter."),
            new ClsQuestion(true,86,32,"https://go.mos360.vn/mosexcel048","Trên slide 4, thiết lập hiệu ứng chuyển động của danh sách đánh dấu theo hướng From Left và thay đổi thời gian diễn ra hiệu ứng là 1.5 giây.","On slide 4, set the bulleted list animation effect direction to From Left and change the duration to 1.5 seconds."),
            new ClsQuestion(true,97,33,"https://go.mos360.vn/mosexcel048","Đối với tất cả các slide, thiết lập thời lượng chuyển tiếp là 2 giây.","For all slides, set the transition duration to 2 seconds."),
            new ClsQuestion(true,102,34,"https://go.mos360.vn/mosexcel048","Tạo section tên \"Student Clubs\" bao gồm slide 3 đến 7.","You are creating a presentation to inform Bellows College students about new student clubs. Create a section named \"Student Clubs\" that includes only slides 3 through 7."),
            new ClsQuestion(true,8,35,"https://go.mos360.vn/mosexcel048","Trên Handout Master, thay đổi bên trái đầu trang để hiển thị \"First Up Consultants\" và bên trái chân trang hiển thị \"www.firstupconsultants.com\".","You are creating a presentation for potential clients of First Up Consultants. On the Handout Master, change the left header to display. \"First Up Consultants\" and the left footer to display \"www.firstupconsultants.com\"."),
            new ClsQuestion(true,32,36,"https://go.mos360.vn/mosexcel048","Trên slide \"Specialties\", chuyển đổi danh sách đánh dấu đầu dòng thành sơ đồ Segmented Cycle","On the \"Specialties\"slide, convert the bulleted list to a Segmented Cycle SmartArt graphic."),
            new ClsQuestion(true,96,37,"https://go.mos360.vn/mosexcel048","Đối với tất cả các slide, biến đổi chuyển tiếp thành Wedge.","For all slides, set the transition variation to Wedge."),
            new ClsQuestion(true,132,38,"https://go.mos360.vn/mosexcel048","Sau slide \"First Up Consultants\" chèn slide Summary Zoom để liên kết đến các slide \"Mission\", \"Goals\", \"Specialties\" và \"Consulting Team\". Không bao gồm liên kết đến slide \"First Up Consultants\".","After the \"First Up Consultants\"slide, insert a Summary Zoom slide that links only to the \"Mission\", \"Goals\", \"Specialties\", and \"Consulting Team\" slides. Do not include a link to the \"First Up Consultants\" slide."),
            new ClsQuestion(true,133,39,"https://go.mos360.vn/mosexcel048","Chỉ trên slide \"Consulting Team\", thiết lập nền slide thành hình Hands từ thư mục Documents. Cài đặt hình nền có độ trong suốt là 75%","On the \"Consulting Team\" slide only, set the slide background to the Hands image from the Documents folder. Set the background image transparency to 75%."),
            new ClsQuestion(true,134,40,"https://go.mos360.vn/mosexcel048","Trên slide \"First Up Consultants\" chuyển đổi văn bản “www.firstupconsultants.com\" thành một liên kết. Thay đổi văn bản hiển thị thành \"Contact Us\"","On the \"First Up Consultants\" slide convert the test \"www.firstupconsultants.com\" to a hyperlink. Change the display text to \"Contact Us\"."),
            new ClsQuestion(true,7,41,"https://go.mos360.vn/mosexcel048","Trên Slide Master, nhân đôi bố cục slide \"Blank\". Tên bố cục slide mới là \"Picture with Text\". Chèn nơi giữ chỗ cho hình ở bên trái và nơi giữ chỗ cho văn bản ở bên phải.","You are preparing a presentation for propective students of the School of Fine Art. On the Slide Master, duplicate the \"Blank\" slide layout. Name the new slide layout \"Picture with Text\". Insert a picture placeholder on the left and a text placeholder on the right."),
            new ClsQuestion(true,31,42,"https://go.mos360.vn/mosexcel048","Trên slide \"Who We Are\", áp dụng kiểu Bevel Perspective Left, White và hiệu ứng nghệ thuật (Artistic Effect) tên Paint Brush cho hình.","On the 'Who We Are\" slide, apply the Bevel Perspective Left, White picture styles and the Paint Brush artistic effect to the image."),
            new ClsQuestion(true,56,43,"https://go.mos360.vn/mosexcel048","Ở cuối bản trình chiếu, chèn các slide từ bản trình chiếu Campus trong thư mục Documents. Sau khi bạn chèn slide mới, slide 6 phải là \"Discover Your Campus\" và slide 7 phải là \"Display Your Art\".","At the end of the presentation, insert the slides from the Campus presentation in the Documents folder. After you insert the slides, slide 6 should be \"Discover Your Campus\" and slide 7 should be \"Display Your Art'."),
            new ClsQuestion(true,75,44,"https://go.mos360.vn/mosexcel048","Trên slide \"Get Involved\", chèn video Sailing từ thư mục video. đặt video ở góc dưới bên phải của slide. Kích thước và vị trícủa video thì không cần chính xác.","On the \"Get Involved\" slide, insert the Sailing video from the Videos folder. Place the video in the lower-right corner of the slide. The exact size and position of the video do not matter."),
            new ClsQuestion(true,85,45,"https://go.mos360.vn/mosexcel048","Trên slide \"Admission Information\", chỉnh đường di chuyển của dấu tích để chuyển động xuống.","On the \"Admission Information\" slide, configure a Down motion path animation for the check mark icon."),
            new ClsQuestion(true,131,46,"https://go.mos360.vn/mosexcel048","Tùy chỉnh chức năng in để in Notes Pages cho tất cả các slide thành 3 bản sao. Tất cả các bản sao của trang 1 được in trước các bản sao của trang 2.","Configure the printing options to print three copies of the Notes Pages for all slides. All copies of page 1 should print before any copies of page2"),
            new ClsQuestion(true,6,47,"https://go.mos360.vn/mosexcel048","Trên bố cục \"Ingredients\" của Slide Master, thay đổi cấp đánh dấu đầu dòng đầu tiên thành hình CheckBox trong thư mục Documents.","You are revising a recipe presentation for your cooking class. On the \"Ingredients\" layout of the Slide Master, change the first level bullet to use the CheckBox image from the Documents folder."),
            new ClsQuestion(true,26,48,"https://go.mos360.vn/mosexcel048","Trên slide 4, áp dụng hiệu ứng đổ bóng (Shadow) là Inside: Top Left cho cả 2 mũi tên. Thiết lập khoảng cách (Distance) bóng là 3pt.","On slide 4, apply the Inside: Top Left shadow effect to both arrows. Set the shadow distance to 3 pt."),
            new ClsQuestion(true,27,49,"https://go.mos360.vn/mosexcel048","Nhóm (Group) 3 cái hình trên slide số 5.","On slide 5, group the three images."),
            new ClsQuestion(true,74,50,"https://go.mos360.vn/mosexcel048","Trên slide 1, chỉnh đoạn âm thanh để hiệu ứng Fade chạy trong 2 giây khi người sử dụng ấn vào biểu tượng âm thanh. Cài đặt đoạn âm thanh chỉ chơi 1 lần duy nhất nhưng vẫn tiếp qua nhiều slide.","On slide 1, configure the audio clip to fade in for 2 seconds when the user clicks the audio icon. Configure the settings so the audio clip plays only one time but continues cross multiple slides."),
            new ClsQuestion(true,119,51,"https://go.mos360.vn/mosexcel048","Gỡ bỏ thuộc tính ẩn và thông tin cá nhân từ bản trình chiếu. Không gỡ bỏ bất kỳ nội dung khác.","Remove hidden properties and personal information from the presentation. Do not remove any other content."),
            new ClsQuestion(true,127,52,"https://go.mos360.vn/mosexcel048","Thiết lập trình chiếu để người xem có thể chuyển tiếp các slide (Advance Slide) bằng tay (Manually)","Set up the slide show to require the viewer to manually advance the slides."),
            new ClsQuestion(true,36,53,"https://go.mos360.vn/mosexcel048","Trên slide 2, căn chỉnh các cạnh trên của các ảnh. Không di chuyển hình ảnh theo chiều ngang.","On slide 2, align the top edges of the images. Do not move the images horizontally."),
            new ClsQuestion(true,45,54,"https://go.mos360.vn/mosexcel048","Trên slide 5, thay đổi biểu đồ để hiển thị Data Table With Legend Keys.","On slide 5, modify the chart to display a data table with legend keys."),
            new ClsQuestion(true,76,55,"https://go.mos360.vn/mosexcel048","Trên slide 6, chèn video River từ thư mục video. Đặt video ở bên phải cái hình. Kích thước và vị trí của video thì không cần chính xác.","On slide 6, insert the River video from the Videos folder. Place the video to the right of the image. The exact size and position of the video do not matter."),
            new ClsQuestion(true,87,56,"https://go.mos360.vn/mosexcel048","Trên slide 4, tùy chỉnh đường di chuyển cho hình ngôi sao chuyển động theo hình trái tim.","On slide 4, configure a heart motion path animation for the star shape."),
            new ClsQuestion(true,135,57,"https://go.mos360.vn/mosexcel048","Tùy chỉnh chức năng in để in thành 5 bản sao của Notes Pages cho bản trình chiếu theo hướng giấy dọc (Portrait Orientation). Tất cả bản sao của trang đầu tiên nên được in trước bản sao của trang thứ 2.","You are creating a presentation for potential clients of Fabrikam Residences. Configure the printing options to print five coppies of the Notes Pages for the presentation in Portrait Orientation. All copies of page 1 should print before any copies of page 2."),
            new ClsQuestion(true,136,58,"https://go.mos360.vn/mosexcel048","Trên slide 1, chèn một Slide Zoom để liên kết đến slide \"Customer Satisfaction Rating\". Vị trí hình thu nhỏ slide nằm phía dưới góc phải slide. Kích thước và vị trí của hình thu nhỏ thì không bắt buộc.","On slide 1, insert a Slide Zoom link to the \"Customer Satisfaction Rating\" slide. Place the slide zoom thumbnail in the lower-right corner of the slide. The exact size and position of the thumbnail do not matter."),
            new ClsQuestion(true,30,59,"https://go.mos360.vn/mosexcel048","Trên slide 3, thay đổi chế độ xem của mô hình 3D (3D Model Views) thành Left.","On slide 3, change the view of the 3D model to Left."),
            new ClsQuestion(true,84,60,"https://go.mos360.vn/mosexcel048","Trên slide 3, áp dụng hiệu ứng chuyển động Jump & Turn cho mô hình 3D.","On slide 3, apply the Jump & Turn animation efect to the 3D model."),
            new ClsQuestion(true,95,61,"https://go.mos360.vn/mosexcel048","Đối với tất cả các slide, biến đổi chuyển tiếp thành From Left.","For all slides, are the transition variation to From Left."),
            new ClsQuestion(true,101,62,"https://go.mos360.vn/mosexcel048","Trên slide 2, chèn vào phần liên kết thu phóng để “Section 2: Products and Services\", \"Section 3: Clubs and Teams\" và \"Section 4: Contact Us\". Đặt lại vị trí để các hình thu nhỏ nằm trong khung hình chữ nhật màu đen và để chúng không chồng lên nhau.","On slide 2, insert Section Zoom links to \"Section 2: Products and Services\", \"Section 3: Clubs and Teams\", and \"Section 4: Contact Us\". Reposition the section thumbnails within the black rectangle so they are not stacked on top of each other."),
            new ClsQuestion(true,129,63,"https://go.mos360.vn/mosexcel048","Trên slide 2, chèn vào chân trang của slide để hiển thị số slide (slide number) và \"www.advanture-works.com\". Áp dụng chân trang cho tất cả các slide ngoài trừ slide tiêu đề.","You are completing a presentation about Adventure Works Cycles. On slide 2, insert a slide footer that displays the slide number and \"www.adventure-works.com\". Apply the footer to all slides except the title slide."),
            new ClsQuestion(true,130,64,"https://go.mos360.vn/mosexcel048","Trên slide 8, định dạng danh sách đánh dấu hiển thị thành 2 cột.","On slide 8, format the bulleted list to display in two columns."),
            new ClsQuestion(true,161,65,"https://go.mos360.vn/mosexcel048","Trên slide 3, thêm bảng 3 cột và 4 dòng năm bên phải hình ngôi sao.","On slide 3, add a table with three columns and four rows to the right of the star shape."),
            new ClsQuestion(true,162,66,"https://go.mos360.vn/mosexcel048","Trên slide 8, thay đổi kiểu biểu đồ thành Style 11 và màu biểu đồ thành Colorful Palette 4 (Palette 4 trong phần Colorful).","On slide 8, change the chart style to Style II and the chart colors to Colorful Palette 4 (Palette 4 in the Colorful section)."),
            new ClsQuestion(true,163,67,"https://go.mos360.vn/mosexcel048","Tùy chỉnh tính năng in để in 4 bản sao của bản trình chiếu với 3 slide mỗi trang. Tất cả bản sao của trang đầu tiên nên được in trước bản sao của trang thứ 2.","Configure the printing options to print four copies of the presentation with three slides per page. All copies of the first page should print before the copies of the second page."),
            new ClsQuestion(true,164,68,"https://go.mos360.vn/mosexcel048","Áp dụng màu nền Gradient mặc định cho slide 6.","Apply a default Gradient fill background to slide 6."),
            new ClsQuestion(true,145,69,"https://go.mos360.vn/mosexcel048","Thêm tất cả các slide từ vineyard.pptx trong thư mục Documents vào cuối bản trình chiếu theo thứ tự.","Add all of the slides from vineyard.pptx in the Documents folder to the end of the presentation in order."),
            new ClsQuestion(true,147,70,"https://go.mos360.vn/mosexcel048","Áp dụng chủ đề (Theme) Ion Boardroom cho Slide Master.","Apply the lon Boardroom theme to the Slide Master."),
            new ClsQuestion(true,149,71,"https://go.mos360.vn/mosexcel048","Trên slide 3, thay đổi màu SmartArt thành Colorful Range - Accent Color 4 to 5.","On slide 3, change the SmartArt graphic colors to Colorful Range -Accent Colors 4 to 5."),
            new ClsQuestion(true,150,72,"https://go.mos360.vn/mosexcel048","Trên slide 5, làm chuyển động danh sách đánh dấu để mục đánh dấu đầu tiên chuyển động Wipe hướng từ phía trên (From Top) ngay khi slide hiển thị. Và mỗi mục đánh dấu tiếp theo cũng chuyển động Wipes từ phía trên sau 1 giây so với mục đánh dấu trước đó. Giữ mặc định thời lượng chuyển động.","On slide 5, animate the bulleted list so that the first list item wipes from the top immediately when the slide is shown and each subsequent list item wipes from the top one second after the"),
            new ClsQuestion(true,165,73,"https://go.mos360.vn/mosexcel048","Thêm biểu đồ Pareto Histogram mặc định cho slide 7.","Add a default Pareto Histogram chart to slide 7."),
            new ClsQuestion(true,166,74,"https://go.mos360.vn/mosexcel048","Trên slide 4, thêm bảng từ tập tin Revenue.xlsx trong thư mục Documents.","On slide 4, add the table from the Revenue.xlsx file in the Documents folder."),
            new ClsQuestion(true,167,75,"https://go.mos360.vn/mosexcel048","Trên slide 2, áp dụng Subtle Effect - Teal, Accent 5 cho hộp văn bản. Thay đổi đường viền thành 3 pt và áp dụng Angle Bevel.","On slide 2, apply the Subtle Effect - Teal, Accent 5 style to the text box. Change its outline to 3 pt and apply an Angle Bevel."),
            new ClsQuestion(true,168,76,"https://go.mos360.vn/mosexcel048","Thêm các slide dựa trên tập tin Outline.docx từ bên ngoài trong thư mục Documents ở cuối bản trình chiếu.","Add slides based on the outline from Outline.docx in the Documents folder to the end of the presentation."),
            new ClsQuestion(true,3,77,"https://go.mos360.vn/mosexcel048","Tạo mới một bố cục cho slide tên “Trevorslayout\" để giữ chỗ cho ảnh ở bên trái và giữ chỗ cho văn bản ở bên phải. Giữ mặc định tất cả các giữ chỗ. Kích thước và vị trí của các giữ chỗ mới thì không quan trọng.","Create a new slide layout called \"Trevorslayout\" with a picture placeholder on the left, and text placeholder on the right. Keep all default placeholders. The size and exact position of the new placeholders."),
            new ClsQuestion(true,4,78,"https://go.mos360.vn/mosexcel048","Thay đổi Handout Master để ghi \"First Copy\" bên trái chân trang.","Change the Handout Master have the left footer read \"First Copy\"."),
            new ClsQuestion(true,5,79,"https://go.mos360.vn/mosexcel048","Thêm một nơi giữ media cho bố cục Media. Vị trí nằm gần tiêu đề. Thẳng hàng lề trái và lề phải so với lề của tiêu đề.","Add a media placeholder to the \"Media\" layout. Position it beneath the title placeholder. Align the left and right margins to the margins of the title placeholder"),
            new ClsQuestion(true,18,80,"https://go.mos360.vn/mosexcel048","Trên slide 5, áp dụng kiểu Snip Diagonal Corner, White và hiệu ứng Pastels Smooth cho hình.","On slide 5, apply the Snip Diagonal Corner, White style and the Pastels Smooth effect to the image."),
            new ClsQuestion(true,49,81,"https://go.mos360.vn/mosexcel048","Thêm một slide mới năm giữa slide 2 và slide 3 từ tập tin Inclusive Practices.docx","Add a new slide in between slides 2 and 3 from the Inclusive Practices.docx"),
            new ClsQuestion(true,50,82,"https://go.mos360.vn/mosexcel048","Ở cuối bản trình chiếu, thêm slide từ bên ngoài có tên Niagra Facts.docx","At the end of the presentation, add the slide from the outline that is named Niagra Facts.docx"),
            new ClsQuestion(true,16,83,"https://go.mos360.vn/mosexcel048","Sắp xếp các hình trên slide 2 để chúng được căn chỉnh về điểm giữa (Align Middle).","Arrange the images on slide 2 so the middles of the pictures are aligned."),
            new ClsQuestion(true,80,83,"https://go.mos360.vn/mosexcel048","Sắp xếp lại thứ tự các hoạt ảnh của các hình trên slide 2 để chúng mờ dần (Fade) từng cái một từ trái qua phải.","Reorder the animations of the images on slide 2 so that they fade in one-by-one from left to right."),
            new ClsQuestion(true,111,84,"https://go.mos360.vn/mosexcel048","Tùy chỉnh bản trình chiếu chỉ in section \"Course Introduction\"","Configure presentation to only print the \"Course Introduction\" Section."),
            new ClsQuestion(true,112,85,"https://go.mos360.vn/mosexcel048","Thay đổi kích thước bản trình chiếu thành 8 inches (20.32 cm) chiều rộng và 11 inches (27.94 cm) chiều cao. Đảm bảo rằng phù hợp với nội dung","Change the size of the presentation to 8 inches (20.32 cm) wide and 11 inches (27.94) high. Ensure that the contents fits."),
            new ClsQuestion(true,19,86,"https://go.mos360.vn/mosexcel048","Trên slide 3, di chuyển cái hình “Mint Chocolate\" ra phía trước.","On slide 3, move the \"Mint Chocolate\" shape to the front."),
            new ClsQuestion(true,20,87,"https://go.mos360.vn/mosexcel048","Trên slide 1, áp dụng kiểu WordArt tên Fill: White; Outline: Orange, Accent colour 2; Hard Shadow: Orange, Accent colour 2 cho văn bản \"Himalayan Coffee House\".","On slide 1, apply WordArt style fill; white; Outline; Orange, Accent colour 2; Hard Shadow: Orange, Accent colour 2 to the text \"Himalayan Coffee House\""),
            new ClsQuestion(true,115,88,"https://go.mos360.vn/mosexcel048","Di chuyển slide Try our new flavours\" năm giữa slide tiêu đề và slide Menu","Move the \"Try our new flavours\" slide in between the title and the Menu slides."),
            new ClsQuestion(true,1,89,"https://go.mos360.vn/mosexcel048","Tạo mới một bố cục cho slide tên \"Customl\" để giữ chổ cho ảnh ở bên trái và giữ chỗ cho văn bản ở bên phải. Giữ mặc định tất cả các giữ chỗ. Kích thước và vị trí của các giữ chỗ mới thì không quan trọng.","Create a new slide layout named \"Custom1\" with a picture placeholder on the left and a text placeholder on thight. Keep all default placeholders. Size and position of the new placeholders do not matter."),
            new ClsQuestion(true,10,90,"https://go.mos360.vn/mosexcel048","Sắp xếp các hình trên slide 2 để chúng được căn chỉnh về điểm giữa (Align Middle).","Arrange the images on slide 2 so that their middles are aligned"),
            new ClsQuestion(true,11,91,"https://go.mos360.vn/mosexcel048","Trên slide 3, thay đổi màu của xe hơi thành màu Blue và màu đường viền là Yellow.","On slide 3, change the color of the car icon to Blue and add a Yellow outline."),
            new ClsQuestion(true,77,92,"https://go.mos360.vn/mosexcel048","Sắp xếp lại thứ tự hiệu ứng của các hình trên slide 2 để chúng mờ dần (Fade) từng cái một từ trái qua phải.","Reorder the animation of the images on slide 2 so that they fade in one by one from left to right."),
            new ClsQuestion(true,78,93,"https://go.mos360.vn/mosexcel048","Trên slide 3, làm cho biểu tượng xe hơi chuyển động để nó bay vào từ bên phải.","On slide 3, animate the car icon so that it flies in from the right."),
            new ClsQuestion(true,103,94,"https://go.mos360.vn/mosexcel048","Lưu bản trình chiếu vào thư mục Documents thành tập tin PDF tên “Presentation\"","Save the presentation to the Documents folder as a PDF file named \"Presentation\"."),
            new ClsQuestion(true,104,95,"https://go.mos360.vn/mosexcel048","Tùy chỉnh cài đặt in để chỉ in section \"Introduction\"","Configure printing to print only the \"Introduction\" section."),
            new ClsQuestion(true,15,96,"https://go.mos360.vn/mosexcel048","Trên slide 5, thêm Pyramid List chứa văn bản \"Gold\", \"Silver\" và \"Bronze\" từ trên xuống dưới. Áp dụng kiểu Inset. Thay đổi kích thước sơ đồ kim tự tháp thì không bắt buộc.","On slide 5, add a Pyramid List that contains the text \"Gold\", \"Silver\", and \"Bronze\" from top to bottom. Apply the Inset style. Resizing the pyramid is optional."),
            new ClsQuestion(true,41,97,"https://go.mos360.vn/mosexcel048","Trên slide 4, thay đổi biểu đồ để nhãn các danh mục được liệt kê ở giữa phía trên của biểu đồ. Các nhãn nằm chồng lên biểu đồ.","On slide 4, modify the chart so that the category labels are listed across the top center of the chart. The labels should overlap the chart."),
            new ClsQuestion(true,63,98,"https://go.mos360.vn/mosexcel048","Trên slide 3, xóa dòng \"Sinusitis\" từ bảng sau đó chèn một cột mới ở bên phải với tiêu đề “Percentage Uninsured\".","On slide 3, delete the \"Sinusitis\" row from the table then insert a new column titled \"Percentage Uninsured\" on the right."),
            new ClsQuestion(true,109,99,"https://go.mos360.vn/mosexcel048","Tùy chỉnh cài đặt in để in Notes Pages cho tất cả các slide.","Configure the printing options to print the Notes Pages for all slides."),
            new ClsQuestion(true,12,100,"https://go.mos360.vn/mosexcel048","Mang văn bản trên slide 4 lên trước hình cánh tay. Sau đó gửi hình gia đình ra phía sau.","Bring the text on slide 4 in front of the image of the hands. Then, send the family image to the back."),
            new ClsQuestion(true,106,101,"https://go.mos360.vn/mosexcel048","Thay đổi chủ đề chính thành Office Theme và thay đổi phông chữ thành Arial","Change the master design theme to Office Theme and change the font to Arial"),
            new ClsQuestion(true,9,102,"https://go.mos360.vn/mosexcel048","Trên slide \"Achievements\", áp dụng hiệu ứng Angle Bevel cho 6 cái hình.","On the \"Achievements\" slide, apply the Angle Bevel effect to all six photographs."),
            new ClsQuestion(true,47,103,"https://go.mos360.vn/mosexcel048","Sau slide \"Achievements\", nhập các slide mới từ tài liệu Word bên ngoài có tên Presenter Order.docx trong thư mục Documents.","After the \"Achievements\" slide, import new slides from the Word document outline titled Presenter Order.docx in the Documents folder."),
            new ClsQuestion(true,2,104,"https://go.mos360.vn/mosexcel048","Thay đổi danh sách đánh dấu đầu dòng cao nhất của Slide Master sử dụng tập tin Check.png trong thư mục Picture.","Change the top-level bullet of the Slide Master to use the Check.png file in the Pictures folder."),
            new ClsQuestion(true,61,105,"https://go.mos360.vn/mosexcel048","Trên slide 3, thêm liên kết đến website \"http://www.humongousinsurance.com\" cho câu \"Click here to view on website\".","On slide 3, add a hyperlink to the website \"http://www.humongousinsurance.com\" to the sentence \"Click here to view on website\"."),
            new ClsQuestion(true,62,106,"https://go.mos360.vn/mosexcel048","Chèn một bình luận ghi nội dung “Update\" cho biểu đồ trên slide 2.","Insert a comment that reads \"Update\" on the chart on slide 2."),
            new ClsQuestion(true,48,107,"https://go.mos360.vn/mosexcel048","Thêm một slide mới vào cuối bản trình chiếu sử dụng tập tin Vanessa.docx trong thư mục Documents.","Sau slide \"Achievements\", nhập các slide mới từ tài liệu Word bên ngoài có tên Presenter Order.docx trong thư mục Documents."),
            new ClsQuestion(true,71,108,"https://go.mos360.vn/mosexcel048","Chỉnh video trên slide 2 để bắt đầu \"00:00.500\" và kết thúc tại \"00:02.500\"","Configure the video on slide 2 to start at \"00:00.500\" and end at \"00:02.500\"."),
            new ClsQuestion(true,79,109,"https://go.mos360.vn/mosexcel048","Làm cho văn bản trên trang chiếu 3 chuyển động để mỗi dấu đầu dòng di chuyển kiểu Wipe từ trái sang phải riêng lẻ khi nhấp vào.","Animate the text on slide 3 so that each bullet wipes in from left to right individually on click."),
            new ClsQuestion(true,108,110,"https://go.mos360.vn/mosexcel048","Kiểm tra và gỡ bỏ các chú thích và nội dung nằm bên ngoài slide.","Inspect for and remove annotations and content that is positioned off the slide."),
            new ClsQuestion(true,40,111,"https://go.mos360.vn/mosexcel048","Tạo biểu đồ đường Line trên slide 7 sử dụng dữ liệu bảng được cung cấp trên cùng slide. Sử dụng các năm như Categories và \"New Customers\" như Series. Thay đổi kích thước biểu đồ là không bắt buộc.","Create a Line chart on slide 7 using the figures provided in the table on the same slide. Use the years as the Categories and \"New Customers\" as the Series. Resizing the chart is optional."),
            new ClsQuestion(true,59,112,"https://go.mos360.vn/mosexcel048","Trên slide 5, thêm vào chân trang với văn bản “Company Confidential\".","On slide 5 only, add a footer with the text \"Company Confidential\"."),
            new ClsQuestion(true,70,113,"https://go.mos360.vn/mosexcel048","Trên slide 4, thêm video New Advert.avi từ thư mục Videos. Vị trí của nó tại 2\" (5.08 cm) từ Top Left Corner theo chiều ngang và chiều dọc.","On slide 4, add the video New Advert.avi from the Videos folder. Position it at 2\" (5.08 cm) from the Top Left Corner, vertically and horizontally."),
            new ClsQuestion(true,66,114,"https://go.mos360.vn/mosexcel048","Trên slide 4, chuyển đổi danh sách thành sơ đồ Vertical Picture Accent List.","On slide 4, convert the list into a Vertical Picture Accent List SmartArt graphic."),
            new ClsQuestion(true,72,115,"https://go.mos360.vn/mosexcel048","Trên slide 3, chỉnh đoạn âm thanh để nó chạy tự động. Ấn biểu tượng trong suốt quá trình trình chiếu.","On Slide 3, configure the audio clip so that it plays automatically. Hide the icon during the show."),
            new ClsQuestion(true,73,116,"https://go.mos360.vn/mosexcel048","Trên slide 3, cắt (crop) video để vùng hiển thị là 6.5\" (16.51 cm) từ lề trái và cắt chiều rộng thành 4.4\" (11.18 cm)","On slide 3, crop the video so that the area shown is 6.5\" (16.51 cm) from the left margin and also crop the width to 4.4\" (11.18 cm)"),
            new ClsQuestion(true,82,117,"https://go.mos360.vn/mosexcel048","Trên slide 2, áp dụng đường di chuyển (Motion Path) hình tròn cho hình có nội dung “No Way\"","On slide 2, apply the circle shape motion path to the shape that says \"no way\"."),
            new ClsQuestion(true,92,118,"https://go.mos360.vn/mosexcel048","Đối với tất cả các chuyển tiếp của các slide, thiết lập thời lượng chuyển tiếp là 3 giây và âm thanh Breeze.","For all slide transitions, set the duration to 3 seconds and the sound to Breeze."),
            new ClsQuestion(true,42,119,"https://go.mos360.vn/mosexcel048","Thay đổi biểu đồ trên slide 6 thành biểu đồ 3-D Clustered.","Change the chart on slide 6 to a 3-D Clustered Column chart"),
            new ClsQuestion(true,64,120,"https://go.mos360.vn/mosexcel048","Trên slide 7, áp dụng Medium Style 2, Accent 1 cho bảng. Thay đổi bảng để các cột thay đổi màu và các dòng thì không.","On slide 7, apply the medium style 2, accent 1 to the table. Modify the table so that the columns alternate colour and the rows do not."),
            new ClsQuestion(true,113,121,"https://go.mos360.vn/mosexcel048","Tạo một tùy chỉnh trình chiếu slide tên \"Important Findings\" để chỉ bao gồm slide 3 và 5","Create a custom slide show named \"Important Findings\" that includes slides 3 and 5 only."),
            new ClsQuestion(true,114,122,"https://go.mos360.vn/mosexcel048","Nhúng các phong chữ chỉ cho các ký tự được sử dụng trong bản trình chiếu này. Lưu bản trình chiếu.","Embed the fonts for only the characters used in this presentation. Save the presentation."),
            new ClsQuestion(true,116,123,"https://go.mos360.vn/mosexcel048","Ẩn các bình luận trong bản trình chiếu.","Hide comments in the presentation"),
            new ClsQuestion(true,117,124,"https://go.mos360.vn/mosexcel048","Tùy chỉnh cài đặt in để in thành 3 bản sao của bản trình chiếu với 3 slide mỗi trang. Tất cả bản sao của trang đầu tiên nên được in trước bản sao của trang thứ 2.","Configure the printing options to print 3 copies of the presentation with 3 slides per page. All copies of the first page should print before the second"),
            new ClsQuestion(true,22,125,"https://go.mos360.vn/mosexcel048","Trên slide 3, nhóm (Group) tất cả hình ảnh.","On slide 3, group all of the pictures."),
            new ClsQuestion(true,52,126,"https://go.mos360.vn/mosexcel048","Xóa slide 5, \"Our Sports Products\".","Delete slide 5, \"Our Sports Products\""),
            new ClsQuestion(true,67,127,"https://go.mos360.vn/mosexcel048","Trên slide 2, canh lề văn bản “You miss 100 percent off the shots you don't take\" lên phía trên của hộp văn bản và thêm hiệu ứng Small Caps.","On slide 2, align the text \"You miss 100 percent of the shots you don't take\" to the top of the text box and add the Small Caps effect."),
            ////TODO: Below question not yet be implemented, need to comment source code
        };

        public static string GetEngQuestion(int index)
        {
            var obj = listQuestion.Where(x => x.CorrectIndex == index).FirstOrDefault();
            if (obj == null)
            {
                return string.Empty;
            }
            return obj.EngQuestion;
        }

        public static string GetVNQuestion(int index)
        {
            var obj = listQuestion.Where(x => x.CorrectIndex == index).FirstOrDefault();
            if (obj == null)
            {
                return string.Empty;
            }
            return obj.VnQuestion;
        }

        public static int GetCorrectIndex(int maskIndex)
        {
            var obj = listQuestion.Where(x => x.MaskIndex == maskIndex).FirstOrDefault();
            if (obj == null)
            {
                return 0;
            }
            return obj.CorrectIndex;
        }

        public static ClsQuestion GetQuestion(int maskIndex)
        {
            var obj = listQuestion.Where(x => x.MaskIndex == maskIndex).FirstOrDefault();
            if (obj == null)
            {
                throw new ArgumentException("Số câu hỏi không hợp lệ");
            }
            return obj;
        }

        public static int GetQuestionNumber()
        {
            return listQuestion.Count;
        }

    }

}
