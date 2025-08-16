using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using MOS_EXCEL_LEARN;
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
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;

namespace WindowsFormsApplication1.Base
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
            new ClsQuestion(true,48,1,"youtube.com","Thêm một trang tính mới tên \"Microsoft Word\" cho tập tin này.","Add a new worksheet named \"Microsoft Word\" to this file."),
            new ClsQuestion(true,31,2,"youtube.com","Đổi tên trang tính \"Outdoor toys\" thành \"Outdoor sports\".","Rename the \"Outdoor toys\" worksheet to \"Outdoor sports\"."),
            new ClsQuestion(true,9,3,"youtube.com","Di chuyển trang tính \"Donor List\" sang bên phải những trang tính khác.","Move the \"Donor List\" worksheet to the right of the other worksheets."),
            new ClsQuestion(true,67,4,"youtube.com","Tạo một bảng sao chép trang tính tên \"October\" để bên phải trang tính \"October\".","Create a copy of the \"October\" worksheet to the right of the \"October\" worksheet."),
            new ClsQuestion(true,28,5,"youtube.com","Làm cho trang tính \"Outdoor toys\" để nó không hiển thị nhưng chúng ta có thể tham chiếu từ những trang tính khác.","Configure the \"Outdoor toys\" worksheet so that it doesn't show but we can refer from other worksheet."),
            new ClsQuestion(true,29,6,"youtube.com","Bỏ ẩn tất cả các trang tính trong tập tin này.","Unhide all the worksheets in this file."),
            new ClsQuestion(true,30,7,"youtube.com","Thay đổi màu nhãn (thẻ), của trang tính \"Outdoor toys\" thành màu \"Blue, Accent 1\".","Change the label color of the \"Outdoor toys\" worksheet to \"Blue, Accent 1\"."),
            new ClsQuestion(true,1,8,"youtube.com","Trên trang tính \"Instructional Hours\", điều chỉnh độ rộng của cột B:G chính xác là 12.","On the \"Instructional Hours\" worksheet, adjust the width of columns B:G to exactly 12."),
            new ClsQuestion(true,32,9,"youtube.com","Trên trang tính \"Shareholders Info\", thay đổi chiều cao của cột chứa văn bản \"Annual Report\" thành 30.","On the \"Shareholders Info\" worksheet, change the height of the column containing the text \"Annual Report\" to 30."),
            new ClsQuestion(true,10,10,"youtube.com","Trên trang tính \"Donor Contact Info\", thêm cột tên \"Cell Phone\" để bên phải cột \"Home Phone\".","On the \"Donor Contact Info\" worksheet, add a column named \"Cell Phone\" to the right of the \"Home Phone\" column."),
            new ClsQuestion(true,85,11,"youtube.com","Xóa cột D trong trang tính \"Students\".","Delete column D in the \"Students\" worksheet."),
            new ClsQuestion(true,86,12,"youtube.com","Trên trang tính \"roster\" không xóa dòng 1 đến dòng 6 nhưng không cho chúng hiển thị.","On the \"roster\" worksheet, do not delete rows 1 through 6, but do not make them visible."),
            new ClsQuestion(true,87,13,"youtube.com","Trên trang tính roster bỏ ẩn dòng 1 đến dòng 6.","On the roster worksheet, unhide rows 1 through 6."),
            new ClsQuestion(true,71,14,"youtube.com","Trên trang tính \"Product\", giữ các dòng 1 và 2 luôn duy trì giữ lại trong khi cuộn.","You are preparing inventory data to present to your manager at Lamna Healthcare Company. On the \"Product\" worksheet, keep rows 1 and 2 always retained while scrolling."),
            new ClsQuestion(true,17,15,"youtube.com","Trên trang tính \"Last Semester\" , điều chỉnh vùng A3:F3 để văn bản tự động xuống hàng.","You are compiling data about courses at Bellows College. On the \"Last Semester\" worksheet, adjust the range A3:F3 so that the text automatically wraps to a new line."),
            new ClsQuestion(true,61,16,"youtube.com","Trên trang tính \"New Policies\", thay đổi căn chỉnh nằm ngang tên \"Agent\" thành \"Left (Indent),. Thiết lập thụt lề là 1.","You are analyzing insurance sales data for the first six months of the year. On the \"New Policies\" worksheet, change the horizontal alignment of the \"Agent\" name to \"Left (Indent),. Set the indent to 1."),
            new ClsQuestion(true,72,17,"youtube.com","Trên trang tính \"Products\", căn trái văn bản ở ô A1.","On the \"Products\" worksheet, left-align the text in cell A1."),
            new ClsQuestion(true,55,18,"youtube.com","Trên trang tính \"Materials\", điều chỉnh chiều rộng tất cả cột A:N để tự động làm cho vừa với mục nhập vào lớn nhất.","On the \"Materials\" worksheet, adjust the width of all columns A:N to automatically fit the largest entry."),
            new ClsQuestion(true,11,19,"youtube.com","Sao chép vùng A59:C63 trên trang tính \"Donor List\" đến vùng A2:C6 trên trang tính \"Top Donors\".","Copy the range A59:C63 on the \"Donor List\" worksheet to the range A2:C6 on the \"Top Donors\" worksheet."),
            new ClsQuestion(true,12,20,"youtube.com","Sao chép vùng A9:A11 trên trang tính \"Donor List\" đến vùng A2:C2 trên trang tính \"Top Donors\".","Copy the range A9:A11 on the \"Donor List\" worksheet to the range A2:C2 on the \"Top Donors\" worksheet."),
            new ClsQuestion(true,13,21,"youtube.com","Trong cột \"Zip Code\" trên trang tính \"Donor Info\", tìm tất cả giá trị \"64044\" thay thế thành \"64043\".","In the \"Zip Code\" column on the \"Donor Info\" worksheet, find all values ​​of \"64044\" and replace them with \"64043\"."),
            new ClsQuestion(true,78,22,"youtube.com","Trên trang tính \"Quote\", dời các ô E8:F32 lên cho khớp với dữ liệu bằng cách xóa các ô E7:F7.","You are updating the sales spreadsheet for an online clothing store. On the \"Quote\" worksheet, move cells E8:F32 up to match the data by deleting cells E7:F7."),
            new ClsQuestion(true,79,23,"youtube.com","Trên trang tính \"Quote\" , dời các ô E6:F20 xuống cho khớp với bảng dữ liệu bằng cách chèn các ô E6:F6","On the \"Quote\" worksheet, move cells E6:F20 down to match the data table by inserting cells E6:F6"),
            new ClsQuestion(true,7,24,"youtube.com","Trên trang tính \"Customers by Order\", sử dụng công cụ dữ liệu excel để gỡ bỏ tất cả những bản ghi (dòng), có giá trị \"CustomerID\" trùng lặp khỏi bảng. Không gỡ bỏ những bản ghi khác.","On the \"Customers by Order\" worksheet, use the Excel data tool to remove all records (rows), with duplicate \"CustomerID\" values ​​from the table. Do not remove any other records."),
            new ClsQuestion(true,2,25,"youtube.com","Trên trang tính \"Substitutes\", bắt đầu tại ô A1, thêm dữ liệu từ file \"Substitutes\" trong thư mục \"Documents\". Dùng dòng đầu tiên của dữ liệu như tiêu đề.","You are compiling class data to present to the enrollment director at the School of Fine Art. On the \"Substitutes\" worksheet, starting in cell A1, add the data from the \"Substitutes\" file in the \"Documents\" folder. Use the first row of data as the header."),
            new ClsQuestion(true,92,26,"youtube.com","Nhập nội dung của tập tin \"FBClickthroughs.csv\" trong thư mục \"Document\" vào trang tính \"Social Media Ads\" bắt đầu từ ô A17. Chọn dấu phẩy làm dấu phân cách.( Chấp nhận những mặc định khác.),","Import the contents of the \"FBClickthroughs.csv\" file in the \"Documents\" folder into the \"Social Media Ads\" worksheet starting in cell A17. Select a comma as the separator. (Accept the other defaults.),"),
            new ClsQuestion(true,70,27,"youtube.com","Nhập nội dung của tập tin \"top25.txt\" trong thư mục \"Document\" vào trang tính \"Top 25\" bắt đầu từ ô A2. Chọn dấu tab làm dấu phân cách (Chấp nhận những mặc định khác),.","Import the contents of the \"top25.txt\" file in the \"Documents\" folder into the \"Top 25\" worksheet starting in cell A2. Select a tab as the separator. (Accept the other defaults.),"),
            new ClsQuestion(true,14,28,"youtube.com","Tạo dữ liệu tổng phụ bằng cách đếm tên mỗi màu thay đổi trong trang tính \"Donor List\", tổng số phụ phải có dòng tóm tắt bên dưới mỗi nhóm và ngắt trang giữa các nhóm (page break between groups),.","Create subtotals by counting each color change in the \"Donor List\" worksheet, the subtotals should have a summary row below each group and a page break between groups."),
            new ClsQuestion(true,42,29,"youtube.com","Điều hướng đến vùng dữ liệu tên \"Rate\" và xóa bỏ nội dung của toàn bộ các ô đã được chọn.","You work for Woodgrove Bank. You are preparing information for a meeting. Navigate to the data range named \"Rate\" and delete the contents of all selected cells."),
            new ClsQuestion(true,15,30,"youtube.com","Đi đến dãy tên \"DonorLevel\" và xóa nội dung bên trong nó.","Go to the range named \"DonorLevel\" and delete the contents within it."),
            new ClsQuestion(true,16,31,"youtube.com","Tìm \"Table4\" và cập nhật dữ liệu dòng 5 cột 3 thành \"300\".","Find \"Table4\" and update row 5 column 3 data to \"300\"."),
            new ClsQuestion(true,43,32,"youtube.com","Trên trang tính \"Exchange Rates\", trong vùng gồm ô B4:D8, định dạng các ô để hiển thị số có 2 chữ số phần thập phân.","On the \"Exchange Rates\" worksheet, in the range B4:D8, format the cells to display numbers with 2 decimal places."),
            new ClsQuestion(true,80,33,"youtube.com","Trên trang tính \"Prices\", dùng kiểu \"Title\" cho ô A1.","On the \"Prices\" worksheet, use the \"Title\" style for cell A1."),
            new ClsQuestion(true,105,34,"youtube.com","Sao chép định dạng của tiêu đề và phụ đề của trang \"Tasks\" và gắn nó vào tiêu đề và phụ đề của trang \"Projects\".","You are completing the grade spreadsheet for the current semester at Bellows College. Copy the title and subtitle formatting from the \"Tasks\" worksheet and paste it into the title and subtitle of the \"Projects\" worksheet."),
            new ClsQuestion(true,56,35,"youtube.com","Trên trang tính \"Materials\", thay đổi những ô thuộc dãy A1:N1 vào một ô không có thay đổi lề của văn bản.","On the \"Materials\" worksheet, change the cells in the range A1:N1 to a cell with no text margins."),
            new ClsQuestion(true,33,36,"youtube.com","Trên trang tính \"Games\", gộp dòng 12 đến dòng 18 trong cột \"Category\" và \"Column1\" của dãy \"Costs\" thành một cột 7 dòng tên \"Category\". Những giá trị còn lại nên thẳng hàng bên trái.","On the \"Games\" worksheet, merge rows 12 through 18 in the \"Category\" column and \"Column1\" in the \"Costs\" range into a 7-row column named \"Category\". The remaining values ​​should align to the left."),
            new ClsQuestion(true,50,37,"youtube.com","Trên trang \"London\", mở rộng công thức ô E5 đến cuối cột của bảng.","You are collecting information about the distances between popular travel destinations. On the \"London\" worksheet, extend the formula in cell E5 to the end of the table column."),
            new ClsQuestion(true,23,38,"youtube.com","Lợi nhuận bằng giá bán trừ giá mua. Trên trang tính \"Non-Fiction\", thêm công thức để tính lợi nhuận mỗi quyển sách tại cột \"Margin\".","Profit equals the selling price minus the purchase price. On the \"Non-Fiction\" worksheet, add a formula to calculate the profit per book in the \"Margin\" column."),
            new ClsQuestion(true,18,39,"youtube.com","Trên trang tính \"Enrollment Summary\", đặt tên vùng A3:B7 là \"Enrollment\".","On the \"Enrollment Summary\" worksheet, name the range A3:B7 \"Enrollment\"."),
            new ClsQuestion(true,73,40,"youtube.com","Trên trang tính \"Products\", trên cột \"Projected Value\", thêm công thức nhân giá trị ở cột \"Current Value\" bởi vùng \"Increase\". Dùng tên vùng trên công thức thay vì tham chiếu ô hoặc giá trị.","On the \"Products\" worksheet, in the \"Projected Value\" column, add a formula that multiplies the value in the \"Current Value\" column by the \"Increase\" range. Use the range name in the formula instead of cell references or values."),
            new ClsQuestion(true,94,41,"youtube.com","Trên trang tính \"Projections\", trên cột \"Quarters 2\" thêm 1 công thức nhân giá trị cột \"Quarter 1\" bởi vùng \"Q2_Increase\". Dùng tên vùng trên công thức thay vì tham chiếu các ô hoặc giá trị.","On the \"Projections\" worksheet, in the \"Quarters 2\" column, add a formula that multiplies the value in the \"Quarter 1\" column by the \"Q2_Increase\" range. Use the range name in the formula instead of cell references or values."),
            new ClsQuestion(true,81,42,"youtube.com","Trên trang tính \"Prices\", trên cột \"Tax\", dùng công thức nhân giá trị ở cột \"Unit Price\" với ô L2.","On the \"Prices\" worksheet, in the \"Tax\" column, use a formula that multiplies the value in the \"Unit Price\" column by cell L2."),
            new ClsQuestion(true,95,43,"youtube.com","Trên trang tính \"Summary\", trên ô B15, dùng 1 hàm để hiển thị số giảm giá lớn nhất từ cột \"Total Sales\".","On the \"Summary\" worksheet, in cell B15, use a function to display the largest discount from the \"Total Sales\" column."),
            new ClsQuestion(true,106,44,"youtube.com","Trên trang tính \"Grade Criteria\", trên ô B28, thêm vào công thức tính tổng các giá trị của các vùng \"Total1\", \"Total2\",\"Total3\". Dùng tên vùng trên công thức thay vì tham chiếu ô hoặc giá trị.","On the \"Grade Criteria\" worksheet, in cell B28, add a formula that sums the values ​​in the \"Total1\", \"Total2\", and \"Total3\" ranges. Use range names in the formula instead of cell references or values."),
            new ClsQuestion(true,107,45,"youtube.com","Trên trang tính \"Exams\", trên ô E35, dùng 1 hàm để xác định có bao nhiêu học sinh không có kết quả \"Exam3\".","On the \"Exams\" worksheet, in cell E35, use a function to determine how many students did not have an \"Exam3\" result."),
            new ClsQuestion(true,62,46,"youtube.com","Trên trang tính \"New Policies\", ở cột \"Inactive Month\", dùng 1 hàm để đếm số của các tháng không có chính sách mới, bởi \"Agent\".","On the \"New Policies\" worksheet, in the \"Inactive Month\" column, use a function to count the number of months without a new policy, by \"Agent\"."),
            new ClsQuestion(true,51,47,"youtube.com","Trên trang \"New York City\", trong ô D23, dùng hàm để hiển thị số lớn nhất từ cột \"Air Miles\".","On the \"New York City\" worksheet, in cell D23, use a function to display the largest number from the \"Air Miles\" column."),
            new ClsQuestion(true,44,48,"youtube.com","Trên trang tính \"Key Accounts\", trong cột \"Monthly Average\", dùng 1 hàm để tính giá trị trung bình mỗi tháng cho mỗi tài khoản từ tháng 1 đến tháng 4.","On the \"Key Accounts\" worksheet, in the \"Monthly Average\" column, use a function to calculate the average monthly value for each account from January to April."),
            new ClsQuestion(true,37,49,"youtube.com","Trên trang tính \"February\", trên cột \"Policy Type\", dùng 1 hàm để hiển thị 2 ký tự đầu của \"Policy Number\" từ cột B.","On the \"February\" worksheet, in the \"Policy Type\" column, use a function to display the first two characters of the \"Policy Number\" from column B."),
            new ClsQuestion(true,99,50,"youtube.com","Trên trang tính \"Sales\", trên cột \"City Code\", sửa đổi công thức để chữ in hoa.","On the \"Sales\" worksheet, in the \"City Code\" column, modify the formula to uppercase."),
            new ClsQuestion(true,88,51,"youtube.com","Trên trang tính \"roster\" trong cột C chèn một công thức để sao chép và chuyển đổi thành các ký tự hoa trên cột A.","On the \"roster\" worksheet, in column C, insert a formula to copy and convert to uppercase characters in column A."),
            new ClsQuestion(true,89,52,"youtube.com","Trên trang tính \"roster\" trong cột B chèn một công thức chuyển đổi tên (first name), thành chữ thường.","On the \"roster\" worksheet, in column B, insert a formula to convert the first name to lowercase."),
            new ClsQuestion(true,90,53,"youtube.com","Trong ô C8 trên trang tính \"Roster\", sử dụng 1 hàm để sao chép ô A8 chỉ để ký tự đầu tiên được viết hoa.","In cell C8 on the \"Roster\" worksheet, use a function to copy cell A8 so that only the first letter is capitalized."),
            new ClsQuestion(true,63,54,"youtube.com","Trên trang tính \"Contact\", ở cột \"Email Address\", dùng 1 hàm để tạo ra địa chỉ email cho mỗi người dùng First Name và \"@humongousinsurance.com\".","On the \"Contact\" worksheet, in the \"Email Address\" column, use a function to generate an email address for each user using the First Name and \"@humongousinsurance.com\"."),
            new ClsQuestion(true,45,55,"youtube.com","Trên trang tính \"Contact\", trong cột \"Email Address\", dùng 1 hàm để tạo địa chỉ email cho mỗi người, dùng tên họ của từng người và \"@woodgrovebank.com\" (dùng tên họ và \"@woodgrovebank.com\" để tạo email),.","On the \"Contact\" worksheet, in the \"Email Address\" column, use a function to generate an email address for each person using each person's first name and \"@woodgrovebank.com\" (use the first name and \"@woodgrovebank.com\" to generate the email),."),
            new ClsQuestion(true,100,56,"youtube.com","Trên trang tính \"Historical Sales\", hiển thị các công thức thay vì giá trị.","On the \"Historical Sales\" worksheet, display the formulas instead of the values."),
            new ClsQuestion(true,38,57,"youtube.com","Trên trang tính \"February\", trên cột \"Discount\", dùng 1 hàm để hiển thị \"Yes\" nếu \"Years as Member\" cao hơn 3. Ngược lại, hiển thị \"No\".","On the \"February\" worksheet, in the \"Discount\" column, use a function to display \"Yes\" if \"Years as Member\" is greater than 3. Otherwise, display \"No\"."),
            new ClsQuestion(true,82,58,"youtube.com","Trên trang tính \"Prices\", ở cột \" Inventory Notice\" dùng 1 hàm để hiển thị \"Low\" nếu \"Inventory Level%\" thấp hơn 15%. Ngược lại gỡ bỏ \"Inventory Notice\" thành khoảng trống.","On the \"Prices\" worksheet, in the \"Inventory Notice\" column, use a function to display \"Low\" if \"Inventory Level%\" is less than 15%. Otherwise, remove \"Inventory Notice\" to a blank space."),
            new ClsQuestion(true,101,59,"youtube.com","Trên trang tính \"Authors\", trên cột \"Bonus\", dùng 1 hàm hiển thị \"500\" nếu \"Books Sold\" lớn hơn 10.000. Ngược lại, hiển thị \"100\".","On the \"Authors\" worksheet, in the \"Bonus\" column, use a function to display \"500\" if \"Books Sold\" is greater than 10,000. Otherwise, display \"100\"."),
            new ClsQuestion(true,49,60,"youtube.com","Thêm 1 hàm tại ô I2 trong trang tính \"Key Applications\" để hiển thị từ \"Yes\" nếu giá trị ô H2 cao hơn \"719\", ngược lại hiển thị từ \"No\". Điền vào những ô trong cột I để hiển thị những sinh viên có vượt qua hay không.","Add a function in cell I2 in the \"Key Applications\" worksheet to display the word \"Yes\" if the value of cell H2 is greater than \"719\", otherwise display the word \"No\". Fill in the cells in column I to display whether the students passed or not."),
            new ClsQuestion(true,24,61,"youtube.com","Trong ô F37 của trang tính \"Non-Fiction\", sử dụng 1 công thức Excel để tính trung bình giá bán của những quyển sách được xuất bản bơi \"Lucerne Publishing\".","In cell F37 of the \"Non-Fiction\" worksheet, use an Excel formula to calculate the average selling price of books published by \"Lucerne Publishing\"."),
            new ClsQuestion(true,68,62,"youtube.com","Trong ô E37 của trang tính \"October\" sử dụng 1 hàm để tính trung bình các ô E11:E35 có giá trị nợ (debit), trên \"$300\".","In cell E37 of the \"October\" worksheet, use a function to calculate the average of cells E11:E35 with a debit value above \"$300\"."),
            new ClsQuestion(true,8,63,"youtube.com","Trên trang tính \"Orders\", sử dụng 1 phương thức để tự động định dạng các ô chứa giá trị trên trung bình trong cột \"OrderTotal\" áp dụng \"Green Fill with Dark Green Text\". Sử dụng kỹ thuật tự động cập nhật định dạng nếu giá trị thay đổi.","On the \"Orders\" worksheet, use a method to automatically format cells containing values ​​above the average in the \"OrderTotal\" column using \"Green Fill with Dark Green Text\". Use a technique that automatically updates the format if the value changes."),
            new ClsQuestion(true,52,64,"youtube.com","Gỡ bỏ định dạng có điều kiện từ trang \"London\".","Remove conditional formatting from the \"London\" worksheet."),
            new ClsQuestion(true,74,65,"youtube.com","Trên trang tính \"Products\", trên cột \"Quantity\", dùng định dạng có điều kiện \"3 Traffic Lights (Unrimmed), định dạng cho các giá trị.","On the \"Products\" worksheet, in the \"Quantity\" column, use the \"3 Traffic Lights (Unrimmed),\" conditional format to format the values."),
            new ClsQuestion(true,96,66,"youtube.com","Trên trang tính \"Summary\", trên các ô \"F4:F11\", dùng định dạng có điều kiện để định dạng \"Yellow Fill with Dark Yellow Text\" cho các ô chứa giá trị hơn \"$5,000,000\".","You are revising a sales summary report for Northwind Traders. On the \"Summary\" worksheet, in cells \"F4:F11\", use the \"Yellow Fill with Dark Yellow Text\" conditional format for cells containing values ​​greater than \"$5,000,000\"."),
            new ClsQuestion(true,25,67,"youtube.com","Trên trang tính \"Fiction\" xoay hình ảnh thành 0 độ.","On the \"Fiction\" worksheet, rotate the image to 0 degrees."),
            new ClsQuestion(true,69,68,"youtube.com","Thêm tập tin \"coins.jpg\" trong thư mục picture bên phải tiêu đề \"Bank Statement\" trên trang tính \"October\".","Add the \"coins.jpg\" file in the picture folder to the right of the \"Bank Statement\" heading on the \"October\" worksheet."),
            new ClsQuestion(true,26,69,"youtube.com","Trên trang tính \"Non-Fiction\" worksheet, gỡ bỏ nền của hình ảnh và áp dụng 20% fill pattern.","On the \"Non-Fiction\" worksheet, remove the background of the image and apply a 20% fill pattern."),
            new ClsQuestion(true,39,70,"youtube.com","Trên trang tính \"Summary\", thêm vào mô tả văn bản thay thế \"Renewal data\" cho biểu đồ.","On the \"Summary\" worksheet, add the \"Renewal data\" alt text description to the chart."),
            new ClsQuestion(true,19,71,"youtube.com","Trên trang tính \"Reveue\", chuyển đổi các ô A3:B7 thành dạng bảng với tiêu đề. Dùng kiểu \"Green, Table Style Light 14\"","On the \"Reveue\" worksheet, convert cells A3:B7 to a table with headings. Use the \"Green, Table Style Light 14\" style"),
            new ClsQuestion(true,75,72,"youtube.com","Trên trang tính \"Products\" dùng kiểu \"White, Table Style Medium 1\" cho bảng.","On the \"Products\" worksheet Use the \"White, Table Style Medium 1\" style for the table."),
            new ClsQuestion(true,46,73,"youtube.com","Trên trang tính \"New Accounts\", gỡ bỏ dòng chứa dữ liệu \"Tailspin Toys\". Không thay đổi nội dung bên ngoài bảng.","On the \"New Accounts\" worksheet, remove the row containing the \"Tailspin Toys\" data. Do not change the content outside the table."),
            new ClsQuestion(true,20,74,"youtube.com","Trên trang tính \"Last Semester\" , gỡ bỏ dòng của bảng có chứa dữ liệu \"Agriculture\". Không thay đổi nội dung bên ngoài bảng.","On the \"Last Semester\" worksheet, remove the table row containing the \"Agriculture\" data. Do not change the content outside the table."),
            new ClsQuestion(true,108,75,"youtube.com","Trên trang tính \"Tasks\", điều chỉnh tùy chọn kiểu của bảng là tự động tô đậm mỗi dòng khác nhau.","On the \"Tasks\" worksheet, adjust the table style option to automatically bold each different row."),
            new ClsQuestion(true,64,76,"youtube.com","Trên trang tính \"New Policies\", thêm \"Total Row\" vào bảng. Cấu hình \"Total Row\" để hiển thị số chính sách mới cho mỗi tháng và tổng mỗi tháng 6.","On the \"New Policies\" worksheet, add a \"Total Row\" to the table. Configure the \"Total Row\" to display the number of new policies for each month and the total for each month 6."),
            new ClsQuestion(true,97,77,"youtube.com","Trên trang tính \"Region 1\", dùng chức năng sắp xếp đa tầng. Sắp xếp dữ liệu của bảng bởi \"Products\" (A đến Z), và sau đó bởi \"Total Sales\" ( Lớn nhất đến nhỏ nhất),.","On the \"Region 1\" worksheet, use the multi-level sort function. Sort the table data by \"Products\" (A to Z), and then by \"Total Sales\" (Largest to Smallest),."),
            new ClsQuestion(true,53,78,"youtube.com","Trên trang \"New York City\", dùng chức năng sắp xếp đa tầng. Sắp xếp dữ liệu của bảng bởi \"Country or Region\" (A đến Z), và sau đó bởi \"City\" ( A đến Z),.","On the \"New York City\" sheet, use the multi-level sort function. Sort the table data by \"Country or Region\" (A to Z), and then by \"City\" (A to Z),."),
            new ClsQuestion(true,83,79,"youtube.com","Trên trang tính \"Orders\", lọc dữ liệu của bảng để hiển thị các đơn đặt hàng từ \"Alpine Ski House\".","On the \"Orders\" sheet, filter the table data to show orders from \"Alpine Ski House\"."),
            new ClsQuestion(true,40,80,"youtube.com","Trên trang tính \"March\", lọc dữ liệu của bảng chỉ hiển thị các chính sách với \"Policy Type\" của \"MP\".","On the \"March\" sheet, filter the table data to show only policies with a \"Policy Type\" of \"MP\"."),
            new ClsQuestion(true,109,81,"youtube.com","Trên trang tính \"Tasks\", đặt tên cho bảng là \"Tasks\".","On the \"Tasks\" sheet, name the table \"Tasks\"."),
            new ClsQuestion(true,3,82,"youtube.com","Trên trang tính \"Classes\", chuyển đổi phạm vi của bảng thành ô. Giữ nguyên định dạng.","On the \"Classes\" sheet, convert the table range to cells. Keep the formatting the same."),
            new ClsQuestion(true,21,83,"youtube.com","Trên trang tính \"Next Semester\", tạo 1 biểu đồ \"Clustered Column\" để hiển thị tên \"Program\" và dữ liệu \"Average Cost per Student\".Vị trí biểu đồ bên phải của bảng. Kích thước chính xác và vị trí của biểu đồ không quan trọng.","On the \"Next Semester\" sheet, create a \"Clustered Column\" chart to show the \"Program\" name and the \"Average Cost per Student\" data. Position the chart to the right of the table. The exact size and position of the chart do not matter."),
            new ClsQuestion(true,110,84,"youtube.com","Trên trang biểu đồ \"Score Distribution\", gỡ bỏ chú giải và chỉ hiển thị các giá trị như nhãn dữ liệu phía trên mỗi cột.","On the \"Score Distribution\" chart sheet, remove the legend and display only the values ​​as data labels above each column."),
            new ClsQuestion(true,84,85,"youtube.com","Trên trang có biểu đồ \"Inventory\", chèn tiêu đề phía trên biểu đồ, ở trên cùng của vùng chứa đồ thị. Hiển thị giá trị phần trăm như \"Data Labels\" vào bên phải của mỗi thanh dữ liệu.","On the sheet with the \"Inventory\" chart, insert a title above the chart, at the top of the chart container. Display the percentage values ​​as \"Data Labels\" to the right of each data bar."),
            new ClsQuestion(true,54,86,"youtube.com","Trên trang \"London\", biểu đồ \"Air Miles\", hiển thị dữ liệu của bảng không có \"Legend Keys\".","On the \"London\" sheet, the \"Air Miles\" chart displays the table data without \"Legend Keys\"."),
            new ClsQuestion(true,4,87,"youtube.com","Trên trang tính \"Instructional Hours\", sửa đổi biểu đồ hiển thị \"Hours\" như \"Primary Vertical Axis Title\".","On the \"Instructional Hours\" sheet, modify the chart to display \"Hours\" as \"Primary Vertical Axis Title\"."),
            new ClsQuestion(true,65,88,"youtube.com","Trên trang tính \"New policies\", thay đổi các thành phần trên biểu đồ bởi việc thay đổi bố cục thành \"Layout 3\".","On the \"New policies\" sheet, change the chart elements by changing the layout to \"Layout 3\"."),
            new ClsQuestion(true,76,89,"youtube.com","Trên trang tính \"Summary\", dùng màu \"Colorful Palette 2\" cho biểu đồ.","On the \"Summary\" sheet, use the \"Colorful Palette 2\" color for the chart."),
            new ClsQuestion(true,22,90,"youtube.com","Trên trang tính \"Enrollment Summary\", dùng kiểu \"Style 7\" và màu \"Monochromatic Palette 6\" cho biểu đồ.","On the \"Enrollment Summary\" sheet, use the \"Style 7\" style and \"Monochromatic Palette 6\" color for the chart."),
            new ClsQuestion(true,98,91,"youtube.com","Trên trang có biểu đồ \"Comparison\", chuyển đổi dữ liệu giữa các trục.","On the sheet with the \"Comparison\" chart, switch the data between the axes."),
            new ClsQuestion(true,47,92,"youtube.com","Trên trang tính \"New Account\", ở biểu đồ \"Account Balances\", chuyển đổi dữ liệu trên trục.","On the \"New Account\" sheet, in the \"Account Balances\" chart, switch the data on the axes."),
            new ClsQuestion(true,5,93,"youtube.com","Từ trang tính \"Graduation\", dời biểu đồ đến trang biểu đồ mới đặt tên \"Graduation Chart\".","From the \"Graduation\" worksheet, move the chart to a new chart sheet named \"Graduation Chart\"."),
            new ClsQuestion(true,102,94,"youtube.com","Trên trang tính \"Summary\", mở rộng biểu đồ để bao gồm dữ liệu \"Current Year\".","On the \"Summary\" worksheet, expand the chart to include the \"Current Year\" data."),
            new ClsQuestion(true,66,95,"youtube.com","Trên trang tính \"New Policies\", ở các ô J5:J13, chèn biểu đồ \"Win/Loss\" để đối chiếu giá trị từ tháng 1 đến tháng 6.","On the \"New Policies\" worksheet, in cells J5:J13, insert a \"Win/Loss\" chart to compare values ​​from January to June."),
            new ClsQuestion(true,6,96,"youtube.com","Trên trang tính \"Enrollment\", trên các ô G5:G25, chèn biểu đồ cột Sparklines đối chiếu các giá trị của khóa học cuối cùng, hiện tại và tiếp theo cho mỗi tiết học.","On the \"Enrollment\" worksheet, in cells G5:G25, insert a Sparklines column chart that compares the values ​​of the last, current, and next course for each class."),
            new ClsQuestion(true,57,97,"youtube.com","Trên trang tính \"Materials\" tại ô A6 tạo một liên kết đến ô A18 trên trang tính \"Categories\".","On the \"Materials\" worksheet, in cell A6, create a link to cell A18 on the \"Categories\" worksheet."),
            new ClsQuestion(true,34,98,"youtube.com","Tại ô C5 của trang tính \"Shareholders Info\" thêm 1 liên kết đến \"http://tailspintoys.com/beyond.html\". Hiển thị văn bản \"More Info\" trong ô.","In cell C5 of the \"Shareholders Info\" worksheet, add a link to \"http://tailspintoys.com/beyond.html\". Display the text \"More Info\" in the cell."),
            new ClsQuestion(true,103,99,"youtube.com","Trên trang tính \"Summary\", từ văn bản đã có ở ô A2, chèn 1 Hyperlink (siêu liên kết), đến \"http://www.nodpublishers.com\" với ScreenTip hiển thị \"Company Website\".","On the \"Summary\" worksheet, from the text already in cell A2, insert a Hyperlink to \"http://www.nodpublishers.com\" with a ScreenTip showing \"Company Website\"."),
            new ClsQuestion(true,58,100,"youtube.com","Trên trang tính \"Materials\" chèn \"Confidential\" ở trên đầu (header), phía bên phải của mỗi trang.","On the \"Materials\" worksheet, insert \"Confidential\" at the top right of each page."),
            new ClsQuestion(true,59,101,"youtube.com","Chèn số trang ở giữa chân trang sử dụng kiểu \"page 1 of ?\".","Insert a page number in the center of the footer using the style \"page 1 of ?\"."),
            new ClsQuestion(true,41,102,"youtube.com","Cấu hình trang tính \"January\" chỉ có các ô A4:F20 được in ra.","Configure the \"January\" worksheet so that only cells A4:F20 are printed."),
            new ClsQuestion(true,93,103,"youtube.com","Trên trang tính \"Inbound calls\" thiết lập vùng in (print area), A1:C19.","On the \"Inbound calls\" worksheet, set the print area to A1:C19."),
            new ClsQuestion(true,104,104,"youtube.com","Gỡ bỏ thông tin thuộc tính và cá nhân từ file được lưu. Không gỡ bỏ nội dung khác.","Remove the property and personal information from the saved file. Do not remove other content."),
            new ClsQuestion(true,60,105,"youtube.com","Thay đổi cài đặt in cho trang tính \"Materials\" để in tất cả các cột trên một trang theo hướng ngang.","Change the print settings for the \"Materials\" worksheet to print all columns on one page in landscape orientation."),
            new ClsQuestion(true,91,106,"youtube.com","Phục hồi lại hình dạng trang tính \"roster\" để hiển thị dòng 7 trên mỗi trang khi in.","Reshape the \"roster\" worksheet to display row 7 on each page when printed."),
            new ClsQuestion(true,35,107,"youtube.com","Sửa đổi cài đặt in để phù hợp với từng trang tính trên 1 trang.","Modify the print settings to fit each sheet on one page."),
            new ClsQuestion(true,77,108,"youtube.com","Trên trang tính \"Q2_Sales\" thiết lập các lề thành \"Narrow\".","On the \"Q2_Sales\" sheet set the margins to \"Narrow\"."),
            new ClsQuestion(true,36,109,"youtube.com","Cấu hình lại các lề trên và dưới là \"1\", trái và phải là \"1.5\".","Reconfigure the top and bottom margins to \"1\", and the left and right margins to \"1.5\"."),
            new ClsQuestion(true,27,110,"youtube.com","Trong phần thuộc tính của tài liệu (document properties), thêm \"Lucerne Publishing\" cho tên công ty (company name),.","In the document properties add \"Lucerne Publishing\" for the company name."),
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
