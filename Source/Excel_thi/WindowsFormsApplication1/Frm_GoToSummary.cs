using DevExpress.Utils;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.XtraTreeList;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MOS_WORD_TEST.WindowsFormsApplication1
{
    public partial class Frm_GoToSummary : Form
    {
        public delegate void ChangeAnswer(ProjectEventArgs e);
        public DataTable DataSource { get; set; }
        public event ChangeAnswer OnChangeAnswer;
        //private static bool isButtonDisabled = false;
        public Frm_GoToSummary()
        {
            InitializeComponent();
            
            if (Home.isButtonDisabled)
            {
                btnFinish.Enabled = !Home.isButtonDisabled;
                btnFinish.BackgroundColor = Color.Gray;
            }
        }

        private void Frm_GoToSummary_Load(object sender, EventArgs e)
        {
            //gridView1.Columns.Clear();
            //gridControl1.DataSource = DataSource;
            //gridView1.BestFitColumns();
            treeList1.DataSource = DataSource;
            treeList1.ParentFieldName = "ParentKey";
            treeList1.KeyFieldName = "QuestionKey";
            //treeList1.Columns["ProjectIndex"].Visible = true;
            treeList1.PopulateColumns();
            treeList1.RowHeight = -1;
            treeList1.OptionsBehavior.AutoNodeHeight = true;
            var memo = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
            memo.WordWrap = true;
            treeList1.RepositoryItems.Add(memo);
            treeList1.Appearance.HeaderPanel.Options.UseTextOptions = true;
            treeList1.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            treeList1.Appearance.HeaderPanel.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            treeList1.Columns[0].Visible = false;
            treeList1.Columns[1].Caption = "Project";
            treeList1.Columns[2].Caption = "Question Number";
            treeList1.Columns[2].Width = 80;
            treeList1.Columns[2].AppearanceCell.Options.UseTextOptions = true;
            treeList1.Columns[2].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            treeList1.Columns[2].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            treeList1.Columns[3].Caption = "Question Content";
            treeList1.Columns[3].ColumnEdit = memo;
            treeList1.Columns[3].Width = 300;
            treeList1.Columns[3].AppearanceCell.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            treeList1.Columns[4].Caption = "Mark Complete";
            treeList1.Columns[4].Width = 80;
            treeList1.Columns[4].AppearanceCell.Options.UseTextOptions = true;
            treeList1.Columns[4].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            treeList1.Columns[4].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            treeList1.Columns[5].Caption = "Mark for Review";
            treeList1.Columns[5].Width = 80;
            treeList1.Columns[5].AppearanceCell.Options.UseTextOptions = true;
            treeList1.Columns[5].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            treeList1.Columns[5].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            treeList1.Columns[6].Visible = false;
            treeList1.Columns[7].Visible = Home.isButtonDisabled;
            treeList1.Columns[7].Caption = "Result";
            treeList1.Columns[7].AppearanceCell.Options.UseTextOptions = true;
            treeList1.Columns[7].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            treeList1.Columns[7].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;

            var formatCondition = new DevExpress.XtraTreeList.StyleFormatConditions.StyleFormatCondition
            {
                Column = treeList1.Columns[7],          // cột áp dụng
                ApplyToRow = false,                     // chỉ đổi màu ô
                Condition = FormatConditionEnum.Equal,  // so sánh bằng
                Value1 = "False"                        // giá trị cần kiểm tra
            };

            // Cài đặt màu sắc cho Appearance
            formatCondition.Appearance.ForeColor = Color.Red;
            formatCondition.Appearance.Options.UseForeColor = true;

            // Thêm vào danh sách điều kiện format của TreeList
            treeList1.FormatConditions.Add(formatCondition);

            //treeList1.BestFitColumns();
            treeList1.ExpandAll();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnFinish_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc nộp bài?", "Cảnh Báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
            {
                btnFinish.Enabled = false;
                btnFinish.BackgroundColor = Color.Gray;
                Home.isButtonDisabled = true;
                treeList1.Columns[7].Visible = true;

                Form1 frm1 = Application.OpenForms["Form1"] as Form1;
                if (frm1 != null)
                {
                    frm1.buttonxhdh.Enabled = false;
                    frm1.buttonxhdh.BackColor = Color.Gray;
                    frm1.buttonRefresh.Enabled = false;
                    frm1.buttonRefresh.BackColor = Color.Gray;
                    frm1.buttonReset.Enabled = false;
                    frm1.buttonReset.BackColor = Color.Gray;
                    frm1.buttonSubmit.Enabled = false;
                    frm1.buttonSubmit.BackColor = Color.Gray;
                    frm1.btnMarkComplete.Enabled = false;
                    frm1.btnMarkComplete.BackColor = Color.Gray;
                    frm1.btnMaskReview.Enabled = false;
                    frm1.btnMaskReview.BackColor = Color.Gray;
                }

                int correctAnswer = 0;
                string str1 = "";
                foreach (DataRow r in DataSource.Rows)
                {
                    string value = r["Value"].ToString().Trim();
                    if (value == "True")
                    {
                        correctAnswer++;

                        //if (r["QuestionIndex"].ToString() != "")
                        //    str1 = str1 + "Project " + r["ProjectIndex"] + ": Câu " + r["QuestionIndex"] + ", ";
                    }
                }

                int total = (int)Math.Round((correctAnswer * 28.6), 0);
                if (total > 1000)
                {
                    total = 1000;
                }

                //Kết quả thi

                //string str2 = "Câu đúng: " + correctAnswer.ToString() + "; " + str1;
                string str2 = "Điểm thi: " + total.ToString();
                string str3 = "<p>" + DateTime.Now.ToShortDateString() + "; " + str2 + "</p>";
                //if (correctAnswer <= 1)
                //    return;
                string str4 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "MOS360");
                if (!Directory.Exists(str4))
                    Directory.CreateDirectory(str4);
                string path = Path.Combine(str4, "MOS360 THI - EXCEL.html");
                if (File.Exists(path))
                {
                    TextWriter textWriter = (TextWriter)new StreamWriter(path, true);
                    textWriter.WriteLine(str3);
                    textWriter.Close();
                }
                else
                {
                    TextWriter textWriter = (TextWriter)new StreamWriter(path, true);
                    string str5 = "<h2 style=\"text-align: center;\"><span style=\"color: #ff0000;\"><strong>MOS360 - X&oacute;a tan nỗi lo CHUẨN ĐẦU RA cho sinh vi&ecirc;n.</strong></span></h2>\r\n<p style=\"padding-left: 40px;\">Kh&ocirc;ng cần ho&agrave;n hảo - chỉ cần bắt đầu. Gửi b&agrave;i cho MOS360 để giữ vững phong độ nh&eacute; 🚀</p>\r\n<p style=\"padding-left: 40px;\">Zalo: <a href=\"https://zalo.me/0912888360\" target=\"_blank\" rel=\"noopener\">0912.888.360</a></p>\r\n<p style=\"padding-left: 40px;\">Fanpage: <a href=\"https://www.facebook.com/mos360.vn\" target=\"_blank\">https://www.facebook.com/mos360.vn</a></p>\r\n<p style=\"padding-left: 40px;\">Website: <a href=\"https://mos360.vn\" target=\"_blank\">https://mos360.vn</a></p>\r\n<p style=\"padding-left: 40px;\">Ch&uacute;c bạn học MOS thật vui, l&agrave;m b&agrave;i thật 'phi&ecirc;u' v&agrave; điểm cao v&egrave;o v&egrave;o nh&eacute; 🚀🎯</p>\r\n<p style=\"text-align: center;\"><strong>KẾT QUẢ THI EXCEL</strong></p>\r\n<p style=\"padding-left: 40px;\">&nbsp;</p>";
                    textWriter.WriteLine(str5);
                    textWriter.WriteLine(str3);
                    textWriter.Close();
                }

                //End Kết quả thi

                MessageBox.Show($@"
Số câu đúng: {correctAnswer}/35

Điểm thi của bạn là: {total}/1000

Bạn có thể xem lại kết quả thi ở folder MOS360 trên Desktop!", "Kết quả thi");
            }
        }

        private void gridView1_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            //var a = Convert.ToInt32(gridView1.GetRowCellValue(e.RowHandle, "ProjectIndex"));
            //ProjectEventArgs changeData = new ProjectEventArgs();
            //changeData.ProjectIndex = Convert.ToInt32(gridView1.GetRowCellValue(e.RowHandle, "ProjectIndex"));
            //changeData.QuestionIndex = Convert.ToInt32(gridView1.GetRowCellValue(e.RowHandle, "QuestionIndex"));
            //OnChangeAnswer?.Invoke(changeData);
            //btnClose_Click(null,null);
        }

        private void treeList1_RowCellClick(object sender, DevExpress.XtraTreeList.RowCellClickEventArgs e)
        {
            if (e.Node == null || e.Node.GetValue("QuestionIndex").ToString() == string.Empty)
                return;

            var projectIndex = Convert.ToInt32(e.Node.GetValue("ProjectIndex"));
            ProjectEventArgs changeData = new ProjectEventArgs();
            changeData.ProjectIndex = projectIndex;
            changeData.QuestionIndex = Convert.ToInt32(e.Node.GetValue("QuestionIndex"));

            OnChangeAnswer?.Invoke(changeData);
            btnClose_Click(null, null);
        }
    }

    public class ProjectEventArgs : EventArgs
    {
        public int ProjectIndex { get; set; }
        public int QuestionIndex { get; set; }
    }
}
