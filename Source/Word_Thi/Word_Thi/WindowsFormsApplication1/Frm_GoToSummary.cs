using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
        public Frm_GoToSummary()
        {
            InitializeComponent();
        }

        private void Frm_GoToSummary_Load(object sender, EventArgs e)
        {
            gridView1.Columns.Clear();
            gridControl1.DataSource = DataSource;
            gridView1.BestFitColumns();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnFinish_Click(object sender, EventArgs e)
        {
            int correctAnswer = 0;
            foreach (DataRow r in DataSource.Rows)
            {
                bool value = Convert.ToBoolean(r["Status"]);
                if (value) correctAnswer++;
            }

            int total = (int)Math.Round((correctAnswer * 28.6),0);
            if(total > 1000)
            {
                total = 1000;
            }
            MessageBox.Show($@"Điểm của bạn là {total}/1000");
        }

        private void gridView1_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            var a = Convert.ToInt32(gridView1.GetRowCellValue(e.RowHandle, "ProjectIndex"));
            ProjectEventArgs changeData = new ProjectEventArgs();
            changeData.ProjectIndex = Convert.ToInt32(gridView1.GetRowCellValue(e.RowHandle, "ProjectIndex"));
            changeData.QuestionIndex = Convert.ToInt32(gridView1.GetRowCellValue(e.RowHandle, "QuestionIndex"));
            OnChangeAnswer?.Invoke(changeData);
            btnClose_Click(null,null);
        }
    }

    public class ProjectEventArgs : EventArgs
    {
        public int ProjectIndex { get; set; }
        public int QuestionIndex { get; set; }
    }
}
