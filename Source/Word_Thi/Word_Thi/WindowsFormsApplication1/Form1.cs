using Microsoft.Office.Interop.Word;
using MOS_WORD_TEST.Base;
using MOS_WORD_TEST.Properties;
using MOS_WORD_TEST.WindowsFormsApplication1;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Point = System.Drawing.Point;

namespace MOS_WORD_TEST
{
    public partial class Form1 : Form
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        public static List<Exam> listExam;
        public static Exam currentExam;
        private Project currentProject;
        private Question currentQuestion;
        public static string Language;
        private string pathWork;
        private string pathRun;
        private string pathFileOfficeMaHoa;
        private string pathFileOffice;
        private string pathReset;
        private int works;
        private int workIndex;
        private Microsoft.Office.Interop.Word.Application a;
        private string[] ImageFile;
        private int screen_height;
        private int screen_width;
        private Document d;
        private DateTime timeStrart;
        private int Cau_So;
        private int next = 0;
        private int soLanReSet = 0;
        private double Diem = 0.0;
        private bool[] CacCauDaCheck;
        private int So_Cau_Dung = 0;
        private int So_Cau_Sai = 0;
        private int Tong_So_Cau = 35;
        private List<int> lsViTri;
        private bool chotat = false;
        private int currentest = 0;
        public object missing = (object)Missing.Value;
        public Form1()
        {
            InitializeComponent();
            this.textBox1.MouseClick += new MouseEventHandler(this.textBox1_MouseClick);
            this.getEnviroment();
        }

        private void a_DocumentBeforeClose(Document Doc, ref bool Cancel)
        {
            if (this.chotat)
                return;
            Cancel = true;
        }

        private void textBox1_MouseClick(object sender, MouseEventArgs e)
        {
            this.textBox1.Text = this.panel2.VerticalScroll.Value.ToString();
            if (this.lsViTri.Count > this.Cau_So)
                this.lsViTri[this.Cau_So] = this.panel2.VerticalScroll.Value;
            else
                this.lsViTri.Add(this.panel2.VerticalScroll.Value);
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //this.Cau_So = this.checkedListBox1.SelectedIndex;
            //if (this.lsViTri.Count < this.checkedListBox1.Items.Count)
            //{
            //    this.Cau_So = this.checkedListBox1.SelectedIndex;
            //    this.panel2.VerticalScroll.Value = this.Cau_So * this.panel2.VerticalScroll.Maximum / this.Tong_So_Cau + this.panel2.VerticalScroll.Minimum;
            //    this.richTextBox1.Text = "KQ";
            //}
            //else
            //{
            //    this.Cau_So = this.checkedListBox1.SelectedIndex;
            //    this.panel2.VerticalScroll.Value = this.lsViTri[this.Cau_So];
            //    this.richTextBox1.Text = "KQ";
            //}
            this.currentProject = currentExam.ProjectIndex[this.checkedListBox1.SelectedIndex];
            this.currentQuestion = currentProject.Questions[0];
            loadcaucanhoiNew(this.checkedListBox1.SelectedIndex);
        }

        private void getEnviroment()
        {
            int length = 0;// listExam.Count;// Directory.GetDirectories(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\MaHoa")).Length;
            //Random r = new Random();
            if (Program.TypeOfTest == 0)
            {
                //Làm đề số 1
                this.button1.Visible = true;
                this.textBox2.Visible = true;
                //currentExam = listExam[0];
                currentest = 0;
                length = currentExam.ProjectIndex.Count;
                Program.Tong = length;
                Program.Lessons = new int[length];
                for (int index = 0; index < length; ++index)
                    Program.Lessons[index] = index;
            }
            else
            {
                Random rExam = new Random();
                int examIndex = rExam.Next(0, listExam.Count - 1);
                //currentExam = listExam[examIndex];
                //Làm đề random
                this.button1.Visible = false;
                this.textBox2.Visible = false;
                Program.Tong = currentExam.ProjectIndex.Count;
                Program.Lessons = new int[length];
                for (int index = 0; index < length; ++index)
                    Program.Lessons[index] = index;
            }

            //if (Program.TypeOfTest == 0)
            //{
            //    this.button1.Visible = true;
            //    this.textBox2.Visible = true;
            //    Program.Tong = length;
            //    Program.Lessons = new int[length];
            //    for (int index = 0; index < length; ++index)
            //        Program.Lessons[index] = index;
            //}
            //else
            //{
            //    this.button1.Visible = false;
            //    this.textBox2.Visible = false;
            //    Program.Tong = 7;
            //    Program.Lessons = new int[length];
            //    for (int i = 0; i < length; ++i)
            //        Program.Lessons[i] = this.chonBai(0, length, i, r);
            //}
        }

        private int chonBai(int min, int max, int i, Random r)
        {
            int kq;
            do
            {
                kq = r.Next(min, max);
            }
            while (this.tontai(kq, i));
            return kq;
        }

        private bool tontai(int kq, int c)
        {
            for (int index = 0; index < c; ++index)
            {
                if (Program.Lessons[index] == kq)
                    return true;
            }
            return false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = true;
            int num = this.Height / 5;
            int width = this.Width;
            int y = 5;
            int leftPadRight = 5;
            int rightPadLeft = 5;
            this.screen_height = this.Height;
            this.screen_width = this.Width;
            this.WindowState = FormWindowState.Normal;
            this.Height = num;
            this.Width = width;
            this.Location = new Point(0, num * 4);

            //Project
            this.checkedListBox1.Location = new Point(0, 0);
            this.checkedListBox1.Height = this.Height;
            //this.panelQuestion.Location = new Point(this.checkedListBox1.Width, this.buttonExit.Height + y);
            this.panelQuestion.Location = new Point((screen_width / 10) * 1, this.buttonExit.Height + y + y);
            this.panelQuestion.Height = this.Height - 40;
            //this.panelQuestion.Width = this.Width - this.checkedListBox1.Width - 200;
            this.panelQuestion.Width = (screen_width / 10) * 8;
            //var a = tabControl1.GetTabRect(0);
            //var c = a.Width;
            //this.panelJumpQuestion.Location = new Point(c, 0);
            this.panelJumpQuestion.Location = new Point((screen_width / 10) * 1, this.buttonExit.Height + y + y + y);
            //this.panelJumpQuestion.Height = 25;
            //this.panelJumpQuestion.Width = this.panelQuestion.Width - c;
            this.panelJumpQuestion.Height = this.buttonExit.Height;
            this.panelJumpQuestion.Width = screen_width;

            //Panel Right
            //this.panel2.Location = new Point(0, 0);
            //this.panel2.Height = this.tabPage2.Height;
            //this.panel2.Width = this.tabPage2.Width;

            // richTextTA
            this.richTextTA.Location = new Point((screen_width / 10) * 1, this.buttonExit.Height + this.panelJumpQuestion.Height + (y * 3));
            this.richTextTA.Height = this.Height - this.buttonExit.Height - this.panelJumpQuestion.Height - 50;
            this.richTextTA.Width = (screen_width / 10) * 8;


            //panelButton
            this.panelButton.Location = new Point(0, 0);
            this.panelButton.Height = this.buttonExit.Height + y + y;
            this.panelButton.Width = screen_width;
            this.panelButton.Controls.Add(this.label1);
            this.panelButton.Controls.Add(this.buttonCheck);

            this.panelButton.Controls.Add(this.labelProject);

            this.panelButton.Controls.Add(this.buttonExit);
            this.panelButton.Controls.Add(this.buttonRefresh);
            this.panelButton.Controls.Add(this.buttonSubmit);
            this.panelButton.Controls.Add(this.buttonReset);
            this.panelButton.Controls.Add(this.buttonxhdh);

            //int panelButtonWidth = this.checkedListBox1.Width + this.panelQuestion.Width;
            int panelButtonWidth = screen_width;

            // Exit
            this.buttonExit.Location = new Point(panelButtonWidth - this.buttonExit.Width - rightPadLeft, y);

            // HDSD
            //this.buttonxhdh.Location = new Point(panelButtonWidth - this.buttonxhdh.Width - this.buttonExit.Width - rightPadLeft, y);

            // Bố trí lại
            //this.buttonRefresh.Location = new Point(panelButtonWidth - this.buttonRefresh.Width - this.buttonxhdh.Width - this.buttonExit.Width - rightPadLeft, y);

            // Submit Project
            //this.buttonSubmit.Location = new Point(panelButtonWidth - this.buttonSubmit.Width - this.buttonRefresh.Width - this.buttonxhdh.Width - this.buttonExit.Width - rightPadLeft, y);
            this.buttonSubmit.Location = new Point(panelButtonWidth - this.buttonSubmit.Width - this.buttonExit.Width - rightPadLeft, y);

            // Reset Project
            //this.buttonReset.Location = new Point(panelButtonWidth - this.buttonReset.Width - this.buttonRefresh.Width - this.buttonSubmit.Width - this.buttonxhdh.Width - this.buttonExit.Width - rightPadLeft, y);
            this.buttonReset.Location = new Point(panelButtonWidth - this.buttonReset.Width - this.buttonSubmit.Width - this.buttonExit.Width - rightPadLeft, y);

            // Timer
            //this.label1.Location = new Point(this.checkedListBox1.Width + y, y);
            this.label1.Location = new Point(y, y);

            // Check -> 2.	Go To Summary
            //this.buttonCheck.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width, this.Height - this.buttonCheck.Height);
            this.buttonCheck.Location = new Point(this.label1.Width + y, y);

            // HDSD
            this.buttonxhdh.Location = new Point(this.label1.Width + this.buttonCheck.Width + leftPadRight, y);

            // Bố trí lại
            this.buttonRefresh.Location = new Point(this.label1.Width + this.buttonCheck.Width + this.buttonxhdh.Width + leftPadRight, y);

            // Label Project
            this.labelProject.Location = new Point((panelButtonWidth / 2) - (this.labelProject.Width / 2), y);


            //this.btnMarkComplete.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width, this.Height - this.btnMarkComplete.Height - y);
            //this.btnMaskReview.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width, this.Height - this.btnMarkComplete.Height - this.btnMaskReview.Height - y - y);

            this.btnMarkComplete.Location = new Point((screen_width / 2) - 100 - this.btnMarkComplete.Width, this.Height - this.btnMarkComplete.Height - y);
            this.btnMaskReview.Location = new Point((screen_width / 2) + 100, this.Height - this.btnMarkComplete.Height - y);



            this.button2.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.button2.Width, 10000);
            this.comboBox1.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.button2.Width - this.comboBox1.Width, 10000);
            this.buttonSave.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.comboBox1.Width - this.button2.Width, 10000);
            this.textBox1.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.comboBox1.Width - this.textBox1.Width, 10000);
            this.button1.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.textBox1.Width - this.comboBox1.Width - this.button1.Width, 10000);
            this.textBox2.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.textBox1.Width - this.button1.Width - this.comboBox1.Width - this.textBox2.Width, 10000);


            this.buttonShowHide.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.textBox1.Width - this.button1.Width - this.textBox2.Width - this.comboBox1.Width - this.buttonShowHide.Width, 10000);
            this.buttonHelp.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width - this.buttonExit.Width - this.buttonReset.Width - this.label1.Width - this.buttonSave.Width - this.textBox1.Width - this.button1.Width - this.textBox2.Width - this.comboBox1.Width - this.buttonShowHide.Width - this.buttonHelp.Width, 10000);

            this.panel4.Location = new Point(this.checkedListBox1.Width + this.panelQuestion.Width, this.Height - this.buttonCheck.Height - this.panel4.Height);



            //this.loadcaucanhoiNew(currentProject.ProjectIndex);
            this.a = (Microsoft.Office.Interop.Word.Application)Activator.CreateInstance(System.Type.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
            this.a.Visible = true;
            this.a.WindowState = WdWindowState.wdWindowStateNormal;
            this.a.Top = 0;
            this.a.Left = -10;
            this.a.Width = width;
            this.a.Height = num * 3;
            // ISSUE: method pointer
            // ISSUE: object of a compiler-generated type is created
            ((ApplicationEvents4_Event)this.a).DocumentBeforeClose += this.a_DocumentBeforeClose;
            LoadProject(0);
            this.SetUp(currentExam.ProjectIndex[0].ProjectIndex);
            currentQuestion = currentProject.Questions[0];
            loadcaucanhoiNew(0);
            this.timeStrart = DateTime.Now;
            this.timeStrart = this.timeStrart.AddMinutes(50.0);
            this.timer1.Start();
        }

        private void LoadProject(int projectIndex)
        {
            //this.labelProject.Text = $@"Project {(projectIndex + 1).ToString()} of {currentExam.ProjectIndex.Count.ToString()} : {currentExam.ProjectIndex[projectIndex].ProjectName}";
            this.labelProject.Text = $@"Project {(projectIndex + 1).ToString()} of {currentExam.ProjectIndex.Count.ToString()}";

            this.checkedListBox1.Items.Clear();
            for (int i = 0; i < currentExam.ProjectIndex.Count; i++)
            {
                var project = currentExam.ProjectIndex[i];
                this.checkedListBox1.Items.Add((object)(i + 1).ToString());
            }

            currentProject = currentExam.ProjectIndex[projectIndex];
        }

        private void loadcaucanhoiNew(int projectIndex)
        {

            this.screen_width = this.Width;

            this.panelJumpQuestion.Controls.Clear();
            currentProject = currentExam.ProjectIndex[projectIndex];
            //if(currentProject.Questions.Count > 1)
            //{
            AddBackButton();
            //}

            int numQuestions = currentProject.Questions.Count;

            int startX = (((this.screen_width / 10) * 8) - ((numQuestions + 2) * 110)) / 2;

            //MessageBox.Show(startX.ToString());

            for (int i = 1; i <= numQuestions; i++)
            {
                var btn = new MOS_WORD_TEST.Base.RJButton();
                // btn.Text = "Project " +  i.ToString() + " of " + currentProject.Questions.Count + ":";
                string prefix = "";

                if (currentProject.Questions[i - 1].MaskForReview == true)
                {
                    prefix = "🚩 ";
                }

                if (currentProject.Questions[i - 1].MaskForComplete == true)
                {
                    prefix = "✔️ ";
                }

                btn.Text = $@"{prefix + i.ToString()}";
                btn.Location = new Point(((i) * 110) + startX, 0);
                //btn.Width = 50;
                //btn.Height = 32;
                //btn.BackColor = System.Drawing.Color.DarkBlue;
                //btn.ForeColor = System.Drawing.Color.White;

                btn.AutoSize = true;
                btn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
                btn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
                btn.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
                btn.BorderColor = System.Drawing.Color.White;
                btn.BorderRadius = 10;
                btn.BorderSize = 0;
                btn.Cursor = System.Windows.Forms.Cursors.Hand;
                btn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
                btn.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
                btn.MinimumSize = new System.Drawing.Size(100, 0);
                btn.Padding = new System.Windows.Forms.Padding(2);

                btn.Tag = currentProject.Questions[i - 1];
                btn.Click += LoadQuestionContent;
                this.panelJumpQuestion.Controls.Add(btn);
                if (i == 1)
                {
                    LoadQuestionContent(btn, EventArgs.Empty);
                }
            }
            //if (currentProject.Questions.Count > 1)
            //{
            AddNextButton();
            //}
        }

        private void AddNextButton()
        {
            var btn = new MOS_WORD_TEST.Base.RJButton();
            btn.Text = ">";
            //btn.Location = new Point((currentProject.Questions.Count + 1) * 60, 0);
            btn.Location = new Point(((screen_width / 10) * 8) - 100, 0);
            //btn.BackColor = System.Drawing.Color.DarkBlue;

            btn.AutoSize = true;
            btn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            btn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            btn.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            btn.BorderColor = System.Drawing.Color.White;
            btn.BorderRadius = 10;
            btn.BorderSize = 0;
            btn.Cursor = System.Windows.Forms.Cursors.Hand;
            btn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btn.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            btn.MinimumSize = new System.Drawing.Size(100, 0);
            btn.Padding = new System.Windows.Forms.Padding(2);

            //btn.Tag = currentProject.Questions[i - 1];
            btn.Click += LoadNextQuestion;
            this.panelJumpQuestion.Controls.Add(btn);
        }

        private void AddBackButton()
        {
            var btn = new MOS_WORD_TEST.Base.RJButton();
            btn.Text = "<";
            btn.Location = new Point(0, 0);
            //btn.Width = 50;
            //btn.Height = 32;
            //btn.BackColor = System.Drawing.Color.DarkBlue;
            //btn.ForeColor = System.Drawing.Color.White;

            btn.AutoSize = true;
            btn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            btn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            btn.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            btn.BorderColor = System.Drawing.Color.White;
            btn.BorderRadius = 10;
            btn.BorderSize = 0;
            btn.Cursor = System.Windows.Forms.Cursors.Hand;
            btn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btn.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            btn.MinimumSize = new System.Drawing.Size(100, 0);
            btn.Padding = new System.Windows.Forms.Padding(2);

            //btn.Tag = currentProject.Questions[i - 1];
            btn.Click += LoadLastQuestion;
            this.panelJumpQuestion.Controls.Add(btn);
        }

        private void LoadNextQuestion(object sender, EventArgs e)
        {
            int curQuestionIndex = currentQuestion.Index;
            if (curQuestionIndex >= currentProject.Questions.Count)
            {
                return;
            }

            Button currentBT = GetLastQuestionButton();
            currentBT.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            currentBT.ForeColor = System.Drawing.Color.White;

            Button nextBT = GetLastQuestionButtonByIndex(currentQuestion.Index + 1);
            bool isNext = currentProject.Questions.Where(x => x.Index == (curQuestionIndex + 1)).Any();
            if (!isNext)
            {
                return;
            }

            currentQuestion = currentProject.Questions.Where(x => x.Index == (curQuestionIndex + 1)).FirstOrDefault();
            nextBT.BackColor = System.Drawing.Color.White;
            nextBT.ForeColor = System.Drawing.Color.Black;
            ClsQuestion questionObj = ClsListQuestion.GetQuestion(currentQuestion.QuestionNumber);
            richTextTA.Text = Language == "EN" ? questionObj.EngQuestion : questionObj.VnQuestion;
        }

        private void LoadLastQuestion(object sender, EventArgs e)
        {
            int curQuestionIndex = currentQuestion.Index;
            if (curQuestionIndex <= 1)
            {
                return;
            }

            Button currentBT = GetLastQuestionButton();
            currentBT.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            currentBT.ForeColor = System.Drawing.Color.White;

            Button lastBT = GetLastQuestionButtonByIndex(currentQuestion.Index - 1);
            bool isLast = currentProject.Questions.Where(x => x.Index == (curQuestionIndex - 1)).Any();
            if (!isLast)
            {
                return;
            }

            currentQuestion = currentProject.Questions.Where(x => x.Index == (curQuestionIndex - 1)).FirstOrDefault();
            lastBT.BackColor = System.Drawing.Color.White;
            lastBT.ForeColor = System.Drawing.Color.Black;
            ClsQuestion questionObj = ClsListQuestion.GetQuestion(currentQuestion.QuestionNumber);
            richTextTA.Text = Language == "EN" ? questionObj.EngQuestion : questionObj.VnQuestion;
        }

        private void LoadQuestionContent(object sender, EventArgs e)
        {
            Button lastBT = GetLastQuestionButton();
            if (lastBT != null)
            {
                lastBT.BackColor = Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
                lastBT.ForeColor = Color.White;
            }
            else
            {

            }

            var btn = sender as Button;
            var question = btn.Tag as Question;
            currentQuestion = question;
            btn.BackColor = Color.White;
            btn.ForeColor = Color.Black;
            ClsQuestion questionObj = ClsListQuestion.GetQuestion(question.QuestionNumber);
            richTextTA.Text = Language == "EN" ? questionObj.EngQuestion : questionObj.VnQuestion;
        }

        private Button GetLastQuestionButtonByIndex(int questionIndex)
        {

            Button btnTimDuoc = panelJumpQuestion.Controls
                          .OfType<Button>()
                            //.FirstOrDefault(b => b.Text == questionIndex.ToString());
                            .FirstOrDefault(b => b.Text.EndsWith(questionIndex.ToString()));
            return btnTimDuoc;
        }

        private Button GetLastQuestionButton()
        {
            Button btnTimDuoc = panelJumpQuestion.Controls
                            .OfType<Button>()
                            .FirstOrDefault(b => b.Text.EndsWith(currentQuestion.Index.ToString()));
            return btnTimDuoc;
        }

        private void loadcaucanhoi()
        {
            this.comboBox1.Items.Clear();
            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "zip//luu");
            if (!File.Exists(path))
                return;
            TextReader textReader = (TextReader)new StreamReader(path);
            string str;
            do
            {
                str = textReader.ReadLine();
                if (str != null)
                    this.comboBox1.Items.Add((object)str);
            }
            while (str != null);
            textReader.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //string str = this.CheckCauLon(this.currentQuestion.QuestionNumber);
            //if (!this.currentQuestion.Status)
            //{
            //    this.currentQuestion.Status = true;
            //    if (str == "True")
            //    {
            //        this.currentQuestion.Value = true;
            //    }
            //    else
            //    {
            //        this.currentQuestion.Value = false;
            //    }
            //    //this.CacCauDaCheck[this.Cau_So] = true;
            //}
            //this.richTextBox1.Text = str;
            //string str = string.Empty;
            //for(int i =0; i < currentExam.ProjectIndex.Count; i++)
            //{
            //    for(int j =0; j < currentExam.ProjectIndex[i].Questions.Count; j++)
            //    {
            //        string kq = currentExam.ProjectIndex[i].Questions[j].Value == true ? "Đúng" : "sai";
            //        str += $@"Project: {currentExam.ProjectIndex[i].ProjectName} - Câu hỏi:{currentExam.ProjectIndex[i].Questions[j].Index} - Kết quả: {kq}{Environment.NewLine}";
            //    }
            //}
            //MessageBox.Show(str);

            //Chấm điểm
            this.ChamDiem();

            Frm_GoToSummary frm = new Frm_GoToSummary();
            frm.OnChangeAnswer += Frm_OnChangeAnswer;
            frm.DataSource = this.GetDataSource();
            frm.ShowDialog();
        }

        private System.Data.DataTable GetDataSource()
        {
            var dt = new System.Data.DataTable();
            dt.Columns.Add("ProjectIndex", typeof(int));
            dt.Columns.Add("ProjectName", typeof(string));
            dt.Columns.Add("QuestionIndex", typeof(int));
            dt.Columns.Add("QuestionKey", typeof(string));
            dt.Columns.Add("QuestionContent", typeof(string));
            dt.Columns.Add("ParentKey", typeof(string));
            //dt.Columns.Add("QuestionNumber", typeof(int));
            dt.Columns.Add("MaskForComplete", typeof(string));
            dt.Columns.Add("MaskForReview", typeof(string));
            dt.Columns.Add("Status", typeof(string));
            dt.Columns.Add("Value", typeof(string));
            for (int i = 0; i < currentExam.ProjectIndex.Count; i++)
            {
                var row1 = dt.NewRow();
                row1["ProjectIndex"] = currentExam.ProjectIndex[i].ProjectIndex;
                //row1["ProjectName"] = currentExam.ProjectIndex[i].ProjectName;
                row1["ProjectName"] = $@"Project {i + 1}";
                string currentParent = $@"{currentExam.ProjectIndex[i].ProjectIndex}_0";
                row1["QuestionKey"] = currentParent;
                dt.Rows.Add(row1);

                for (int j = 0; j < currentExam.ProjectIndex[i].Questions.Count; j++)
                {
                    var row = dt.NewRow();
                    row["ProjectIndex"] = currentExam.ProjectIndex[i].ProjectIndex;
                    row["ProjectName"] = string.Empty;//currentExam.ProjectIndex[i].ProjectName;
                    string questionIndex = $@"{currentExam.ProjectIndex[i].ProjectIndex}_{currentExam.ProjectIndex[i].Questions[j].Index}";
                    row["QuestionKey"] = questionIndex;
                    row["ParentKey"] = currentParent;
                    row["QuestionIndex"] = currentExam.ProjectIndex[i].Questions[j].Index;
                    int questionNumber = currentExam.ProjectIndex[i].Questions[j].QuestionNumber;
                    row["QuestionContent"] = Language == "EN" ? ClsListQuestion.GetQuestion(questionNumber).EngQuestion : ClsListQuestion.GetQuestion(questionNumber).VnQuestion;
                    //row["QuestionNumber"] = currentExam.ProjectIndex[i].Questions[j].QuestionNumber;
                    row["MaskForComplete"] = currentExam.ProjectIndex[i].Questions[j].MaskForComplete == true ? "✔️" : string.Empty;
                    row["MaskForReview"] = currentExam.ProjectIndex[i].Questions[j].MaskForReview == true ? "🚩" : string.Empty;
                    row["Status"] = currentExam.ProjectIndex[i].Questions[j].Status == true ? "✔️" : string.Empty;
                    row["Value"] = currentExam.ProjectIndex[i].Questions[j].Value == true ? "True" : "False";
                    dt.Rows.Add(row);
                }
            }
            return dt;
        }

        private void Frm_OnChangeAnswer(ProjectEventArgs e)
        {
            bool currentExist = currentExam.ProjectIndex.Where(x => x.ProjectIndex == e.ProjectIndex).Any();
            if (!currentExist)
            {
                MessageBox.Show("Project không tồn tại. Vui lòng thử lại");
                return;
            }
            currentProject = currentExam.ProjectIndex.Where(x => x.ProjectIndex == e.ProjectIndex).FirstOrDefault();
            //this.SetUp(currentProject.ProjectIndex);
            this.pathFileOffice = currentProject.PathFileOffice;
            this.pathFileOfficeMaHoa = currentProject.PathFileOfficeMaHoa;
            OpenDocument();
            LoadProject(currentProject.ProjectIndex - 1);
            loadcaucanhoiNew(currentProject.ProjectIndex - 1);
            ChangeCurrentBT();
            bool currentQuestionExist = currentProject.Questions.Where(x => x.Index == e.QuestionIndex).Any();
            if (!currentQuestionExist)
            {
                MessageBox.Show("Question không tồn tại. Vui lòng thử lại");
                return;
            }
            currentQuestion = currentProject.Questions.Where(x => x.Index == e.QuestionIndex).FirstOrDefault();
            Button changeQuestionBT = GetLastQuestionButtonByIndex(currentQuestion.Index);
            LoadQuestionContent(changeQuestionBT, EventArgs.Empty);
        }

        private void ChangeCurrentBT()
        {
            Button currentBT = GetLastQuestionButton();
            currentBT.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(2)))), ((int)(((byte)(34)))), ((int)(((byte)(91)))));
            currentBT.ForeColor = System.Drawing.Color.White;
        }

        private void button3_Click(object sender, EventArgs e) => this.submit();

        private void submit()
        {
            if (MessageBox.Show($"Bạn có chắc chắn Submit Project {currentProject.ProjectIndex.ToString()}?", "Cảnh Báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
            {
                //bool reviewQestion = currentProject.Questions.Where(x => x.MaskForReview == true).Any();
                //if (reviewQestion)
                //{
                //    MessageBox.Show($@"Poject có question cần review nên không thể submit. Vui lòng kiểm tra câu hỏi và chuyển thành Mask for complete");
                //    return;
                //}
                this.ChamDiem();
                if (currentProject.ProjectIndex == currentExam.ProjectIndex[6].ProjectIndex)
                {
                    button2_Click(null, null);
                    return;
                    //this.next = 0;
                    //string str = "Kết quả của " + Home.thongtin + " = Đúng " + (object)this.So_Cau_Dung + "câu; Sai " + (object)this.So_Cau_Sai + " câu => Điểm: " + (1000.0 / (double)(this.So_Cau_Dung + this.So_Cau_Sai) * (double)this.So_Cau_Dung).ToString("#,##0.0");
                    //this.luuDiem(str);
                    //int num = (int)MessageBox.Show(str);
                    //this.So_Cau_Dung = 0;
                    //this.So_Cau_Sai = 0;
                }
                else
                    ++this.next;

                //this.labelProject.Text = $@"Project {(currentProject.ProjectIndex + 1).ToString()} of {currentExam.ProjectIndex.Count.ToString()}: {currentProject.ProjectName}";
                this.labelProject.Text = $@"Project {(currentProject.ProjectIndex + 1).ToString()} of {currentExam.ProjectIndex.Count.ToString()}";

                this.checkedListBox1.SetItemChecked(currentProject.ProjectIndex - 1, true);
                int nexproject = currentProject.ProjectIndex + 1;
                if (nexproject <= currentExam.ProjectIndex.Count)
                {
                    currentProject = currentExam.ProjectIndex[nexproject - 1];
                    this.checkedListBox1.SelectedIndex = nexproject - 1;
                    this.SetUp(currentProject.ProjectIndex);
                }
                else
                {
                    int num = (int)MessageBox.Show("Hết bài rồi bạn nhé");
                }
            }
            this.chotat = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show($"Bạn có muốn Restart Project {currentProject.ProjectIndex.ToString()} không? (Restart Project sẽ làm lại từ đầu)", "Cảnh Báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) != DialogResult.Yes)
                    return;
                this.chotat = true;
                for (int i = 0; i < currentProject.Questions.Count; i++)
                {
                    currentProject.Questions[i].Status = false;
                    currentProject.Questions[i].Value = false;
                    currentProject.Questions[i].MaskForComplete = false;
                    currentProject.Questions[i].MaskForReview = false;
                }

                //Remove all maskreview and markcomplete in buttons
                foreach (Control control in this.panelJumpQuestion.Controls)
                {
                    if (control is Button btn)
                    {
                        string btnText = btn.Text;
                        btnText = btnText.Replace("🚩 ", "").Replace("✔️ ", "");
                        btn.Text = btnText;
                    }
                }

                //while (this.a.Documents.Count >= 1)
                //{
                //    object SaveChanges = (object)false;
                //    object missing1 = System.Type.Missing;
                //    object missing2 = System.Type.Missing;
                //    object Index = (object)1;
                //    // ISSUE: reference to a compiler-generated method
                //    this.a.Documents[ref Index].Close(ref SaveChanges, ref missing1, ref missing2);
                //}
                object SaveChanges = (object)false;
                object missing1 = System.Type.Missing;
                object missing2 = System.Type.Missing;
                this.d.Close(ref SaveChanges, ref missing1, ref missing2);

                this.chotat = false;
                Home.DecryptFile(this.pathFileOfficeMaHoa, this.pathFileOffice);
                object pathFileOffice = (object)this.pathFileOffice;
                object missing3 = System.Type.Missing;
                object missing4 = System.Type.Missing;
                object missing5 = System.Type.Missing;
                object PasswordDocument = (object)"271565234";
                object missing6 = System.Type.Missing;
                object missing7 = System.Type.Missing;
                object missing8 = System.Type.Missing;
                object missing9 = System.Type.Missing;
                object missing10 = System.Type.Missing;
                object missing11 = System.Type.Missing;
                object missing12 = System.Type.Missing;
                object missing13 = System.Type.Missing;
                object missing14 = System.Type.Missing;
                object missing15 = System.Type.Missing;
                object XMLTransform = (object)Missing.Value;
                // ISSUE: reference to a compiler-generated method
                this.d = this.a.Documents.Open(ref pathFileOffice, ref missing3, ref missing4, ref missing5, ref PasswordDocument, ref missing6, ref missing7, ref missing8, ref missing9, ref missing10, ref missing11, ref missing12, ref missing13, ref missing14, ref missing15, ref XMLTransform);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Đóng các cửa sổ thông báo của Word trước khi Restart Project");
            }
        }

        private void TurnOffFile()
        {
            object saveChanges = (object)Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            object missing = Type.Missing;
            while (this.a.Documents.Count >= 1)
            {
                this.a.Documents[1].Close(ref saveChanges, ref missing, ref missing);
            }
        }

        private void luuDiem(string diem)
        {
            TextWriter textWriter1 = (TextWriter)new StreamWriter(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip//diem.html"), true);
            string shortDateString = DateTime.Now.ToShortDateString();
            textWriter1.WriteLine("<p style=\"text-align:center;\">" + shortDateString + ":" + diem + "</p>");
            textWriter1.Close();
            TextWriter textWriter2 = (TextWriter)new StreamWriter(Path.Combine(System.Windows.Forms.Application.StartupPath, "diem.html"), true);
            textWriter2.WriteLine("<p style=\"text-align:center;\">" + shortDateString + ":" + diem + "</p>");
            textWriter2.Close();
            string str1 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), nameof(diem));
            if (!Directory.Exists(str1))
                Directory.CreateDirectory(str1);
            string path = Path.Combine(str1, "Word.html");
            if (File.Exists(path))
            {
                TextWriter textWriter3 = (TextWriter)new StreamWriter(path, true);
                textWriter3.WriteLine("<p style=\"text-align:center;\">" + shortDateString + ":" + diem + "</p>");
                textWriter3.Close();
            }
            else
            {
                TextWriter textWriter4 = (TextWriter)new StreamWriter(path, true);
                string str2 = "<P style=\"text-align:center;color:red;font-size:160%\"> Điều Kiện để được thi lại miễn phí là phải học hành chăm chỉ, gửi điểm thường xuyên cho thầy</p><P style=\"text-align:center;\">Khi Có Điểm bạn copy Điểm này gửi cho thầy qua Website hỗ trợ https://mos360.vn</p><P style=\"text-align:center;\">================================================================== </p> ";
                textWriter4.WriteLine(str2);
                textWriter4.WriteLine("<p style=\"text-align:center;\">" + shortDateString + ":" + diem + "</p>");
                textWriter4.Close();
            }
        }

        private void ChamDiem()
        {
            for (int i = 0; i < currentExam.ProjectIndex.Count; i++)
            {
                for (int j = 0; j < currentExam.ProjectIndex[i].Questions.Count; j++)
                {
                    if (currentExam.ProjectIndex[i].ProjectIndex != currentProject.ProjectIndex)
                    {
                        continue;
                    }

                    if (!currentExam.ProjectIndex[i].Questions[j].Status == true)
                    {
                        int cau = currentExam.ProjectIndex[i].Questions[j].QuestionNumber;
                        currentExam.ProjectIndex[i].Questions[j].Status = true;
                        if (this.CheckCauLon(cau) == "True")
                        {
                            currentExam.ProjectIndex[i].Questions[j].Value = true;
                            ++this.So_Cau_Dung;
                        }
                        else
                        {
                            currentExam.ProjectIndex[i].Questions[j].Value = false;
                            ++this.So_Cau_Sai;
                        }
                    }
                }
            }

            //this.So_Cau_Dung = 0;
            //this.So_Cau_Sai = 0;
            //this.So_Cau_Dung = currentExam.ProjectIndex.SelectMany(p => p.Questions).Count(x => x.Status == true && x.Value == true);
            //this.So_Cau_Sai = 35 - this.So_Cau_Dung;
            //int a = currentExam.ProjectIndex.SelectMany(p => p.Questions).Count(x => x.Status == true && x.Value == false);
            //for (int cau = 0; cau < this.Tong_So_Cau; ++cau)
            //{
            //    if (!this.CacCauDaCheck[cau])
            //    {
            //        currentQuestion.Status = true;
            //        if (this.CheckCauLon(cau) == "True")
            //        {
            //            ++this.So_Cau_Dung;
            //            currentQuestion.Value = true;
            //        }
            //        else
            //        {
            //            currentQuestion.Value = false;
            //            ++this.So_Cau_Sai;
            //        }
            //    }
            //}
            //this.chotat = true;
            //object obj = (object)WdSaveOptions.wdDoNotSaveChanges;
            //while (this.a.Documents.Count >= 1)
            //{
            //    object SaveChanges = obj;
            //    object missing1 = System.Type.Missing;
            //    object missing2 = System.Type.Missing;
            //    object Index = (object)1;
            //    // ISSUE: reference to a compiler-generated method
            //    this.a.Documents[ref Index].Close(ref SaveChanges, ref missing1, ref missing2);
            //}
            //this.chotat = false;
        }
        private void SetUpNew(int Index)
        {

        }

        private void SetUp(int Index)
        {
            string[] directories = Directory.GetDirectories(Path.Combine(System.Windows.Forms.Application.StartupPath, $@"Zip\\Exam\\Exam{currentExam.ExamIndex}"));
            this.works = directories.Length;
            Array.Sort<string>(directories, (IComparer<string>)StringComparer.InvariantCulture);
            this.workIndex = Index - 1;
            this.pathWork = directories[this.workIndex];
            this.pathRun = Path.Combine(System.Windows.Forms.Application.StartupPath, "Zip\\Tam");
            //this.ImageFile = Directory.GetFiles(Path.Combine(this.pathWork, "Test"));
            //this.checkedListBox1.Items.Clear();
            //for (int index = 0; index < this.ImageFile.Length; ++index)
            //    this.checkedListBox1.Items.Add((object)(index + 1).ToString());
            if (string.IsNullOrEmpty(currentProject.PathFileOfficeMaHoa))
            {
                this.pathFileOfficeMaHoa = Path.Combine(this.pathWork, "Source\\Main");
                this.pathFileOfficeMaHoa = Directory.GetFiles(this.pathFileOfficeMaHoa)[0];
                currentProject.PathFileOfficeMaHoa = this.pathFileOfficeMaHoa;
            }
            else
            {
                this.pathFileOfficeMaHoa = currentProject.PathFileOfficeMaHoa;
            }

            if (string.IsNullOrEmpty(currentProject.PathFileOffice))
            {
                this.pathFileOffice = Path.Combine(System.Windows.Forms.Application.StartupPath, "Zip\\Tam\\A\\");
                this.pathFileOffice = Path.Combine(this.pathFileOffice, Path.GetFileName(this.pathFileOfficeMaHoa));
                currentProject.PathFileOffice = this.pathFileOffice;
            }
            else
            {
                this.pathFileOffice = currentProject.PathFileOffice;
            }
            //this.pathFileOffice = Path.Combine(System.Windows.Forms.Application.StartupPath, "Zip\\Tam\\A\\");
            //this.pathFileOffice = Path.Combine(this.pathFileOffice, Path.GetFileName(this.pathFileOfficeMaHoa));
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            string path = Path.Combine(this.pathWork, "Source\\Main");
            try
            {
                foreach (string file in Directory.GetFiles(path))
                    File.Copy(file, Path.Combine(folderPath, Path.GetFileName(file)), true);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Thư mục Document của bạn không truy cập được, các file dữ liệu cần cho bài làm sẽ ở thư mục Data" + ex.Message);
                foreach (string file in Directory.GetFiles(path))
                    File.Copy(file, Path.Combine(System.Windows.Forms.Application.StartupPath + "\\DATA", Path.GetFileName(file)), true);
            }
            //this.CacCauDaCheck = new bool[this.Tong_So_Cau];
            //for (int index = 0; index < this.Tong_So_Cau; ++index)
            //    this.CacCauDaCheck[index] = false;

            TextReader textReader = (TextReader)new StreamReader(Path.Combine(this.pathWork, "a.txt"));
            this.lsViTri = new List<int>();
            for (string s = textReader.ReadLine(); s != null; s = textReader.ReadLine())
                this.lsViTri.Add(int.Parse(s));
            textReader.Close();
            this.Cau_So = 0;
            //this.checkedListBox1.SelectedIndex = this.Cau_So;
            if (!currentProject.DocumentOpened)
            {
                Home.DecryptFile(this.pathFileOfficeMaHoa, this.pathFileOffice);
                currentProject.DocumentOpened = true;
            }
            OpenDocument();
            //for (int index = 0; index < this.CacCauDaCheck.Length; ++index)
            //    this.CacCauDaCheck[index] = false;
        }

        private void OpenDocument()
        {
            object pathFileOffice = (object)this.pathFileOffice;
            object missing1 = System.Type.Missing;
            object missing2 = System.Type.Missing;
            object missing3 = System.Type.Missing;
            object PasswordDocument = (object)"271565234";
            object missing4 = System.Type.Missing;
            object missing5 = System.Type.Missing;
            object missing6 = System.Type.Missing;
            object missing7 = System.Type.Missing;
            object missing8 = System.Type.Missing;
            object missing9 = System.Type.Missing;
            object missing10 = System.Type.Missing;
            object missing11 = System.Type.Missing;
            object missing12 = System.Type.Missing;
            object missing13 = System.Type.Missing;
            object XMLTransform = (object)Missing.Value;
            // ISSUE: reference to a compiler-generated method
            this.d = this.a.Documents.Open(ref pathFileOffice, ref missing1, ref missing2, ref missing3, ref PasswordDocument, ref missing4, ref missing5, ref missing6, ref missing7, ref missing8, ref missing9, ref missing10, ref missing11, ref missing12, ref missing13, ref XMLTransform);
            this.soLanReSet = 0;
            if (currentProject.DocumentOpened)
            {
                this.d.Activate();
                IntPtr hwnd = (IntPtr)a.ActiveWindow.Hwnd;
                SetForegroundWindow(hwnd);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.chotat = true;
            while (this.a.Documents.Count >= 1)
            {
                object SaveChanges = (object)false;
                object missing1 = System.Type.Missing;
                object missing2 = System.Type.Missing;
                object Index = (object)1;
                // ISSUE: reference to a compiler-generated method
                this.a.Documents[ref Index].Close(ref SaveChanges, ref missing1, ref missing2);
            }
            object SaveChanges1 = (object)Missing.Value;
            object OriginalFormat = (object)Missing.Value;
            object RouteDocument = (object)Missing.Value;
            // ISSUE: reference to a compiler-generated method
            this.a.Quit(ref SaveChanges1, ref OriginalFormat, ref RouteDocument);
            this.chotat = false;
            this.Close();
        }

        private void timer1_Tick(object sender, EventArgs e) => this.label1.Text = (this.timeStrart - DateTime.Now).ToString().Substring(3, 5);

        private void LoadCau(int Cau)
        {
            this.checkedListBox1.SelectedIndex = Cau;
            foreach (string file in Directory.GetFiles(this.pathRun))
                File.Delete(file);
            this.ImageFile[this.checkedListBox1.SelectedIndex].Replace(Path.Combine(this.pathWork, "Test"), this.pathRun);
            Home.DecryptFile(this.pathFileOfficeMaHoa, this.pathFileOffice);
            object pathFileOffice = (object)this.pathFileOffice;
            object missing1 = System.Type.Missing;
            object missing2 = System.Type.Missing;
            object missing3 = System.Type.Missing;
            object PasswordDocument = (object)"271565234";
            object missing4 = System.Type.Missing;
            object missing5 = System.Type.Missing;
            object missing6 = System.Type.Missing;
            object missing7 = System.Type.Missing;
            object missing8 = System.Type.Missing;
            object missing9 = System.Type.Missing;
            object missing10 = System.Type.Missing;
            object missing11 = System.Type.Missing;
            object missing12 = System.Type.Missing;
            object missing13 = System.Type.Missing;
            object XMLTransform = (object)Missing.Value;
            // ISSUE: reference to a compiler-generated method
            this.d = this.a.Documents.Open(ref pathFileOffice, ref missing1, ref missing2, ref missing3, ref PasswordDocument, ref missing4, ref missing5, ref missing6, ref missing7, ref missing8, ref missing9, ref missing10, ref missing11, ref missing12, ref missing13, ref XMLTransform);
        }

        private string CheckCauLon(int cau)
        {
            try
            {
                IQuestionCheck questionCheck = new QuestionCheck();
                //return questionCheck.CheckAnswer(currentQuestion.QuestionNumber, this.a, this.d);
                return questionCheck.CheckAnswer(cau, this.a, this.d);
            }
            catch (Exception ex)
            {
                return "False (Something wrong)";
            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.Items.Count > 10)
            {
                int num = (int)MessageBox.Show("Quá 10 câu không biết rồi bạn gọi thầy đi");
            }
            else
            {
                TextWriter textWriter = (TextWriter)new StreamWriter(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\luu"), true);
                textWriter.WriteLine((Program.Lessons[this.next] + 1).ToString() + "|" + (object)(this.checkedListBox1.SelectedIndex + 1));
                textWriter.Close();
                this.loadcaucanhoi();
            }
        }

        private void pictureBox2_MouseHover(object sender, EventArgs e) => this.panel2.Focus();

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                this.next = int.Parse(this.textBox2.Text) - 2;
                if (this.next >= -1)
                {
                    this.submit();
                }
                else
                {
                    int num = (int)MessageBox.Show("so lon hon 0");
                    this.textBox2.Focus();
                }
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("nhap so");
                this.textBox2.Focus();
            }
        }

        private void buttonHelp_Click(object sender, EventArgs e)
        {
            string str = Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\tam\\1.avi");
            string fileNameSource = Path.Combine(this.pathWork, "help\\" + (object)(this.checkedListBox1.SelectedIndex + 1));
            try
            {
                Home.DecryptFile(fileNameSource, str);
                Process.Start(str);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Tắt Video help trước khi mở help mới |" + ex.Message);
            }
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                this.a.WindowState = WdWindowState.wdWindowStateNormal;
                this.a.Top = 0;
                this.a.Left = -10;
                // ISSUE: reference to a compiler-generated method
                this.a.Resize(this.screen_width, this.screen_height * 3 / 5);
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("Bạn đã tắt MS Word! Tắt chương trình và làm lại");
            }
        }

        private void buttonxhdh_Click(object sender, EventArgs e)
        {
            if (this.buttonxhdh.Text == "HDSD")
            {
                this.buttonxhdh.Text = "Ẩn Taskbar";
                this.TopMost = false;
                try
                {
                    Process.Start("https://go.mos360.vn/mostesthdsd");
                }
                catch (Exception ex)
                {
                    int num = (int)MessageBox.Show("Vui lòng kết nối mạng để xem Hướng dẫn sử dụng phần mềm " + ex.Message);
                }
            }
            else
            {
                this.buttonxhdh.Text = "HDSD";
                this.TopMost = true;
            }
        }

        private void buttonShowHide_Click(object sender, EventArgs e)
        {
            if (this.TopMost)
            {
                this.TopMost = false;
                this.buttonShowHide.Text = "Tự Làm";
                //Process.Start("https://meet.google.com/uay-ywnp-xzs");
            }
            else
            {
                this.TopMost = true;
                this.buttonShowHide.Text = "Gọi Thầy";
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (this.comboBox1.Items.Count == 0)
            {
                int num1 = (int)MessageBox.Show("Bạn không có câu cần hỏi");
            }
            else if (this.comboBox1.Text == "")
            {
                int num2 = (int)MessageBox.Show("Chọn Project");
            }
            else
            {
                string text = this.comboBox1.Text;
                this.comboBox1.Items.RemoveAt(this.comboBox1.SelectedIndex);
                TextWriter textWriter = (TextWriter)new StreamWriter(Path.Combine(System.Windows.Forms.Application.StartupPath, "zip\\Luu"), false);
                for (int index = 0; index < this.comboBox1.Items.Count; ++index)
                    textWriter.WriteLine(this.comboBox1.Items[index].ToString());
                textWriter.Close();
                this.submit(int.Parse(text.Split('|')[0]) - 1);
            }
        }

        public void submit(int pro)
        {
            if (MessageBox.Show("Bạn có chắc nộp bài?", "Cảnh Báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
            {
                for (int i = 0; i < currentExam.ProjectIndex.Count; i++)
                {
                    for (int j = 0; j < currentExam.ProjectIndex[i].Questions.Count; j++)
                    {
                        if (!currentExam.ProjectIndex[i].Questions[j].Status == true)
                        {
                            int cau = currentExam.ProjectIndex[i].Questions[j].QuestionNumber;
                            currentExam.ProjectIndex[i].Questions[j].Status = true;
                            if (this.CheckCauLon(cau) == "True")
                            {
                                currentExam.ProjectIndex[i].Questions[j].Value = true;
                                ++this.So_Cau_Dung;
                            }
                            else
                            {
                                currentExam.ProjectIndex[i].Questions[j].Value = false;
                                ++this.So_Cau_Sai;
                            }
                        }
                    }
                }
                //for (int cau = 0; cau < this.Tong_So_Cau; ++cau)
                //{
                //    if (!this.CacCauDaCheck[cau])
                //    {
                //        if (this.CheckCauLon(cau) == "True")
                //            ++this.So_Cau_Dung;
                //        else
                //            ++this.So_Cau_Sai;
                //    }
                //}
                this.Cau_So = 0;
                this.checkedListBox1.SelectedIndex = 0;
                this.pathReset = Path.Combine(this.pathWork, "Source\\Sub\\" + this.Cau_So.ToString());
                this.chotat = true;
                while (this.a.Documents.Count >= 1)
                {
                    object SaveChanges = (object)false;
                    object missing1 = System.Type.Missing;
                    object missing2 = System.Type.Missing;
                    object Index = (object)1;
                    // ISSUE: reference to a compiler-generated method
                    this.a.Documents[ref Index].Close(ref SaveChanges, ref missing1, ref missing2);
                }
                this.chotat = false;
                this.SetUp(pro);
                this.soLanReSet = 0;
                for (int index = 0; index < this.CacCauDaCheck.Length; ++index)
                    this.CacCauDaCheck[index] = false;
            }
            this.chotat = false;
            this.Cau_So = 0;
            this.checkedListBox1.SelectedIndex = 0;
        }

        //private void pictureBox1_MouseHover(object sender, EventArgs e) => this.panel1.Focus();

        private void btnMarkComplete_Click(object sender, EventArgs e)
        {
            //currentQuestion.MaskForComplete = true;
            Button currentBT = GetLastQuestionButton();

            if (currentBT.Text.StartsWith("🚩"))
            {
                currentBT.Text = $@"✔️ {currentQuestion.Index.ToString()}";
                currentQuestion.MaskForComplete = true;
                currentQuestion.MaskForReview = false;
            }
            else
            {
                if (currentBT.Text == currentQuestion.Index.ToString())
                {
                    currentBT.Text = $@"✔️ {currentQuestion.Index.ToString()}";
                    currentQuestion.MaskForComplete = true;
                    currentQuestion.MaskForReview = false;
                }
                else
                {
                    currentBT.Text = $@"{currentQuestion.Index.ToString()}";
                    currentQuestion.MaskForComplete = false;
                    currentQuestion.MaskForReview = false;
                }
            }
        }

        private void btnMaskReview_Click(object sender, EventArgs e)
        {
            //currentQuestion.MaskForReview = true;
            Button currentBT = GetLastQuestionButton();

            if (currentBT.Text.StartsWith("✔️"))
            {
                currentBT.Text = $@"🚩 {currentQuestion.Index.ToString()}";
                currentQuestion.MaskForReview = true;
                currentQuestion.MaskForComplete = false;
            }
            else
            {
                if (currentBT.Text == currentQuestion.Index.ToString())
                {
                    currentBT.Text = $@"🚩 {currentQuestion.Index.ToString()}";
                    currentQuestion.MaskForReview = true;
                    currentQuestion.MaskForComplete = false;
                }
                else
                {
                    currentBT.Text = $@"{currentQuestion.Index.ToString()}";
                    currentQuestion.MaskForReview = false;
                    currentQuestion.MaskForComplete = false;
                }
            }
        }
    }
}
