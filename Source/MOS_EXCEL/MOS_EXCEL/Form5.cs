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

namespace MOS_EXCEL_LEARN
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }

        private void Form5_New_Load(object sender, EventArgs e)
        {
            this.buttonExit.Location = new Point(this.Width - this.buttonExit.Width, this.Height - this.buttonExit.Height);
            this.richTextBox1.Width = this.Width;
            this.richTextBox1.Height = this.Height - this.buttonNext.Height;
            this.richTextBox1.Location = new Point(0, 0);
            this.buttonNext.Location = new Point((this.Width - this.buttonNext.Width) / 2, this.Height - this.buttonNext.Height);
            if (!File.Exists("d:\\config\\cs3.rtf"))
                return;
            this.richTextBox1.LoadFile("d:\\config\\cs3.rtf");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Program.key = "";
            this.Close();
        }

        private void buttonNext_Click(object sender, EventArgs e)
        {
            Program.key = "Next";
            this.Close();
        }
    }
}
