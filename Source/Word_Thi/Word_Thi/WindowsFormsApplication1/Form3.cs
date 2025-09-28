using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MOS_WORD_TEST
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
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

        private void Form3_Load(object sender, EventArgs e)
        {
            this.buttonExit.Location = new Point(this.Width - this.buttonExit.Width, this.Height - this.buttonExit.Height);
            this.richTextBox1.Width = this.Width;
            this.richTextBox1.Height = this.Height - this.buttonNext.Height;
            this.richTextBox1.Location = new Point(0, 0);
            this.buttonNext.Location = new Point((this.Width - this.buttonNext.Width) / 2, this.Height - this.buttonNext.Height);
            if (!File.Exists("d:\\config\\cs1.rtf"))
                return;
            this.richTextBox1.LoadFile("d:\\config\\cs1.rtf");
        }
    }
}
