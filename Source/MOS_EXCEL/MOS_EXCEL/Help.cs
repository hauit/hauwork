using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MOS_EXCEL_LEARN
{
    public partial class Help : Form
    {
        public Help()
        {
            InitializeComponent();
            this.SizeChanged += new EventHandler(this.Help_SizeChanged);
        }

        private void Help_Load(object sender, EventArgs e)
        {
            this.Width = Screen.PrimaryScreen.Bounds.Width / 2;
            this.Height = Screen.PrimaryScreen.Bounds.Height / 3;
            this.richTextBox1.Location = new Point(0, 0);
            this.richTextBox1.Width = this.Width - 15;
            this.richTextBox1.Height = this.Height - 2 * this.button1.Height;
            Rectangle bounds = Screen.PrimaryScreen.Bounds;
            int x = bounds.Width - this.Width;
            bounds = Screen.PrimaryScreen.Bounds;
            int y = bounds.Height - this.Height;
            this.Location = new Point(x, y);
        }

        private void Help_SizeChanged(object sender, EventArgs e)
        {
            this.richTextBox1.Width = this.Width - 15;
            this.richTextBox1.Height = this.Height - 2 * this.button1.Height;
        }

        public void refresh()
        {
        }

        private void button1_Click(object sender, EventArgs e) => this.Hide();
    }
}
