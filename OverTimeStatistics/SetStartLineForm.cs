using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OverTimeStatistics
{
    public partial class SetStartLineForm : Form
    {
        public string sheetname = "";
        public int startline = 1;
        public SetStartLineForm()
        {
            InitializeComponent();
        }

        public void set_sheet_name(string tsheetname)
        {
            textBox_sheetname.Text = tsheetname;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            startline = int.Parse(text_startline.Text);
            sheetname = textBox_sheetname.Text.ToString();
            this.Dispose();
        }
    }
}
