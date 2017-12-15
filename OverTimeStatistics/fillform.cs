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
    public partial class Formfilldata : Form
    {
        public string splitcolums = "";
        public string readcolums = "";
        public int start_linenumber = 0;
        public string select_sheetname = "";

        public Formfilldata()
        {
            InitializeComponent();
        }
        public void set_sheetnames(List<string> sheetnames)
        {
            comboBox_sheetnames.Items.Clear();
            foreach (string item in sheetnames)
            {
                comboBox_sheetnames.Items.Add(item);
            }

            comboBox_sheetnames.SelectedIndex = 0;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            splitcolums = cloumIDtextsplit.Text.Trim();
            readcolums = textBoxreadcolum.Text.Trim();
            start_linenumber = int.Parse(this.text_startline.Text);
            select_sheetname = comboBox_sheetnames.SelectedItem.ToString();
            this.Dispose();
        }
    }
}
