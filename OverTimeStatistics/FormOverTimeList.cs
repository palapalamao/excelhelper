using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OverTimeStatistics.OverTimeListDetail;

namespace OverTimeStatistics
{
    public partial class FormOverTimeList : UserControl
    {
        OverTimeListClass otc = null;
        public FormOverTimeList()
        {
            InitializeComponent();
            otc = new OverTimeListClass(AppDomain.CurrentDomain.BaseDirectory + ConfigFile.FileName);
            //textBox1.Text = otc.Deplist;
            //TargetFileName.Text = otc.ImportFileOri;
            //StartDate.Text = otc.CurDate;

            // textBox3point.Text = otc.Threepointsalary;
            // textBox2point.Text = otc.Twopointsalary;
            //textBox1Point5.Text = otc.Onepointfivesalary;
        }

        private void Btn_StartStatistics_Click(object sender, EventArgs e)
        {
            try
            {
                //otc.SaveIniData(TargetFileName.Text, textBoxUnqinueName.Text, StartDate.Text, textBox3point.Text, textBox2point.Text,textBox1Point5.Text);
                //otc.GetIniData(AppDomain.CurrentDomain.BaseDirectory + ConfigFile.FileName);

                otc.StartReadThread();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void buttonopenfolder_Click(object sender, EventArgs e)
        {
            OpenFileDialog opg = new OpenFileDialog();
            opg.Multiselect = true;
            if (opg.ShowDialog() == DialogResult.OK)
            {
                otc = new OverTimeListClass(AppDomain.CurrentDomain.BaseDirectory + ConfigFile.FileName);
                //textBox1.Text = otc.Deplist;
                foreach (var item in opg.FileNames)
                {
                    OrinFilelist.Text += item + ",";
                    otc.orifilelist.Add(item);
                }

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opg = new OpenFileDialog();
            opg.Filter = "Support Files (*.xls)|*.xls|All files (*.*)|*.*";
            opg.Multiselect = false;
            if (opg.ShowDialog() == DialogResult.OK)
            {
                //textBoxUnqinueName.Text = opg.FileName;

            }
        }

        private void textBox1Point5_Enter(object sender, EventArgs e)
        {
            //textBox1Point5.Text = "";

            //try
            //{
            //    string[] yearandmonth = StartDate.Text.Split('.');
            //    DateTime tempdt = new DateTime(Convert.ToInt32(yearandmonth[0]), Convert.ToInt32(yearandmonth[1]), 1);

            //    List<string> dttotal = new List<string>();
            //    string[] dt3 = textBox3point.Text.Split(',');
            //    string[] dt2 = textBox2point.Text.Split(','); 

            //    for (int i = 1; i <= DateTime.DaysInMonth(int.Parse(yearandmonth[0]), int.Parse(yearandmonth[1])); i++)
            //    {
            //        dttotal.Add(i.ToString());
            //    }
            //    foreach (string item in dt3)
            //    {
            //        dttotal.Remove(item);
            //    }
            //    foreach (string item in dt2)
            //    {
            //        dttotal.Remove(item);
            //    }

            //    for (int i = dttotal.Count-1; i >= 0; i--)
            //    {
            //        textBox1Point5.Text = dttotal[i] + "," + textBox1Point5.Text;
            //    }
            //    textBox1Point5.Text = textBox1Point5.Text.TrimEnd(',');

            //}

            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

            //essageBox.Show("get fouces");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog opg = new OpenFileDialog();
            opg.Filter = "Support Files (*.xls)|*.xls|All files (*.*)|*.*";
            opg.Multiselect = false;
            if (opg.ShowDialog() == DialogResult.OK)
            {
                otc.WithoutovertimeFileName = opg.FileName;
                //textBox1.Text = otc.WithoutovertimeFileName;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {

            otc.StartGeneratorList();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //OpenFileDialog opg = new OpenFileDialog();
            //opg.Filter = "Support Files (*.xls)|*.xls|(*.xlsx)|*.xlsx|All files (*.*)|*.*";
            //opg.Multiselect = false;
            //if (opg.ShowDialog() == DialogResult.OK)
            //{
            //    otc.ResultFileName = opg.FileName;
            //    textBox10.Text = opg.FileName;
            //}
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog opg = new OpenFileDialog();
            opg.Filter = "Support Files (*.xls)|*.xls|(*.xlsx)|*.xlsx|All files (*.*)|*.*";
            opg.Multiselect = false;
            if (opg.ShowDialog() == DialogResult.OK)
            {
                otc.ResultFileName = opg.FileName;

            }
        }

        private void textBoxYearStart_MouseEnter(object sender, EventArgs e)
        {
            //
        }

        private void textBoxYearEndDate_MouseEnter(object sender, EventArgs e)
        {
            // textBoxYearEndDate.Text = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString();
        }

        private void Btn_YearStartCal_Click(object sender, EventArgs e)
        {
            otc.StartGeneratorYearStatics();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog opg = new OpenFileDialog();
            opg.Multiselect = true;
            string tfilenames = "";
            if (opg.ShowDialog() == DialogResult.OK)
            {
                otc = new OverTimeListClass(AppDomain.CurrentDomain.BaseDirectory + ConfigFile.FileName);
                //textBox1.Text = otc.Deplist;
                foreach (var item in opg.FileNames)
                {
                    tfilenames += item + ",";
                    otc.orifilelist.Add(item);
                }
                textBox3.Text = tfilenames;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                //otc.SaveIniData(TargetFileName.Text, textBoxUnqinueName.Text, StartDate.Text, textBox3point.Text, textBox2point.Text,textBox1Point5.Text);
                //otc.GetIniData(AppDomain.CurrentDomain.BaseDirectory + ConfigFile.FileName);

                otc.MergeFileThread();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog opg = new OpenFileDialog();
            opg.Multiselect = false;
            string tfilenames = "";
            if (opg.ShowDialog() == DialogResult.OK)
            { 
                otc.modifyFileName = opg.FileName;
                textBoxmodify.Text = opg.FileName;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog opg = new OpenFileDialog();
            opg.Multiselect = false;
            string tfilenames = "";
            if (opg.ShowDialog() == DialogResult.OK)
            {
                otc.oriFileName = opg.FileName;
                textBoxori.Text = opg.FileName;
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            otc = new OverTimeListClass(AppDomain.CurrentDomain.BaseDirectory + ConfigFile.FileName);
            otc.oriFileName = textBoxori.Text;
            otc.modifyFileName = textBoxmodify.Text;
            otc.StartGeneratorList();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
