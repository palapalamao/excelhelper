namespace OverTimeStatistics
{
    partial class FormOverTimeList
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.OrinFilelist = new System.Windows.Forms.TextBox();
            this.buttonopenfolder = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.Btn_StartStatistics = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button5 = new System.Windows.Forms.Button();
            this.textBoxori = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.textBoxmodify = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // OrinFilelist
            // 
            this.OrinFilelist.Location = new System.Drawing.Point(15, 48);
            this.OrinFilelist.Multiline = true;
            this.OrinFilelist.Name = "OrinFilelist";
            this.OrinFilelist.ReadOnly = true;
            this.OrinFilelist.Size = new System.Drawing.Size(227, 50);
            this.OrinFilelist.TabIndex = 29;
            // 
            // buttonopenfolder
            // 
            this.buttonopenfolder.Location = new System.Drawing.Point(248, 48);
            this.buttonopenfolder.Name = "buttonopenfolder";
            this.buttonopenfolder.Size = new System.Drawing.Size(40, 50);
            this.buttonopenfolder.TabIndex = 28;
            this.buttonopenfolder.Text = "……";
            this.buttonopenfolder.UseVisualStyleBackColor = true;
            this.buttonopenfolder.Click += new System.EventHandler(this.buttonopenfolder_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, -16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 12);
            this.label1.TabIndex = 27;
            this.label1.Text = "1.请选择要通知的名单:";
            // 
            // Btn_StartStatistics
            // 
            this.Btn_StartStatistics.Location = new System.Drawing.Point(111, 115);
            this.Btn_StartStatistics.Margin = new System.Windows.Forms.Padding(2);
            this.Btn_StartStatistics.Name = "Btn_StartStatistics";
            this.Btn_StartStatistics.Size = new System.Drawing.Size(81, 33);
            this.Btn_StartStatistics.TabIndex = 26;
            this.Btn_StartStatistics.Text = "开始";
            this.Btn_StartStatistics.UseVisualStyleBackColor = true;
            this.Btn_StartStatistics.Click += new System.EventHandler(this.Btn_StartStatistics_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(16, 21);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(149, 12);
            this.label4.TabIndex = 34;
            this.label4.Text = "1.请选择需要拆分的文件们";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.Btn_StartStatistics);
            this.groupBox1.Controls.Add(this.buttonopenfolder);
            this.groupBox1.Controls.Add(this.OrinFilelist);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Location = new System.Drawing.Point(21, 14);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(316, 411);
            this.groupBox1.TabIndex = 51;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "1.拆表";
            // 
            // toolTip1
            // 
            this.toolTip1.Tag = "ssssss";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Controls.Add(this.textBox3);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Location = new System.Drawing.Point(350, 14);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox2.Size = new System.Drawing.Size(414, 202);
            this.groupBox2.TabIndex = 52;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "2.合表";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(144, 136);
            this.button1.Margin = new System.Windows.Forms.Padding(2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(81, 33);
            this.button1.TabIndex = 26;
            this.button1.Text = "开始";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(277, 48);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(40, 50);
            this.button2.TabIndex = 28;
            this.button2.Text = "……";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(44, 48);
            this.textBox3.Multiline = true;
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(227, 50);
            this.textBox3.TabIndex = 29;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(16, 21);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(383, 12);
            this.label5.TabIndex = 34;
            this.label5.Text = "1.请选择需要合并的文件们:（按照每个excel的sheet页名字进行合并）";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button5);
            this.groupBox3.Controls.Add(this.textBoxori);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.button3);
            this.groupBox3.Controls.Add(this.button4);
            this.groupBox3.Controls.Add(this.textBoxmodify);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Location = new System.Drawing.Point(350, 246);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox3.Size = new System.Drawing.Size(414, 178);
            this.groupBox3.TabIndex = 53;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "2.合并到原始表";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(277, 105);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(40, 19);
            this.button5.TabIndex = 35;
            this.button5.Text = "……";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // textBoxori
            // 
            this.textBoxori.Location = new System.Drawing.Point(44, 105);
            this.textBoxori.Multiline = true;
            this.textBoxori.Name = "textBoxori";
            this.textBoxori.ReadOnly = true;
            this.textBoxori.Size = new System.Drawing.Size(227, 20);
            this.textBoxori.TabIndex = 36;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(16, 78);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(89, 12);
            this.label6.TabIndex = 37;
            this.label6.Text = "2.请选择原始表";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(144, 141);
            this.button3.Margin = new System.Windows.Forms.Padding(2);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(81, 33);
            this.button3.TabIndex = 26;
            this.button3.Text = "开始";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(277, 48);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(40, 19);
            this.button4.TabIndex = 28;
            this.button4.Text = "……";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // textBoxmodify
            // 
            this.textBoxmodify.Location = new System.Drawing.Point(44, 48);
            this.textBoxmodify.Multiline = true;
            this.textBoxmodify.Name = "textBoxmodify";
            this.textBoxmodify.ReadOnly = true;
            this.textBoxmodify.Size = new System.Drawing.Size(227, 20);
            this.textBoxmodify.TabIndex = 29;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 21);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(125, 12);
            this.label3.TabIndex = 34;
            this.label3.Text = "1.请选择修改后的文件";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(19, 523);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 12);
            this.label2.TabIndex = 54;
            this.label2.Text = "Copyright by Allan";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(17, 548);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(215, 12);
            this.label7.TabIndex = 55;
            this.label7.Text = "遇到问题欢迎沟通，1160744812@qq.com";
            this.label7.Click += new System.EventHandler(this.label7_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(494, 552);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(341, 12);
            this.label8.TabIndex = 56;
            this.label8.Text = "做软件不易，欢迎好心的哥哥姐姐大爷大妈，爷爷奶奶资助点儿";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::OverTimeStatistics.Properties.Resources.QQ截图20171117211842;
            this.pictureBox1.Location = new System.Drawing.Point(618, 429);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(146, 120);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 57;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // FormOverTimeList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "FormOverTimeList";
            this.Size = new System.Drawing.Size(848, 587);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox OrinFilelist;
        private System.Windows.Forms.Button buttonopenfolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button Btn_StartStatistics;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.TextBox textBoxori;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox textBoxmodify;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}