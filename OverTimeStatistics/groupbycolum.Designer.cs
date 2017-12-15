namespace OverTimeStatistics
{
    partial class cloumgroup
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
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cloumIDtext = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.text_startline = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox_sheetnames = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(42, 260);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 25);
            this.button1.TabIndex = 0;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(39, 176);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(178, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "请输入按哪列进行分组（如“K”）";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // cloumIDtext
            // 
            this.cloumIDtext.Location = new System.Drawing.Point(41, 213);
            this.cloumIDtext.Name = "cloumIDtext";
            this.cloumIDtext.Size = new System.Drawing.Size(75, 20);
            this.cloumIDtext.TabIndex = 2;
            this.cloumIDtext.Text = "F";
            this.cloumIDtext.TextChanged += new System.EventHandler(this.cloumIDtext_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(38, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(191, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "选择当前需要处理的Sheet页名称：";
            // 
            // text_startline
            // 
            this.text_startline.Location = new System.Drawing.Point(41, 134);
            this.text_startline.Name = "text_startline";
            this.text_startline.Size = new System.Drawing.Size(329, 20);
            this.text_startline.TabIndex = 9;
            this.text_startline.Text = "1";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(39, 105);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(165, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "请输入列所在的行数（如“1”）";
            // 
            // comboBox_sheetnames
            // 
            this.comboBox_sheetnames.FormattingEnabled = true;
            this.comboBox_sheetnames.Location = new System.Drawing.Point(41, 65);
            this.comboBox_sheetnames.Name = "comboBox_sheetnames";
            this.comboBox_sheetnames.Size = new System.Drawing.Size(329, 21);
            this.comboBox_sheetnames.TabIndex = 11;
            // 
            // cloumgroup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(458, 350);
            this.Controls.Add(this.comboBox_sheetnames);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.text_startline);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cloumIDtext);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "cloumgroup";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "分组列的编号";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox cloumIDtext;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox text_startline;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox_sheetnames;
    }
}