namespace OverTimeStatistics
{
    partial class Formfilldata
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
            this.cloumIDtextsplit = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.textBoxreadcolum = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBox_sheetnames = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.text_startline = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // cloumIDtextsplit
            // 
            this.cloumIDtextsplit.Location = new System.Drawing.Point(30, 225);
            this.cloumIDtextsplit.Name = "cloumIDtextsplit";
            this.cloumIDtextsplit.Size = new System.Drawing.Size(100, 20);
            this.cloumIDtextsplit.TabIndex = 5;
            this.cloumIDtextsplit.Text = "B,C,F";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 188);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(190, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "请输入以哪列为比较基准（如“B”）";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(30, 366);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 25);
            this.button1.TabIndex = 3;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBoxreadcolum
            // 
            this.textBoxreadcolum.Location = new System.Drawing.Point(29, 308);
            this.textBoxreadcolum.Name = "textBoxreadcolum";
            this.textBoxreadcolum.Size = new System.Drawing.Size(100, 20);
            this.textBoxreadcolum.TabIndex = 7;
            this.textBoxreadcolum.Text = "Q,R,S";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 271);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(234, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "请输入需要填充哪几列的数据（如“X,Y,Z”）";
            // 
            // comboBox_sheetnames
            // 
            this.comboBox_sheetnames.FormattingEnabled = true;
            this.comboBox_sheetnames.Location = new System.Drawing.Point(30, 68);
            this.comboBox_sheetnames.Name = "comboBox_sheetnames";
            this.comboBox_sheetnames.Size = new System.Drawing.Size(329, 21);
            this.comboBox_sheetnames.TabIndex = 15;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(27, 38);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(191, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "选择当前需要处理的Sheet页名称：";
            // 
            // text_startline
            // 
            this.text_startline.Location = new System.Drawing.Point(29, 143);
            this.text_startline.Name = "text_startline";
            this.text_startline.Size = new System.Drawing.Size(329, 20);
            this.text_startline.TabIndex = 17;
            this.text_startline.Text = "1";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(27, 114);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(165, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "请输入列所在的行数（如“1”）";
            // 
            // Formfilldata
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(395, 420);
            this.Controls.Add(this.text_startline);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.comboBox_sheetnames);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxreadcolum);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cloumIDtextsplit);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "Formfilldata";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "填充数据";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox cloumIDtextsplit;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBoxreadcolum;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBox_sheetnames;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox text_startline;
        private System.Windows.Forms.Label label4;
    }
}