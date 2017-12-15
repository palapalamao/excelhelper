namespace OverTimeStatistics
{
    partial class SetStartLineForm
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
            this.text_startline = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_sheetname = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // text_startline
            // 
            this.text_startline.Location = new System.Drawing.Point(59, 138);
            this.text_startline.Name = "text_startline";
            this.text_startline.Size = new System.Drawing.Size(139, 20);
            this.text_startline.TabIndex = 5;
            this.text_startline.Text = "1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(57, 109);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(165, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "请输入列所在的行数（如“1”）";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(59, 178);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 25);
            this.button1.TabIndex = 3;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(56, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(107, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "当前Sheet页名称：";
            // 
            // textBox_sheetname
            // 
            this.textBox_sheetname.Enabled = false;
            this.textBox_sheetname.Location = new System.Drawing.Point(60, 69);
            this.textBox_sheetname.Name = "textBox_sheetname";
            this.textBox_sheetname.Size = new System.Drawing.Size(138, 20);
            this.textBox_sheetname.TabIndex = 7;
            // 
            // SetStartLineForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.textBox_sheetname);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.text_startline);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "SetStartLineForm";
            this.Text = "SetStartLineForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox text_startline;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_sheetname;
    }
}