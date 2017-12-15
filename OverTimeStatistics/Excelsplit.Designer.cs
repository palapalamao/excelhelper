namespace OverTimeStatistics
{
    partial class Excelsplit
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
            this.formOverTimeList1 = new OverTimeStatistics.FormOverTimeList();
            this.SuspendLayout();
            // 
            // formOverTimeList1
            // 
            this.formOverTimeList1.Location = new System.Drawing.Point(24, 21);
            this.formOverTimeList1.Margin = new System.Windows.Forms.Padding(2);
            this.formOverTimeList1.Name = "formOverTimeList1";
            this.formOverTimeList1.Size = new System.Drawing.Size(854, 586);
            this.formOverTimeList1.TabIndex = 0;
            this.formOverTimeList1.Load += new System.EventHandler(this.formOverTimeList1_Load);
            // 
            // Excelsplit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(879, 604);
            this.Controls.Add(this.formOverTimeList1);
            this.Name = "Excelsplit";
            this.Text = "Excel拆分合并助手";
            this.ResumeLayout(false);

        }

        #endregion

        private FormOverTimeList formOverTimeList1;
    }
}