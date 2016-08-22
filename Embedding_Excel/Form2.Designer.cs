namespace EmbeddedExcel
{
    partial class Form2
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
            this.excelWrapper = new EmbeddedExcel.ExcelWrapper();
            this.SuspendLayout();
            // 
            // excelWrapper
            // 
            this.excelWrapper.Dock = System.Windows.Forms.DockStyle.Fill;
            this.excelWrapper.Location = new System.Drawing.Point(0, 0);
            this.excelWrapper.Margin = new System.Windows.Forms.Padding(5);
            this.excelWrapper.Name = "excelWrapper";
            this.excelWrapper.Size = new System.Drawing.Size(990, 553);
            this.excelWrapper.TabIndex = 1;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(990, 553);
            this.Controls.Add(this.excelWrapper);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form2";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form2_FormClosing);
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private ExcelWrapper excelWrapper;
    }
}