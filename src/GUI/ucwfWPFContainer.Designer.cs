namespace ExcelDna_MVVM.GUI
{
    partial class ucwfWPFContainer
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.WPFContainer = new System.Windows.Forms.Integration.ElementHost();
            this.SuspendLayout();
            // 
            // WPFContainer
            // 
            this.WPFContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.WPFContainer.Location = new System.Drawing.Point(0, 0);
            this.WPFContainer.Name = "WPFContainer";
            this.WPFContainer.Size = new System.Drawing.Size(670, 103);
            this.WPFContainer.TabIndex = 0;
            this.WPFContainer.Child = null;
            // 
            // ucwfWPFContainer
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Controls.Add(this.WPFContainer);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "ucwfWPFContainer";
            this.Size = new System.Drawing.Size(670, 103);
            this.ResumeLayout(false);

        }
        #endregion

        private System.Windows.Forms.Integration.ElementHost WPFContainer;
    }
}
