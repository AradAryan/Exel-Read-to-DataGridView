namespace Tamrin8
{
    partial class MainForm
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
            this.mainDataGird = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.mainDataGird)).BeginInit();
            this.SuspendLayout();
            // 
            // mainDataGird
            // 
            this.mainDataGird.AllowUserToAddRows = false;
            this.mainDataGird.AllowUserToDeleteRows = false;
            this.mainDataGird.AllowUserToResizeColumns = false;
            this.mainDataGird.AllowUserToResizeRows = false;
            this.mainDataGird.BackgroundColor = System.Drawing.Color.Silver;
            this.mainDataGird.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.mainDataGird.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainDataGird.Location = new System.Drawing.Point(0, 0);
            this.mainDataGird.Name = "mainDataGird";
            this.mainDataGird.Size = new System.Drawing.Size(330, 222);
            this.mainDataGird.TabIndex = 2;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(330, 222);
            this.Controls.Add(this.mainDataGird);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Main";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.mainDataGird)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.DataGridView mainDataGird;
    }
}

