
namespace ScheduleFGR
{
    partial class GreenWichSchedule
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialogData = new System.Windows.Forms.OpenFileDialog();
            this.btnOpenData = new System.Windows.Forms.Button();
            this.tbxOpenFile = new System.Windows.Forms.TextBox();
            this.dgvListData = new System.Windows.Forms.DataGridView();
            this.btnExportSchedule = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvListData)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialogData
            // 
            this.openFileDialogData.FileName = "openFileDialogData";
            // 
            // btnOpenData
            // 
            this.btnOpenData.Location = new System.Drawing.Point(25, 35);
            this.btnOpenData.Name = "btnOpenData";
            this.btnOpenData.Size = new System.Drawing.Size(94, 29);
            this.btnOpenData.TabIndex = 0;
            this.btnOpenData.Text = "Open File:";
            this.btnOpenData.UseVisualStyleBackColor = true;
            this.btnOpenData.Click += new System.EventHandler(this.btnOpenData_Click);
            // 
            // tbxOpenFile
            // 
            this.tbxOpenFile.Location = new System.Drawing.Point(157, 37);
            this.tbxOpenFile.Name = "tbxOpenFile";
            this.tbxOpenFile.ReadOnly = true;
            this.tbxOpenFile.Size = new System.Drawing.Size(617, 27);
            this.tbxOpenFile.TabIndex = 1;
            // 
            // dgvListData
            // 
            this.dgvListData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvListData.Location = new System.Drawing.Point(25, 94);
            this.dgvListData.Name = "dgvListData";
            this.dgvListData.RowHeadersWidth = 51;
            this.dgvListData.RowTemplate.Height = 29;
            this.dgvListData.Size = new System.Drawing.Size(749, 342);
            this.dgvListData.TabIndex = 2;
            // 
            // btnExportSchedule
            // 
            this.btnExportSchedule.Enabled = false;
            this.btnExportSchedule.Location = new System.Drawing.Point(613, 460);
            this.btnExportSchedule.Name = "btnExportSchedule";
            this.btnExportSchedule.Size = new System.Drawing.Size(125, 29);
            this.btnExportSchedule.TabIndex = 3;
            this.btnExportSchedule.Text = "Export Schedule";
            this.btnExportSchedule.UseVisualStyleBackColor = true;
            this.btnExportSchedule.Click += new System.EventHandler(this.btnExportSchedule_Click);
            // 
            // GreenWichSchedule
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 512);
            this.Controls.Add(this.btnExportSchedule);
            this.Controls.Add(this.dgvListData);
            this.Controls.Add(this.tbxOpenFile);
            this.Controls.Add(this.btnOpenData);
            this.Name = "GreenWichSchedule";
            this.Text = "GreenWichSchedule";
            ((System.ComponentModel.ISupportInitialize)(this.dgvListData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialogData;
        private System.Windows.Forms.Button btnOpenData;
        private System.Windows.Forms.TextBox tbxOpenFile;
        private System.Windows.Forms.DataGridView dgvListData;
        private System.Windows.Forms.Button btnExportSchedule;
    }
}

