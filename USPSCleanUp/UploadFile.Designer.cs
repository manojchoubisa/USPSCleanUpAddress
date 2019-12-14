namespace USPSCleanUp
{
    partial class UploadFile
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnUpload = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.lblmessage = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnExportNewExcel = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnErrorLog = new System.Windows.Forms.Button();
            this.lblResult = new System.Windows.Forms.Label();
            this.btnCleasing = new System.Windows.Forms.Button();
            this.btnDublicate = new System.Windows.Forms.Button();
            this.btnClearDublicate = new System.Windows.Forms.Button();
            this.lblfilepath = new System.Windows.Forms.Label();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.pnlHide = new System.Windows.Forms.Panel();
            this.lblSheetName = new System.Windows.Forms.Label();
            this.cbSheetList = new System.Windows.Forms.ComboBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.tabPageProcessFiles = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.pnlHide.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(27, 13);
            this.textBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(317, 22);
            this.textBox1.TabIndex = 0;
            // 
            // btnUpload
            // 
            this.btnUpload.Location = new System.Drawing.Point(382, 11);
            this.btnUpload.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(115, 27);
            this.btnUpload.TabIndex = 1;
            this.btnUpload.Text = "Upload";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // lblmessage
            // 
            this.lblmessage.AutoSize = true;
            this.lblmessage.Location = new System.Drawing.Point(11, 54);
            this.lblmessage.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblmessage.Name = "lblmessage";
            this.lblmessage.Size = new System.Drawing.Size(79, 17);
            this.lblmessage.TabIndex = 4;
            this.lblmessage.Text = "lblmessage";
            this.lblmessage.Visible = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 217);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.Size = new System.Drawing.Size(1417, 409);
            this.dataGridView1.TabIndex = 5;
            this.dataGridView1.Visible = false;
            // 
            // btnExportNewExcel
            // 
            this.btnExportNewExcel.Location = new System.Drawing.Point(1086, 14);
            this.btnExportNewExcel.Margin = new System.Windows.Forms.Padding(4);
            this.btnExportNewExcel.Name = "btnExportNewExcel";
            this.btnExportNewExcel.Size = new System.Drawing.Size(109, 28);
            this.btnExportNewExcel.TabIndex = 6;
            this.btnExportNewExcel.Text = "Save Excel";
            this.btnExportNewExcel.UseVisualStyleBackColor = true;
            this.btnExportNewExcel.Visible = false;
            this.btnExportNewExcel.Click += new System.EventHandler(this.btnExportNewExcel_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(628, 85);
            this.lblStatus.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(62, 17);
            this.lblStatus.TabIndex = 8;
            this.lblStatus.Text = "lblStatus";
            this.lblStatus.Visible = false;
            // 
            // btnErrorLog
            // 
            this.btnErrorLog.Location = new System.Drawing.Point(1202, 15);
            this.btnErrorLog.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnErrorLog.Name = "btnErrorLog";
            this.btnErrorLog.Size = new System.Drawing.Size(115, 27);
            this.btnErrorLog.TabIndex = 9;
            this.btnErrorLog.Text = "Error Log";
            this.btnErrorLog.UseVisualStyleBackColor = true;
            this.btnErrorLog.Visible = false;
            this.btnErrorLog.Click += new System.EventHandler(this.btnErrorLog_Click);
            // 
            // lblResult
            // 
            this.lblResult.AutoSize = true;
            this.lblResult.Location = new System.Drawing.Point(9, 176);
            this.lblResult.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(62, 17);
            this.lblResult.TabIndex = 10;
            this.lblResult.Text = "lblResult";
            this.lblResult.Visible = false;
            // 
            // btnCleasing
            // 
            this.btnCleasing.Location = new System.Drawing.Point(689, 14);
            this.btnCleasing.Margin = new System.Windows.Forms.Padding(4);
            this.btnCleasing.Name = "btnCleasing";
            this.btnCleasing.Size = new System.Drawing.Size(143, 28);
            this.btnCleasing.TabIndex = 11;
            this.btnCleasing.Text = "Cleaning Address";
            this.btnCleasing.UseVisualStyleBackColor = true;
            this.btnCleasing.Visible = false;
            this.btnCleasing.Click += new System.EventHandler(this.btnCleasing_Click);
            // 
            // btnDublicate
            // 
            this.btnDublicate.Location = new System.Drawing.Point(839, 14);
            this.btnDublicate.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnDublicate.Name = "btnDublicate";
            this.btnDublicate.Size = new System.Drawing.Size(115, 28);
            this.btnDublicate.TabIndex = 12;
            this.btnDublicate.Text = "View Dublicate";
            this.btnDublicate.UseVisualStyleBackColor = true;
            this.btnDublicate.Visible = false;
            this.btnDublicate.Click += new System.EventHandler(this.btnDublicate_Click);
            // 
            // btnClearDublicate
            // 
            this.btnClearDublicate.Location = new System.Drawing.Point(964, 14);
            this.btnClearDublicate.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnClearDublicate.Name = "btnClearDublicate";
            this.btnClearDublicate.Size = new System.Drawing.Size(115, 28);
            this.btnClearDublicate.TabIndex = 13;
            this.btnClearDublicate.Text = "ClearDublicate";
            this.btnClearDublicate.UseVisualStyleBackColor = true;
            this.btnClearDublicate.Visible = false;
            this.btnClearDublicate.Click += new System.EventHandler(this.btnClearDublicate_Click);
            // 
            // lblfilepath
            // 
            this.lblfilepath.AutoSize = true;
            this.lblfilepath.Location = new System.Drawing.Point(628, 113);
            this.lblfilepath.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblfilepath.Name = "lblfilepath";
            this.lblfilepath.Size = new System.Drawing.Size(73, 17);
            this.lblfilepath.TabIndex = 14;
            this.lblfilepath.Text = "lblFilePath";
            this.lblfilepath.Visible = false;
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(631, 134);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(4);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(798, 22);
            this.txtFilePath.TabIndex = 15;
            this.txtFilePath.TabStop = false;
            this.txtFilePath.Visible = false;
            // 
            // pnlHide
            // 
            this.pnlHide.Controls.Add(this.textBox1);
            this.pnlHide.Controls.Add(this.btnUpload);
            this.pnlHide.Location = new System.Drawing.Point(12, 1);
            this.pnlHide.Name = "pnlHide";
            this.pnlHide.Size = new System.Drawing.Size(518, 50);
            this.pnlHide.TabIndex = 16;
            // 
            // lblSheetName
            // 
            this.lblSheetName.AutoSize = true;
            this.lblSheetName.Location = new System.Drawing.Point(14, 130);
            this.lblSheetName.Name = "lblSheetName";
            this.lblSheetName.Size = new System.Drawing.Size(82, 17);
            this.lblSheetName.TabIndex = 17;
            this.lblSheetName.Text = "SheetName";
            this.lblSheetName.Visible = false;
            // 
            // cbSheetList
            // 
            this.cbSheetList.FormattingEnabled = true;
            this.cbSheetList.Location = new System.Drawing.Point(118, 130);
            this.cbSheetList.Name = "cbSheetList";
            this.cbSheetList.Size = new System.Drawing.Size(462, 24);
            this.cbSheetList.TabIndex = 18;
            this.cbSheetList.Visible = false;
            this.cbSheetList.SelectedIndexChanged += new System.EventHandler(this.cbSheetList_SelectedIndexChanged);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(1323, 15);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(96, 28);
            this.btnCancel.TabIndex = 19;
            this.btnCancel.Text = "Start Over";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // tabPageProcessFiles
            // 
            this.tabPageProcessFiles.Location = new System.Drawing.Point(12, 650);
            this.tabPageProcessFiles.Name = "tabPageProcessFiles";
            this.tabPageProcessFiles.Size = new System.Drawing.Size(1417, 18);
            this.tabPageProcessFiles.TabIndex = 20;
            this.tabPageProcessFiles.Visible = false;
            // 
            // UploadFile
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.ClientSize = new System.Drawing.Size(1440, 677);
            this.Controls.Add(this.tabPageProcessFiles);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.cbSheetList);
            this.Controls.Add(this.lblSheetName);
            this.Controls.Add(this.pnlHide);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.lblfilepath);
            this.Controls.Add(this.btnClearDublicate);
            this.Controls.Add(this.btnDublicate);
            this.Controls.Add(this.btnCleasing);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.btnErrorLog);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnExportNewExcel);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.lblmessage);
            this.ForeColor = System.Drawing.Color.SteelBlue;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "UploadFile";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "File Manage";
            this.Load += new System.EventHandler(this.UploadFile_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.pnlHide.ResumeLayout(false);
            this.pnlHide.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label lblmessage;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnExportNewExcel;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnErrorLog;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Button btnCleasing;
        private System.Windows.Forms.Button btnDublicate;
        private System.Windows.Forms.Button btnClearDublicate;
        private System.Windows.Forms.Label lblfilepath;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Panel pnlHide;
        private System.Windows.Forms.Label lblSheetName;
        private System.Windows.Forms.ComboBox cbSheetList;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ProgressBar tabPageProcessFiles;
    }
}

