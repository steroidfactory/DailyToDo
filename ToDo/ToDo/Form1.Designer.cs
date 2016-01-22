namespace ToDo
{
    partial class Form1
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.reportReadBox = new System.Windows.Forms.DataGridView();
            this.btnImport = new System.Windows.Forms.Button();
            this.exportFile = new System.Windows.Forms.SaveFileDialog();
            this.barProgress = new System.Windows.Forms.ProgressBar();
            this.lblBarProgress = new System.Windows.Forms.Label();
            this.importFile = new System.Windows.Forms.OpenFileDialog();
            this.btnCancel = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.Barcode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Status = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Family = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Lot = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Make = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Model = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PartNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.reportReadBox)).BeginInit();
            this.SuspendLayout();
            // 
            // btnRefresh
            // 
            this.btnRefresh.Enabled = false;
            this.btnRefresh.Location = new System.Drawing.Point(533, 463);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(75, 23);
            this.btnRefresh.TabIndex = 0;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnExport
            // 
            this.btnExport.Enabled = false;
            this.btnExport.Location = new System.Drawing.Point(380, 462);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(103, 24);
            this.btnExport.TabIndex = 1;
            this.btnExport.Text = "Export to Excel";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Enabled = false;
            this.button3.Location = new System.Drawing.Point(765, 463);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 2;
            this.button3.Text = "Print";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Visible = false;
            // 
            // reportReadBox
            // 
            this.reportReadBox.AllowUserToAddRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.reportReadBox.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.reportReadBox.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.reportReadBox.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Barcode,
            this.Status,
            this.Family,
            this.Lot,
            this.Make,
            this.Model,
            this.PartNumber});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.reportReadBox.DefaultCellStyle = dataGridViewCellStyle2;
            this.reportReadBox.Location = new System.Drawing.Point(184, 108);
            this.reportReadBox.Name = "reportReadBox";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.reportReadBox.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.reportReadBox.Size = new System.Drawing.Size(801, 325);
            this.reportReadBox.TabIndex = 3;
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(184, 463);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 4;
            this.btnImport.Text = "Import File";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.button4_Click);
            // 
            // exportFile
            // 
            this.exportFile.FileOk += new System.ComponentModel.CancelEventHandler(this.exportFile_FileOk);
            // 
            // barProgress
            // 
            this.barProgress.Location = new System.Drawing.Point(337, 24);
            this.barProgress.Name = "barProgress";
            this.barProgress.Size = new System.Drawing.Size(442, 23);
            this.barProgress.TabIndex = 5;
            // 
            // lblBarProgress
            // 
            this.lblBarProgress.AutoSize = true;
            this.lblBarProgress.Location = new System.Drawing.Point(539, 63);
            this.lblBarProgress.Name = "lblBarProgress";
            this.lblBarProgress.Size = new System.Drawing.Size(0, 13);
            this.lblBarProgress.TabIndex = 6;
            // 
            // btnCancel
            // 
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(275, 462);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(52, 333);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 8;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // Barcode
            // 
            this.Barcode.HeaderText = "Serial No.";
            this.Barcode.Name = "Barcode";
            // 
            // Status
            // 
            this.Status.HeaderText = "Status";
            this.Status.Name = "Status";
            // 
            // Family
            // 
            this.Family.HeaderText = "Product Family";
            this.Family.Name = "Family";
            // 
            // Lot
            // 
            this.Lot.HeaderText = "Lot No.";
            this.Lot.Name = "Lot";
            // 
            // Make
            // 
            this.Make.HeaderText = "Manufacturer";
            this.Make.Name = "Make";
            // 
            // Model
            // 
            this.Model.HeaderText = "Mfg Model No.";
            this.Model.Name = "Model";
            // 
            // PartNumber
            // 
            this.PartNumber.HeaderText = "Mfg Part No.";
            this.PartNumber.Name = "PartNumber";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1130, 652);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblBarProgress);
            this.Controls.Add(this.barProgress);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.reportReadBox);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnRefresh);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.reportReadBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.DataGridView reportReadBox;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.SaveFileDialog exportFile;
        private System.Windows.Forms.ProgressBar barProgress;
        private System.Windows.Forms.Label lblBarProgress;
        private System.Windows.Forms.OpenFileDialog importFile;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Barcode;
        private System.Windows.Forms.DataGridViewTextBoxColumn Status;
        private System.Windows.Forms.DataGridViewTextBoxColumn Family;
        private System.Windows.Forms.DataGridViewTextBoxColumn Lot;
        private System.Windows.Forms.DataGridViewTextBoxColumn Make;
        private System.Windows.Forms.DataGridViewTextBoxColumn Model;
        private System.Windows.Forms.DataGridViewTextBoxColumn PartNumber;
    }
}

