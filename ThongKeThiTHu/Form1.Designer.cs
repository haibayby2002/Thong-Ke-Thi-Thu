namespace ThongKeThiTHu
{
    partial class frmMain
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
            this.btnReadExcel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnOpen = new System.Windows.Forms.Button();
            this.txtFile = new System.Windows.Forms.TextBox();
            this.grbLop = new System.Windows.Forms.GroupBox();
            this.lblLop = new System.Windows.Forms.Label();
            this.cmbLop = new System.Windows.Forms.ComboBox();
            this.btnThongKeTheoLop = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnThongKeTheoMon = new System.Windows.Forms.Button();
            this.cmbMon = new System.Windows.Forms.ComboBox();
            this.lblMon = new System.Windows.Forms.Label();
            this.grbLop.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnReadExcel
            // 
            this.btnReadExcel.Font = new System.Drawing.Font("Times New Roman", 12.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnReadExcel.Location = new System.Drawing.Point(12, 17);
            this.btnReadExcel.Name = "btnReadExcel";
            this.btnReadExcel.Size = new System.Drawing.Size(79, 29);
            this.btnReadExcel.TabIndex = 0;
            this.btnReadExcel.Text = "Đọc file Excel";
            this.btnReadExcel.UseVisualStyleBackColor = true;
            this.btnReadExcel.Click += new System.EventHandler(this.btnReadExcel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.Font = new System.Drawing.Font("Times New Roman", 12.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.Location = new System.Drawing.Point(317, 17);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(102, 28);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Lưu Backup";
            this.btnSave.UseVisualStyleBackColor = true;
            // 
            // btnOpen
            // 
            this.btnOpen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpen.Font = new System.Drawing.Font("Times New Roman", 12.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpen.Location = new System.Drawing.Point(425, 17);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(113, 28);
            this.btnOpen.TabIndex = 2;
            this.btnOpen.Text = "Mở backup";
            this.btnOpen.UseVisualStyleBackColor = true;
            // 
            // txtFile
            // 
            this.txtFile.Font = new System.Drawing.Font("Times New Roman", 12.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFile.Location = new System.Drawing.Point(85, 18);
            this.txtFile.Name = "txtFile";
            this.txtFile.ReadOnly = true;
            this.txtFile.Size = new System.Drawing.Size(127, 27);
            this.txtFile.TabIndex = 3;
            // 
            // grbLop
            // 
            this.grbLop.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.grbLop.Controls.Add(this.btnThongKeTheoLop);
            this.grbLop.Controls.Add(this.cmbLop);
            this.grbLop.Controls.Add(this.lblLop);
            this.grbLop.Font = new System.Drawing.Font("Times New Roman", 12.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grbLop.Location = new System.Drawing.Point(12, 82);
            this.grbLop.Name = "grbLop";
            this.grbLop.Size = new System.Drawing.Size(200, 167);
            this.grbLop.TabIndex = 4;
            this.grbLop.TabStop = false;
            this.grbLop.Text = "Thống kê theo lớp";
            // 
            // lblLop
            // 
            this.lblLop.AutoSize = true;
            this.lblLop.Font = new System.Drawing.Font("Times New Roman", 12.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLop.Location = new System.Drawing.Point(6, 25);
            this.lblLop.Name = "lblLop";
            this.lblLop.Size = new System.Drawing.Size(41, 19);
            this.lblLop.TabIndex = 0;
            this.lblLop.Text = "Lớp:";
            // 
            // cmbLop
            // 
            this.cmbLop.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLop.FormattingEnabled = true;
            this.cmbLop.Location = new System.Drawing.Point(53, 22);
            this.cmbLop.Name = "cmbLop";
            this.cmbLop.Size = new System.Drawing.Size(141, 27);
            this.cmbLop.TabIndex = 1;
            // 
            // btnThongKeTheoLop
            // 
            this.btnThongKeTheoLop.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnThongKeTheoLop.Location = new System.Drawing.Point(6, 136);
            this.btnThongKeTheoLop.Name = "btnThongKeTheoLop";
            this.btnThongKeTheoLop.Size = new System.Drawing.Size(188, 25);
            this.btnThongKeTheoLop.TabIndex = 2;
            this.btnThongKeTheoLop.Text = "Xuất Excel";
            this.btnThongKeTheoLop.UseVisualStyleBackColor = true;
            this.btnThongKeTheoLop.Click += new System.EventHandler(this.btnThongKeTheoLop_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.btnThongKeTheoMon);
            this.groupBox1.Controls.Add(this.cmbMon);
            this.groupBox1.Controls.Add(this.lblMon);
            this.groupBox1.Font = new System.Drawing.Font("Times New Roman", 12.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(317, 82);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(221, 167);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Thống kê theo lớp";
            // 
            // btnThongKeTheoMon
            // 
            this.btnThongKeTheoMon.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnThongKeTheoMon.Location = new System.Drawing.Point(6, 136);
            this.btnThongKeTheoMon.Name = "btnThongKeTheoMon";
            this.btnThongKeTheoMon.Size = new System.Drawing.Size(209, 25);
            this.btnThongKeTheoMon.TabIndex = 2;
            this.btnThongKeTheoMon.Text = "Xuất Excel";
            this.btnThongKeTheoMon.UseVisualStyleBackColor = true;
            // 
            // cmbMon
            // 
            this.cmbMon.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbMon.FormattingEnabled = true;
            this.cmbMon.Items.AddRange(new object[] {
            "Toán",
            "Sử",
            "Anh",
            "Lí",
            "Địa",
            "Hóa",
            "Sinh",
            "Văn",
            "GDCD"});
            this.cmbMon.Location = new System.Drawing.Point(58, 22);
            this.cmbMon.Name = "cmbMon";
            this.cmbMon.Size = new System.Drawing.Size(157, 27);
            this.cmbMon.TabIndex = 1;
            // 
            // lblMon
            // 
            this.lblMon.AutoSize = true;
            this.lblMon.Font = new System.Drawing.Font("Times New Roman", 12.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMon.Location = new System.Drawing.Point(6, 25);
            this.lblMon.Name = "lblMon";
            this.lblMon.Size = new System.Drawing.Size(46, 19);
            this.lblMon.TabIndex = 0;
            this.lblMon.Text = "Môn:";
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(550, 261);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.grbLop);
            this.Controls.Add(this.txtFile);
            this.Controls.Add(this.btnOpen);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnReadExcel);
            this.Name = "frmMain";
            this.Text = "Thống kê thi thử đại học";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.grbLop.ResumeLayout(false);
            this.grbLop.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnReadExcel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.TextBox txtFile;
        private System.Windows.Forms.GroupBox grbLop;
        private System.Windows.Forms.Button btnThongKeTheoLop;
        private System.Windows.Forms.ComboBox cmbLop;
        private System.Windows.Forms.Label lblLop;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnThongKeTheoMon;
        private System.Windows.Forms.ComboBox cmbMon;
        private System.Windows.Forms.Label lblMon;
    }
}

