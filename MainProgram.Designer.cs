namespace 傑偲工程管理系統_途程簡易新增程式
{
    partial class MainProgram
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainProgram));
            this.dgvMasShow = new System.Windows.Forms.DataGridView();
            this.lblSelectFile = new System.Windows.Forms.Label();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnDot = new System.Windows.Forms.Button();
            this.btnLoad = new System.Windows.Forms.Button();
            this.btnUpload = new System.Windows.Forms.Button();
            this.dgvDtlShow = new System.Windows.Forms.DataGridView();
            this.lblMas = new System.Windows.Forms.Label();
            this.lblDtl = new System.Windows.Forms.Label();
            this.lblExplain = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMasShow)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDtlShow)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvMasShow
            // 
            this.dgvMasShow.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvMasShow.Location = new System.Drawing.Point(12, 139);
            this.dgvMasShow.Name = "dgvMasShow";
            this.dgvMasShow.RowTemplate.Height = 24;
            this.dgvMasShow.Size = new System.Drawing.Size(760, 100);
            this.dgvMasShow.TabIndex = 0;
            // 
            // lblSelectFile
            // 
            this.lblSelectFile.AutoSize = true;
            this.lblSelectFile.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblSelectFile.Location = new System.Drawing.Point(8, 22);
            this.lblSelectFile.Name = "lblSelectFile";
            this.lblSelectFile.Size = new System.Drawing.Size(207, 20);
            this.lblSelectFile.TabIndex = 1;
            this.lblSelectFile.Text = "請選擇要上傳的Excel檔案：";
            // 
            // txtFilePath
            // 
            this.txtFilePath.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtFilePath.Location = new System.Drawing.Point(206, 19);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(400, 29);
            this.txtFilePath.TabIndex = 2;
            // 
            // btnDot
            // 
            this.btnDot.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnDot.Location = new System.Drawing.Point(612, 15);
            this.btnDot.Name = "btnDot";
            this.btnDot.Size = new System.Drawing.Size(40, 34);
            this.btnDot.TabIndex = 3;
            this.btnDot.Text = "...";
            this.btnDot.UseVisualStyleBackColor = true;
            this.btnDot.Click += new System.EventHandler(this.btnDot_Click);
            // 
            // btnLoad
            // 
            this.btnLoad.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnLoad.Location = new System.Drawing.Point(662, 14);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(110, 35);
            this.btnLoad.TabIndex = 4;
            this.btnLoad.Text = "載入資料";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // btnUpload
            // 
            this.btnUpload.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnUpload.Location = new System.Drawing.Point(662, 98);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(110, 35);
            this.btnUpload.TabIndex = 5;
            this.btnUpload.Text = "開始上傳";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // dgvDtlShow
            // 
            this.dgvDtlShow.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDtlShow.Location = new System.Drawing.Point(12, 289);
            this.dgvDtlShow.Name = "dgvDtlShow";
            this.dgvDtlShow.RowTemplate.Height = 24;
            this.dgvDtlShow.Size = new System.Drawing.Size(760, 300);
            this.dgvDtlShow.TabIndex = 6;
            // 
            // lblMas
            // 
            this.lblMas.AutoSize = true;
            this.lblMas.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblMas.Location = new System.Drawing.Point(12, 116);
            this.lblMas.Name = "lblMas";
            this.lblMas.Size = new System.Drawing.Size(166, 20);
            this.lblMas.TabIndex = 7;
            this.lblMas.Text = "EMOdTmpRouteMas";
            // 
            // lblDtl
            // 
            this.lblDtl.AutoSize = true;
            this.lblDtl.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblDtl.Location = new System.Drawing.Point(12, 266);
            this.lblDtl.Name = "lblDtl";
            this.lblDtl.Size = new System.Drawing.Size(157, 20);
            this.lblDtl.TabIndex = 8;
            this.lblDtl.Text = "EMOdTmpRouteDtl";
            // 
            // lblExplain
            // 
            this.lblExplain.AutoSize = true;
            this.lblExplain.Font = new System.Drawing.Font("微軟正黑體", 18F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblExplain.Location = new System.Drawing.Point(133, 67);
            this.lblExplain.Name = "lblExplain";
            this.lblExplain.Size = new System.Drawing.Size(518, 31);
            this.lblExplain.TabIndex = 9;
            this.lblExplain.Text = "＊請先將資料載入，確認無誤後，再進行上傳！";
            // 
            // MainProgram
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(784, 601);
            this.Controls.Add(this.lblExplain);
            this.Controls.Add(this.lblDtl);
            this.Controls.Add(this.lblMas);
            this.Controls.Add(this.dgvDtlShow);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.btnDot);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.lblSelectFile);
            this.Controls.Add(this.dgvMasShow);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainProgram";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "工程途程簡易上傳程式 v1.0";
            ((System.ComponentModel.ISupportInitialize)(this.dgvMasShow)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDtlShow)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvMasShow;
        private System.Windows.Forms.Label lblSelectFile;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnDot;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.DataGridView dgvDtlShow;
        private System.Windows.Forms.Label lblMas;
        private System.Windows.Forms.Label lblDtl;
        private System.Windows.Forms.Label lblExplain;
    }
}

