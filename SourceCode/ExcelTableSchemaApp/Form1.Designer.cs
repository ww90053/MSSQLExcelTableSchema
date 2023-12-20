namespace ExcelTableSchemaApp
{
    partial class Form1
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
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.btnexport = new System.Windows.Forms.Button();
            this.linkchkall = new System.Windows.Forms.LinkLabel();
            this.linkunchkall = new System.Windows.Forms.LinkLabel();
            this.chklistTable = new System.Windows.Forms.CheckedListBox();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtexpfilepath = new System.Windows.Forms.TextBox();
            this.btnBrowser = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtcreuser = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnLoadDB = new System.Windows.Forms.Button();
            this.txtConnString = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProgress = new System.Windows.Forms.Label();
            this.txtOutputFile = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnexport
            // 
            this.btnexport.Location = new System.Drawing.Point(12, 496);
            this.btnexport.Name = "btnexport";
            this.btnexport.Size = new System.Drawing.Size(75, 23);
            this.btnexport.TabIndex = 0;
            this.btnexport.Text = "匯出EXCEL";
            this.btnexport.UseVisualStyleBackColor = true;
            this.btnexport.Click += new System.EventHandler(this.btnexport_Click);
            // 
            // linkchkall
            // 
            this.linkchkall.AutoSize = true;
            this.linkchkall.Location = new System.Drawing.Point(6, 18);
            this.linkchkall.Name = "linkchkall";
            this.linkchkall.Size = new System.Drawing.Size(29, 12);
            this.linkchkall.TabIndex = 2;
            this.linkchkall.TabStop = true;
            this.linkchkall.Text = "全選";
            this.linkchkall.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkchkall_LinkClicked);
            // 
            // linkunchkall
            // 
            this.linkunchkall.AutoSize = true;
            this.linkunchkall.Location = new System.Drawing.Point(41, 18);
            this.linkunchkall.Name = "linkunchkall";
            this.linkunchkall.Size = new System.Drawing.Size(41, 12);
            this.linkunchkall.TabIndex = 3;
            this.linkunchkall.TabStop = true;
            this.linkunchkall.Text = "全不選";
            this.linkunchkall.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkunchkall_LinkClicked);
            // 
            // chklistTable
            // 
            this.chklistTable.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.chklistTable.CheckOnClick = true;
            this.chklistTable.FormattingEnabled = true;
            this.chklistTable.Location = new System.Drawing.Point(20, 36);
            this.chklistTable.Name = "chklistTable";
            this.chklistTable.Size = new System.Drawing.Size(454, 223);
            this.chklistTable.TabIndex = 1;
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(529, 343);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(8, 4);
            this.checkedListBox1.TabIndex = 4;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chklistTable);
            this.groupBox1.Controls.Add(this.linkunchkall);
            this.groupBox1.Controls.Add(this.linkchkall);
            this.groupBox1.Location = new System.Drawing.Point(12, 73);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(490, 282);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Table名稱";
            // 
            // txtexpfilepath
            // 
            this.txtexpfilepath.Location = new System.Drawing.Point(15, 417);
            this.txtexpfilepath.Name = "txtexpfilepath";
            this.txtexpfilepath.Size = new System.Drawing.Size(397, 22);
            this.txtexpfilepath.TabIndex = 6;
            // 
            // btnBrowser
            // 
            this.btnBrowser.Location = new System.Drawing.Point(427, 417);
            this.btnBrowser.Name = "btnBrowser";
            this.btnBrowser.Size = new System.Drawing.Size(75, 23);
            this.btnBrowser.TabIndex = 7;
            this.btnBrowser.Text = "瀏灠";
            this.btnBrowser.UseVisualStyleBackColor = true;
            this.btnBrowser.Click += new System.EventHandler(this.btnBrowser_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 402);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 8;
            this.label1.Text = "檔案匯出路徑";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 358);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 10;
            this.label2.Text = "填表人";
            // 
            // txtcreuser
            // 
            this.txtcreuser.Location = new System.Drawing.Point(15, 373);
            this.txtcreuser.Name = "txtcreuser";
            this.txtcreuser.Size = new System.Drawing.Size(360, 22);
            this.txtcreuser.TabIndex = 9;
            this.txtcreuser.Text = "請輸入填表人";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnLoadDB);
            this.groupBox2.Controls.Add(this.txtConnString);
            this.groupBox2.Location = new System.Drawing.Point(15, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(487, 55);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "連線字串";
            // 
            // btnLoadDB
            // 
            this.btnLoadDB.Location = new System.Drawing.Point(406, 19);
            this.btnLoadDB.Name = "btnLoadDB";
            this.btnLoadDB.Size = new System.Drawing.Size(75, 23);
            this.btnLoadDB.TabIndex = 8;
            this.btnLoadDB.Text = "載入";
            this.btnLoadDB.UseVisualStyleBackColor = true;
            this.btnLoadDB.Click += new System.EventHandler(this.btnLoadDB_Click);
            // 
            // txtConnString
            // 
            this.txtConnString.Location = new System.Drawing.Point(17, 21);
            this.txtConnString.Name = "txtConnString";
            this.txtConnString.Size = new System.Drawing.Size(380, 22);
            this.txtConnString.TabIndex = 0;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 538);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(509, 23);
            this.progressBar1.TabIndex = 12;
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(104, 501);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(20, 12);
            this.lblProgress.TabIndex = 13;
            this.lblProgress.Text = "0/0";
            // 
            // txtOutputFile
            // 
            this.txtOutputFile.Location = new System.Drawing.Point(15, 468);
            this.txtOutputFile.Name = "txtOutputFile";
            this.txtOutputFile.Size = new System.Drawing.Size(162, 22);
            this.txtOutputFile.TabIndex = 14;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 453);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 15;
            this.label3.Text = "匯出檔名";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(183, 478);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(21, 12);
            this.label4.TabIndex = 16;
            this.label4.Text = ".xls";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(533, 575);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtOutputFile);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtcreuser);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnBrowser);
            this.Controls.Add(this.txtexpfilepath);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.btnexport);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Form1";
            this.Text = "TableSchema匯出";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnexport;
        private System.Windows.Forms.LinkLabel linkchkall;
        private System.Windows.Forms.LinkLabel linkunchkall;
        private System.Windows.Forms.CheckedListBox chklistTable;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtexpfilepath;
        private System.Windows.Forms.Button btnBrowser;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtcreuser;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox txtConnString;
        private System.Windows.Forms.Button btnLoadDB;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.TextBox txtOutputFile;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}

