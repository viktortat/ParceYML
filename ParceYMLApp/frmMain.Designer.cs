namespace ParceYmlApp
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
            this.label1 = new System.Windows.Forms.Label();
            this.btnSelFile = new System.Windows.Forms.Button();
            this.txbPathSelector = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnCreatePricat = new System.Windows.Forms.Button();
            this.btnParseInExcel = new System.Windows.Forms.Button();
            this.chbCopyToDB = new System.Windows.Forms.CheckBox();
            this.button3 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.lblInfo = new System.Windows.Forms.Label();
            this.btnParseFromExcel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.BackColor = System.Drawing.Color.White;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(-3, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(417, 24);
            this.label1.TabIndex = 1;
            this.label1.Text = "Разбор YML для каталога";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnSelFile
            // 
            this.btnSelFile.Location = new System.Drawing.Point(375, 42);
            this.btnSelFile.Name = "btnSelFile";
            this.btnSelFile.Size = new System.Drawing.Size(33, 22);
            this.btnSelFile.TabIndex = 17;
            this.btnSelFile.Text = "...";
            this.btnSelFile.UseVisualStyleBackColor = true;
            this.btnSelFile.Click += new System.EventHandler(this.btnSelFile_Click);
            // 
            // txbPathSelector
            // 
            this.txbPathSelector.Location = new System.Drawing.Point(89, 43);
            this.txbPathSelector.Name = "txbPathSelector";
            this.txbPathSelector.Size = new System.Drawing.Size(286, 20);
            this.txbPathSelector.TabIndex = 16;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "Выбор файла:";
            // 
            // btnCreatePricat
            // 
            this.btnCreatePricat.Location = new System.Drawing.Point(282, 69);
            this.btnCreatePricat.Name = "btnCreatePricat";
            this.btnCreatePricat.Size = new System.Drawing.Size(126, 23);
            this.btnCreatePricat.TabIndex = 18;
            this.btnCreatePricat.Text = "Формирование Pricat";
            this.btnCreatePricat.UseVisualStyleBackColor = true;
            this.btnCreatePricat.Click += new System.EventHandler(this.btnCreatePricat_Click);
            // 
            // btnParseInExcel
            // 
            this.btnParseInExcel.Location = new System.Drawing.Point(7, 95);
            this.btnParseInExcel.Name = "btnParseInExcel";
            this.btnParseInExcel.Size = new System.Drawing.Size(203, 28);
            this.btnParseInExcel.TabIndex = 19;
            this.btnParseInExcel.Text = "Разбор в Excel";
            this.btnParseInExcel.UseVisualStyleBackColor = true;
            this.btnParseInExcel.Click += new System.EventHandler(this.btnParseInExcel_Click);
            // 
            // chbCopyToDB
            // 
            this.chbCopyToDB.AutoSize = true;
            this.chbCopyToDB.Checked = true;
            this.chbCopyToDB.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chbCopyToDB.Location = new System.Drawing.Point(7, 72);
            this.chbCopyToDB.Name = "chbCopyToDB";
            this.chbCopyToDB.Size = new System.Drawing.Size(102, 17);
            this.chbCopyToDB.TabIndex = 20;
            this.chbCopyToDB.Text = "Записать в БД";
            this.chbCopyToDB.UseVisualStyleBackColor = true;
            this.chbCopyToDB.CheckedChanged += new System.EventHandler(this.chbCopyToDB_CheckedChanged);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(216, 69);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(68, 23);
            this.button3.TabIndex = 21;
            this.button3.Text = "Test";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(7, 152);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(401, 298);
            this.dataGridView1.TabIndex = 22;
            // 
            // lblInfo
            // 
            this.lblInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lblInfo.Location = new System.Drawing.Point(7, 126);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(401, 23);
            this.lblInfo.TabIndex = 23;
            this.lblInfo.Text = "...";
            this.lblInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnParseFromExcel
            // 
            this.btnParseFromExcel.Location = new System.Drawing.Point(216, 95);
            this.btnParseFromExcel.Name = "btnParseFromExcel";
            this.btnParseFromExcel.Size = new System.Drawing.Size(192, 28);
            this.btnParseFromExcel.TabIndex = 24;
            this.btnParseFromExcel.Text = "Разбор из Excel";
            this.btnParseFromExcel.UseVisualStyleBackColor = true;
            this.btnParseFromExcel.Click += new System.EventHandler(this.btnParseFromExcel_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(411, 462);
            this.Controls.Add(this.btnParseFromExcel);
            this.Controls.Add(this.lblInfo);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.chbCopyToDB);
            this.Controls.Add(this.btnParseInExcel);
            this.Controls.Add(this.btnCreatePricat);
            this.Controls.Add(this.btnSelFile);
            this.Controls.Add(this.txbPathSelector);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Варианты разбора YML";
            this.Load += new System.EventHandler(this.frmMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSelFile;
        private System.Windows.Forms.TextBox txbPathSelector;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnCreatePricat;
        private System.Windows.Forms.Button btnParseInExcel;
        private System.Windows.Forms.CheckBox chbCopyToDB;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.Button btnParseFromExcel;
    }
}

