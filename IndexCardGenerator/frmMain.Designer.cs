namespace IndexCardGenerator
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
            this.btnExecute = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtWordListFile = new System.Windows.Forms.TextBox();
            this.btnBrowseWordListFile = new System.Windows.Forms.Button();
            this.btnBrowseTemplate = new System.Windows.Forms.Button();
            this.txtTemplateFile = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cboEncoding = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btnExecute
            // 
            this.btnExecute.Location = new System.Drawing.Point(347, 92);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(75, 23);
            this.btnExecute.TabIndex = 0;
            this.btnExecute.Text = "&Create";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 47);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Word List:";
            // 
            // txtWordListFile
            // 
            this.txtWordListFile.Location = new System.Drawing.Point(73, 44);
            this.txtWordListFile.Name = "txtWordListFile";
            this.txtWordListFile.Size = new System.Drawing.Size(268, 20);
            this.txtWordListFile.TabIndex = 2;
            // 
            // btnBrowseWordListFile
            // 
            this.btnBrowseWordListFile.Location = new System.Drawing.Point(347, 41);
            this.btnBrowseWordListFile.Name = "btnBrowseWordListFile";
            this.btnBrowseWordListFile.Size = new System.Drawing.Size(75, 23);
            this.btnBrowseWordListFile.TabIndex = 3;
            this.btnBrowseWordListFile.Text = "&Browse";
            this.btnBrowseWordListFile.UseVisualStyleBackColor = true;
            this.btnBrowseWordListFile.Click += new System.EventHandler(this.btnBrowseWordListFile_Click);
            // 
            // btnBrowseTemplate
            // 
            this.btnBrowseTemplate.Location = new System.Drawing.Point(347, 12);
            this.btnBrowseTemplate.Name = "btnBrowseTemplate";
            this.btnBrowseTemplate.Size = new System.Drawing.Size(75, 23);
            this.btnBrowseTemplate.TabIndex = 6;
            this.btnBrowseTemplate.Text = "&Browse";
            this.btnBrowseTemplate.UseVisualStyleBackColor = true;
            this.btnBrowseTemplate.Click += new System.EventHandler(this.btnBrowseTemplate_Click);
            // 
            // txtTemplateFile
            // 
            this.txtTemplateFile.Location = new System.Drawing.Point(73, 15);
            this.txtTemplateFile.Name = "txtTemplateFile";
            this.txtTemplateFile.Size = new System.Drawing.Size(268, 20);
            this.txtTemplateFile.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Template:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 97);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Encoding: ";
            // 
            // cboEncoding
            // 
            this.cboEncoding.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboEncoding.FormattingEnabled = true;
            this.cboEncoding.Items.AddRange(new object[] {
            "ANSI",
            "UTF-8"});
            this.cboEncoding.Location = new System.Drawing.Point(76, 94);
            this.cboEncoding.Name = "cboEncoding";
            this.cboEncoding.Size = new System.Drawing.Size(121, 21);
            this.cboEncoding.TabIndex = 8;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(434, 127);
            this.Controls.Add(this.cboEncoding);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnBrowseTemplate);
            this.Controls.Add(this.txtTemplateFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnBrowseWordListFile);
            this.Controls.Add(this.txtWordListFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExecute);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Index Card Generator v1.0";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExecute;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtWordListFile;
        private System.Windows.Forms.Button btnBrowseWordListFile;
        private System.Windows.Forms.Button btnBrowseTemplate;
        private System.Windows.Forms.TextBox txtTemplateFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboEncoding;
    }
}

