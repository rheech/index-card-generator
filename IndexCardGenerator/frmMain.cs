using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office;
using System.Diagnostics;

namespace IndexCardGenerator
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            cboEncoding.SelectedIndex = 0;
        }

        private System.Text.Encoding getEncoding()
        {
            System.Text.Encoding encoding;

            switch (cboEncoding.Text)
            {
                case "ANSI":
                    encoding = System.Text.Encoding.Default;
                    break;
                case "UTF-8":
                    encoding = System.Text.Encoding.UTF8;
                    break;
                default:
                    encoding = System.Text.Encoding.Default;
                    break;
            }

            return encoding;
        }
        
        private void frmMain_Load(object sender, EventArgs e)
        {
            txtTemplateFile.Text = String.Format("{0}{1}", System.Windows.Forms.Application.StartupPath, "\\Index Card Template.dotx");
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            string[] sList, sLine;
            byte[] sBytes;
            System.IO.FileInfo fileTemplate;
            System.IO.FileInfo fileWordList;
            WordApp app;

            fileTemplate = new System.IO.FileInfo(txtTemplateFile.Text);
            fileWordList = new System.IO.FileInfo(txtWordListFile.Text);

            if (fileTemplate.Exists && fileWordList.Exists)
            {
                app = new WordApp(txtTemplateFile.Text);
                
                sBytes = System.IO.File.ReadAllBytes(txtWordListFile.Text);
                sList = System.IO.File.ReadAllLines(txtWordListFile.Text, getEncoding());

                app.GoToEnd();
                app.SetAlignment(Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter);

                for (int i = 0; i < sList.Length; i++)
                {
                    sLine = sList[i].Split('\t');

                    app.TypeText(String.Format("{0}{1}", "\r\n", sLine[0]));
                    app.InsertPageBreak();
                    app.TypeText(String.Format("{0}{1}", "\r\n", sLine[1]));

                    if (i < sList.Length - 1)
                    {
                        app.InsertPageBreak();
                    }
                    else
                    {
                        app.TypeText("\r\n");
                    }
                }

                app.Visible = true;
            }
            else
            {
                MessageBox.Show("Please select template and list file.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnBrowseWordListFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtWordListFile.Text = ofd.FileName;
            }
        }

        private void btnBrowseTemplate_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtTemplateFile.Text = ofd.FileName;
            }
        }
    }
}
