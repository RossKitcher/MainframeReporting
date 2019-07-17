using System.ComponentModel;

namespace start
{
    partial class Main_Form
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
            this.runReportButton = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.browseFilesButton = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.reportStatusProgress = new System.Windows.Forms.ProgressBar();
            this.reportBgWorker = new System.ComponentModel.BackgroundWorker();
            this.reportStatusLabel = new System.Windows.Forms.Label();
            this.dataFormatButton = new System.Windows.Forms.Button();
            this.userGuideButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // runReportButton
            // 
            this.runReportButton.Location = new System.Drawing.Point(286, 208);
            this.runReportButton.Name = "runReportButton";
            this.runReportButton.Size = new System.Drawing.Size(241, 60);
            this.runReportButton.TabIndex = 0;
            this.runReportButton.Text = "Run Report";
            this.runReportButton.UseVisualStyleBackColor = true;
            this.runReportButton.Click += new System.EventHandler(this.runReportButton_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "xlsx";
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Excel files (*.xls*)|*.xls*";
            this.openFileDialog1.Title = "Browse Excel Files";
            // 
            // browseFilesButton
            // 
            this.browseFilesButton.Location = new System.Drawing.Point(72, 105);
            this.browseFilesButton.Name = "browseFilesButton";
            this.browseFilesButton.Size = new System.Drawing.Size(198, 49);
            this.browseFilesButton.TabIndex = 1;
            this.browseFilesButton.Text = "Browse Data (.xls)";
            this.browseFilesButton.UseVisualStyleBackColor = true;
            this.browseFilesButton.Click += new System.EventHandler(this.browseFilesButton_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(276, 105);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(439, 49);
            this.textBox1.TabIndex = 2;
            // 
            // reportStatusProgress
            // 
            this.reportStatusProgress.Location = new System.Drawing.Point(202, 308);
            this.reportStatusProgress.Name = "reportStatusProgress";
            this.reportStatusProgress.Size = new System.Drawing.Size(397, 23);
            this.reportStatusProgress.TabIndex = 4;
            // 
            // reportBgWorker
            // 
            this.reportBgWorker.WorkerReportsProgress = true;
            this.reportBgWorker.WorkerSupportsCancellation = true;
            this.reportBgWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.reportBgWorker_DoWork);
            this.reportBgWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.reportBgWorker_ProgressChanged);
            this.reportBgWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.reportBgWorker_RunWorkerCompleted);
            // 
            // reportStatusLabel
            // 
            this.reportStatusLabel.AutoSize = true;
            this.reportStatusLabel.Location = new System.Drawing.Point(361, 285);
            this.reportStatusLabel.Name = "reportStatusLabel";
            this.reportStatusLabel.Size = new System.Drawing.Size(98, 20);
            this.reportStatusLabel.TabIndex = 5;
            this.reportStatusLabel.Text = "Not Running";
            // 
            // dataFormatButton
            // 
            this.dataFormatButton.Location = new System.Drawing.Point(598, 398);
            this.dataFormatButton.Name = "dataFormatButton";
            this.dataFormatButton.Size = new System.Drawing.Size(190, 40);
            this.dataFormatButton.TabIndex = 7;
            this.dataFormatButton.Text = "Data format";
            this.dataFormatButton.UseVisualStyleBackColor = true;
            this.dataFormatButton.Click += new System.EventHandler(this.dataFormatButton_Click);
            // 
            // userGuideButton
            // 
            this.userGuideButton.Location = new System.Drawing.Point(12, 398);
            this.userGuideButton.Name = "userGuideButton";
            this.userGuideButton.Size = new System.Drawing.Size(190, 40);
            this.userGuideButton.TabIndex = 8;
            this.userGuideButton.Text = "User guide";
            this.userGuideButton.UseVisualStyleBackColor = true;
            this.userGuideButton.Click += new System.EventHandler(this.userGuideButton_Click);
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Location = new System.Drawing.Point(98, 362);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(600, 2);
            this.label1.TabIndex = 9;
            // 
            // label2
            // 
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.Location = new System.Drawing.Point(98, 181);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(600, 2);
            this.label2.TabIndex = 10;
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.SystemColors.Control;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F);
            this.textBox2.Location = new System.Drawing.Point(181, 12);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(418, 46);
            this.textBox2.TabIndex = 11;
            this.textBox2.Text = "Mainframe Reporting";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3
            // 
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label3.Location = new System.Drawing.Point(98, 75);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(600, 2);
            this.label3.TabIndex = 12;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.userGuideButton);
            this.Controls.Add(this.dataFormatButton);
            this.Controls.Add(this.reportStatusLabel);
            this.Controls.Add(this.reportStatusProgress);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.browseFilesButton);
            this.Controls.Add(this.runReportButton);
            this.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.Name = "Form1";
            this.Text = "Mainframe Reporting Script";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button runReportButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button browseFilesButton;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ProgressBar reportStatusProgress;
        private System.ComponentModel.BackgroundWorker reportBgWorker;
        private System.Windows.Forms.Label reportStatusLabel;
        private System.Windows.Forms.Button dataFormatButton;
        private System.Windows.Forms.Button userGuideButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label3;
    }
}

