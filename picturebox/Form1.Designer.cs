namespace picturebox
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
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.trackBar1 = new System.Windows.Forms.TrackBar();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.tbStartTime = new System.Windows.Forms.TextBox();
            this.tbEndTime = new System.Windows.Forms.TextBox();
            this.tbStartDate = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tbEndDate = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnNext = new System.Windows.Forms.Button();
            this.tbSeconds = new System.Windows.Forms.TextBox();
            this.tbAutoP = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tbTotalPlanned = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.tbTotalNotPlanned = new System.Windows.Forms.TextBox();
            this.btnPrev = new System.Windows.Forms.Button();
            this.tbArea = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.tbTime = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnChart = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(12, 7);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(100, 26);
            this.btnOpenFile.TabIndex = 0;
            this.btnOpenFile.Text = "Open File";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(118, 7);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(100, 26);
            this.btnClear.TabIndex = 3;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.Clear_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // trackBar1
            // 
            this.trackBar1.Location = new System.Drawing.Point(1089, 2);
            this.trackBar1.Maximum = 40;
            this.trackBar1.Minimum = 1;
            this.trackBar1.Name = "trackBar1";
            this.trackBar1.Size = new System.Drawing.Size(274, 56);
            this.trackBar1.TabIndex = 5;
            this.trackBar1.TickFrequency = 3;
            this.trackBar1.Value = 1;
            this.trackBar1.Scroll += new System.EventHandler(this.trackBar1_Scroll);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(3, 19);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(1348, 300);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            this.pictureBox1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.pictureBox1_MouseMove);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(12, 64);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1363, 334);
            this.panel1.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(1037, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 17);
            this.label1.TabIndex = 7;
            this.label1.Text = "Scale";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(433, 2);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 17);
            this.label2.TabIndex = 8;
            this.label2.Text = "Start time:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(438, 33);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(67, 17);
            this.label3.TabIndex = 9;
            this.label3.Text = "End time:";
            // 
            // tbStartTime
            // 
            this.tbStartTime.Location = new System.Drawing.Point(511, 3);
            this.tbStartTime.Name = "tbStartTime";
            this.tbStartTime.ReadOnly = true;
            this.tbStartTime.Size = new System.Drawing.Size(84, 22);
            this.tbStartTime.TabIndex = 10;
            // 
            // tbEndTime
            // 
            this.tbEndTime.Location = new System.Drawing.Point(511, 31);
            this.tbEndTime.Name = "tbEndTime";
            this.tbEndTime.ReadOnly = true;
            this.tbEndTime.Size = new System.Drawing.Size(84, 22);
            this.tbEndTime.TabIndex = 11;
            // 
            // tbStartDate
            // 
            this.tbStartDate.Location = new System.Drawing.Point(649, 3);
            this.tbStartDate.Name = "tbStartDate";
            this.tbStartDate.ReadOnly = true;
            this.tbStartDate.Size = new System.Drawing.Size(78, 22);
            this.tbStartDate.TabIndex = 13;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(601, 2);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(42, 17);
            this.label4.TabIndex = 12;
            this.label4.Text = "Date:";
            // 
            // tbEndDate
            // 
            this.tbEndDate.Location = new System.Drawing.Point(649, 31);
            this.tbEndDate.Name = "tbEndDate";
            this.tbEndDate.ReadOnly = true;
            this.tbEndDate.Size = new System.Drawing.Size(78, 22);
            this.tbEndDate.TabIndex = 15;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(601, 33);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(42, 17);
            this.label5.TabIndex = 14;
            this.label5.Text = "Date:";
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(330, 7);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(100, 26);
            this.btnNext.TabIndex = 16;
            this.btnNext.Text = "Next";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.bNext_Click);
            // 
            // tbSeconds
            // 
            this.tbSeconds.Location = new System.Drawing.Point(152, 404);
            this.tbSeconds.Name = "tbSeconds";
            this.tbSeconds.ReadOnly = true;
            this.tbSeconds.Size = new System.Drawing.Size(78, 22);
            this.tbSeconds.TabIndex = 17;
            // 
            // tbAutoP
            // 
            this.tbAutoP.Location = new System.Drawing.Point(367, 402);
            this.tbAutoP.Name = "tbAutoP";
            this.tbAutoP.ReadOnly = true;
            this.tbAutoP.Size = new System.Drawing.Size(78, 22);
            this.tbAutoP.TabIndex = 18;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(12, 407);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(134, 17);
            this.label6.TabIndex = 19;
            this.label6.Text = "Total Process Time:";
            // 
            // tbTotalPlanned
            // 
            this.tbTotalPlanned.Location = new System.Drawing.Point(367, 430);
            this.tbTotalPlanned.Name = "tbTotalPlanned";
            this.tbTotalPlanned.ReadOnly = true;
            this.tbTotalPlanned.Size = new System.Drawing.Size(78, 22);
            this.tbTotalPlanned.TabIndex = 20;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(236, 407);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(80, 17);
            this.label7.TabIndex = 21;
            this.label7.Text = "Auto Mode:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(236, 433);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(97, 17);
            this.label8.TabIndex = 22;
            this.label8.Text = "Planned Stop:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(236, 463);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(123, 17);
            this.label9.TabIndex = 24;
            this.label9.Text = "Not Planned Stop:";
            // 
            // tbTotalNotPlanned
            // 
            this.tbTotalNotPlanned.Location = new System.Drawing.Point(367, 460);
            this.tbTotalNotPlanned.Name = "tbTotalNotPlanned";
            this.tbTotalNotPlanned.ReadOnly = true;
            this.tbTotalNotPlanned.Size = new System.Drawing.Size(78, 22);
            this.tbTotalNotPlanned.TabIndex = 23;
            // 
            // btnPrev
            // 
            this.btnPrev.Location = new System.Drawing.Point(224, 7);
            this.btnPrev.Name = "btnPrev";
            this.btnPrev.Size = new System.Drawing.Size(100, 26);
            this.btnPrev.TabIndex = 25;
            this.btnPrev.Text = "Previous";
            this.btnPrev.UseVisualStyleBackColor = true;
            this.btnPrev.Click += new System.EventHandler(this.bPrev_Click);
            // 
            // tbArea
            // 
            this.tbArea.Location = new System.Drawing.Point(780, 428);
            this.tbArea.Name = "tbArea";
            this.tbArea.ReadOnly = true;
            this.tbArea.Size = new System.Drawing.Size(78, 22);
            this.tbArea.TabIndex = 27;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(707, 430);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(67, 17);
            this.label10.TabIndex = 28;
            this.label10.Text = "% of total";
            // 
            // tbTime
            // 
            this.tbTime.Location = new System.Drawing.Point(780, 400);
            this.tbTime.Name = "tbTime";
            this.tbTime.ReadOnly = true;
            this.tbTime.Size = new System.Drawing.Size(78, 22);
            this.tbTime.TabIndex = 29;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(707, 401);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(43, 17);
            this.label11.TabIndex = 30;
            this.label11.Text = "Time:";
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(12, 39);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(100, 26);
            this.btnImport.TabIndex = 31;
            this.btnImport.Text = "Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnChart
            // 
            this.btnChart.Location = new System.Drawing.Point(118, 39);
            this.btnChart.Name = "btnChart";
            this.btnChart.Size = new System.Drawing.Size(100, 26);
            this.btnChart.TabIndex = 32;
            this.btnChart.Text = "Chart";
            this.btnChart.UseVisualStyleBackColor = true;
            this.btnChart.Click += new System.EventHandler(this.btnChart_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(1203, 407);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(172, 26);
            this.btnDelete.TabIndex = 33;
            this.btnDelete.Text = "Delete file on server";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(224, 39);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(206, 26);
            this.btnExport.TabIndex = 34;
            this.btnExport.Text = "Export data to excel";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1387, 494);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnChart);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.tbTime);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.tbArea);
            this.Controls.Add(this.btnPrev);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.tbTotalNotPlanned);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.tbTotalPlanned);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.tbAutoP);
            this.Controls.Add(this.tbSeconds);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.tbEndDate);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.tbStartDate);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tbEndTime);
            this.Controls.Add(this.tbStartTime);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnOpenFile);
            this.Controls.Add(this.trackBar1);
            this.Controls.Add(this.panel1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Process Time";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.TrackBar trackBar1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbStartTime;
        private System.Windows.Forms.TextBox tbEndTime;
        private System.Windows.Forms.TextBox tbStartDate;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbEndDate;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.TextBox tbSeconds;
        private System.Windows.Forms.TextBox tbAutoP;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbTotalPlanned;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox tbTotalNotPlanned;
        private System.Windows.Forms.Button btnPrev;
        private System.Windows.Forms.TextBox tbArea;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox tbTime;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnChart;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnExport;
    }
}

