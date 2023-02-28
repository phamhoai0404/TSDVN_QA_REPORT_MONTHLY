
namespace GUI
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.panel1 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.pnlMainMain = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblDisplay = new System.Windows.Forms.Label();
            this.picExecute = new System.Windows.Forms.PictureBox();
            this.picDone = new System.Windows.Forms.PictureBox();
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.btnActionMain = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.chkJCM = new System.Windows.Forms.CheckBox();
            this.chkRiso = new System.Windows.Forms.CheckBox();
            this.chkOkidenki = new System.Windows.Forms.CheckBox();
            this.chkHT = new System.Windows.Forms.CheckBox();
            this.chkKyocera = new System.Windows.Forms.CheckBox();
            this.chkFX = new System.Windows.Forms.CheckBox();
            this.chkTSB = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnSelectFileError = new System.Windows.Forms.Button();
            this.btnSelectFileData = new System.Windows.Forms.Button();
            this.txtFileError = new System.Windows.Forms.TextBox();
            this.txtMonth = new System.Windows.Forms.TextBox();
            this.txtFileData = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.pnlMainMain.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picExecute)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picDone)).BeginInit();
            this.tabMain.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(720, 89);
            this.panel1.TabIndex = 0;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(161, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(448, 42);
            this.label3.TabIndex = 9;
            this.label3.Text = "QA REPORT MONTHLY";
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.panel2.Controls.Add(this.pictureBox1);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Location = new System.Drawing.Point(12, 6);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(77, 67);
            this.panel2.TabIndex = 8;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(5, 7);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(64, 42);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 7);
            this.label2.TabIndex = 2;
            this.label2.Text = "Taishodo Việt Nam";
            // 
            // pnlMainMain
            // 
            this.pnlMainMain.Controls.Add(this.groupBox1);
            this.pnlMainMain.Controls.Add(this.tabMain);
            this.pnlMainMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlMainMain.Location = new System.Drawing.Point(0, 89);
            this.pnlMainMain.Name = "pnlMainMain";
            this.pnlMainMain.Size = new System.Drawing.Size(720, 526);
            this.pnlMainMain.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblDisplay);
            this.groupBox1.Controls.Add(this.picExecute);
            this.groupBox1.Controls.Add(this.picDone);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox1.Location = new System.Drawing.Point(0, 486);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(720, 40);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // lblDisplay
            // 
            this.lblDisplay.AutoSize = true;
            this.lblDisplay.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDisplay.Location = new System.Drawing.Point(38, 16);
            this.lblDisplay.Name = "lblDisplay";
            this.lblDisplay.Size = new System.Drawing.Size(99, 13);
            this.lblDisplay.TabIndex = 18;
            this.lblDisplay.Text = "Sẵn sàng thực hiện";
            // 
            // picExecute
            // 
            this.picExecute.Image = ((System.Drawing.Image)(resources.GetObject("picExecute.Image")));
            this.picExecute.Location = new System.Drawing.Point(12, 19);
            this.picExecute.Name = "picExecute";
            this.picExecute.Size = new System.Drawing.Size(17, 12);
            this.picExecute.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picExecute.TabIndex = 16;
            this.picExecute.TabStop = false;
            // 
            // picDone
            // 
            this.picDone.Image = ((System.Drawing.Image)(resources.GetObject("picDone.Image")));
            this.picDone.Location = new System.Drawing.Point(16, 18);
            this.picDone.Name = "picDone";
            this.picDone.Size = new System.Drawing.Size(10, 12);
            this.picDone.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picDone.TabIndex = 17;
            this.picDone.TabStop = false;
            // 
            // tabMain
            // 
            this.tabMain.Controls.Add(this.tabPage1);
            this.tabMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabMain.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabMain.Location = new System.Drawing.Point(0, 0);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(720, 526);
            this.tabMain.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.btnClearAll);
            this.tabPage1.Controls.Add(this.btnActionMain);
            this.tabPage1.Controls.Add(this.groupBox4);
            this.tabPage1.Controls.Add(this.groupBox3);
            this.tabPage1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage1.Location = new System.Drawing.Point(4, 33);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(712, 489);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Hoạt động chính";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // btnClearAll
            // 
            this.btnClearAll.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnClearAll.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnClearAll.BackgroundImage")));
            this.btnClearAll.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnClearAll.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnClearAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnClearAll.ForeColor = System.Drawing.Color.Navy;
            this.btnClearAll.Location = new System.Drawing.Point(25, 379);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(60, 54);
            this.btnClearAll.TabIndex = 103;
            this.btnClearAll.UseVisualStyleBackColor = false;
            this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
            // 
            // btnActionMain
            // 
            this.btnActionMain.Location = new System.Drawing.Point(105, 379);
            this.btnActionMain.Name = "btnActionMain";
            this.btnActionMain.Size = new System.Drawing.Size(581, 54);
            this.btnActionMain.TabIndex = 1;
            this.btnActionMain.Text = "THỰC HIỆN";
            this.btnActionMain.UseVisualStyleBackColor = true;
            this.btnActionMain.Click += new System.EventHandler(this.btnActionMain_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox4.Controls.Add(this.chkJCM);
            this.groupBox4.Controls.Add(this.chkRiso);
            this.groupBox4.Controls.Add(this.chkOkidenki);
            this.groupBox4.Controls.Add(this.chkHT);
            this.groupBox4.Controls.Add(this.chkKyocera);
            this.groupBox4.Controls.Add(this.chkFX);
            this.groupBox4.Controls.Add(this.chkTSB);
            this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.Location = new System.Drawing.Point(25, 233);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(661, 127);
            this.groupBox4.TabIndex = 0;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Lựa chọn:";
            // 
            // chkJCM
            // 
            this.chkJCM.AutoSize = true;
            this.chkJCM.Location = new System.Drawing.Point(243, 84);
            this.chkJCM.Name = "chkJCM";
            this.chkJCM.Size = new System.Drawing.Size(108, 20);
            this.chkJCM.TabIndex = 0;
            this.chkJCM.Text = "Báo cáo JCM";
            this.chkJCM.UseVisualStyleBackColor = true;
            // 
            // chkRiso
            // 
            this.chkRiso.AutoSize = true;
            this.chkRiso.Location = new System.Drawing.Point(243, 58);
            this.chkRiso.Name = "chkRiso";
            this.chkRiso.Size = new System.Drawing.Size(113, 20);
            this.chkRiso.TabIndex = 0;
            this.chkRiso.Text = "Báo cáo RISO";
            this.chkRiso.UseVisualStyleBackColor = true;
            // 
            // chkOkidenki
            // 
            this.chkOkidenki.AutoSize = true;
            this.chkOkidenki.Location = new System.Drawing.Point(243, 32);
            this.chkOkidenki.Name = "chkOkidenki";
            this.chkOkidenki.Size = new System.Drawing.Size(142, 20);
            this.chkOkidenki.TabIndex = 0;
            this.chkOkidenki.Text = "Báo cáo OKIDENKI";
            this.chkOkidenki.UseVisualStyleBackColor = true;
            // 
            // chkHT
            // 
            this.chkHT.AutoSize = true;
            this.chkHT.Location = new System.Drawing.Point(50, 84);
            this.chkHT.Name = "chkHT";
            this.chkHT.Size = new System.Drawing.Size(134, 20);
            this.chkHT.TabIndex = 0;
            this.chkHT.Text = "Báo cáo HITACHI";
            this.chkHT.UseVisualStyleBackColor = true;
            // 
            // chkKyocera
            // 
            this.chkKyocera.AutoSize = true;
            this.chkKyocera.Location = new System.Drawing.Point(449, 32);
            this.chkKyocera.Name = "chkKyocera";
            this.chkKyocera.Size = new System.Drawing.Size(145, 20);
            this.chkKyocera.TabIndex = 0;
            this.chkKyocera.Text = "Báo cáo KYOCERA";
            this.chkKyocera.UseVisualStyleBackColor = true;
            // 
            // chkFX
            // 
            this.chkFX.AutoSize = true;
            this.chkFX.Location = new System.Drawing.Point(50, 58);
            this.chkFX.Name = "chkFX";
            this.chkFX.Size = new System.Drawing.Size(97, 20);
            this.chkFX.TabIndex = 0;
            this.chkFX.Text = "Báo cáo FX";
            this.chkFX.UseVisualStyleBackColor = true;
            // 
            // chkTSB
            // 
            this.chkTSB.AutoSize = true;
            this.chkTSB.Location = new System.Drawing.Point(50, 32);
            this.chkTSB.Name = "chkTSB";
            this.chkTSB.Size = new System.Drawing.Size(108, 20);
            this.chkTSB.TabIndex = 0;
            this.chkTSB.Text = "Báo cáo TSB";
            this.chkTSB.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.Controls.Add(this.btnSelectFileError);
            this.groupBox3.Controls.Add(this.btnSelectFileData);
            this.groupBox3.Controls.Add(this.txtFileError);
            this.groupBox3.Controls.Add(this.txtMonth);
            this.groupBox3.Controls.Add(this.txtFileData);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(25, 19);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(661, 198);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Nguồn dữ liệu";
            // 
            // btnSelectFileError
            // 
            this.btnSelectFileError.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnSelectFileError.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSelectFileError.BackgroundImage")));
            this.btnSelectFileError.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnSelectFileError.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSelectFileError.Location = new System.Drawing.Point(622, 149);
            this.btnSelectFileError.Name = "btnSelectFileError";
            this.btnSelectFileError.Size = new System.Drawing.Size(33, 29);
            this.btnSelectFileError.TabIndex = 102;
            this.btnSelectFileError.UseVisualStyleBackColor = false;
            this.btnSelectFileError.Click += new System.EventHandler(this.btnSelectFileError_Click);
            // 
            // btnSelectFileData
            // 
            this.btnSelectFileData.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnSelectFileData.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSelectFileData.BackgroundImage")));
            this.btnSelectFileData.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnSelectFileData.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSelectFileData.Location = new System.Drawing.Point(622, 90);
            this.btnSelectFileData.Name = "btnSelectFileData";
            this.btnSelectFileData.Size = new System.Drawing.Size(33, 29);
            this.btnSelectFileData.TabIndex = 102;
            this.btnSelectFileData.UseVisualStyleBackColor = false;
            this.btnSelectFileData.Click += new System.EventHandler(this.btnSelectFileData_Click);
            // 
            // txtFileError
            // 
            this.txtFileError.Location = new System.Drawing.Point(12, 152);
            this.txtFileError.Name = "txtFileError";
            this.txtFileError.Size = new System.Drawing.Size(604, 22);
            this.txtFileError.TabIndex = 1;
            // 
            // txtMonth
            // 
            this.txtMonth.Location = new System.Drawing.Point(118, 39);
            this.txtMonth.Name = "txtMonth";
            this.txtMonth.Size = new System.Drawing.Size(40, 22);
            this.txtMonth.TabIndex = 1;
            this.txtMonth.TextChanged += new System.EventHandler(this.txtFileData_TextChanged);
            // 
            // txtFileData
            // 
            this.txtFileData.Location = new System.Drawing.Point(12, 93);
            this.txtFileData.Name = "txtFileData";
            this.txtFileData.Size = new System.Drawing.Size(604, 22);
            this.txtFileData.TabIndex = 1;
            this.txtFileData.TextChanged += new System.EventHandler(this.txtFileData_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 133);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(54, 16);
            this.label5.TabIndex = 0;
            this.label5.Text = "File Lỗi:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(103, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Tháng báo cáo:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 74);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 16);
            this.label6.TabIndex = 0;
            this.label6.Text = "File Dữ liệu:";
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(720, 615);
            this.Controls.Add(this.pnlMainMain);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "QA Report Monthly";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.pnlMainMain.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picExecute)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picDone)).EndInit();
            this.tabMain.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel pnlMainMain;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TabControl tabMain;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label lblDisplay;
        private System.Windows.Forms.PictureBox picExecute;
        private System.Windows.Forms.PictureBox picDone;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox txtFileError;
        private System.Windows.Forms.TextBox txtFileData;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnSelectFileError;
        private System.Windows.Forms.Button btnSelectFileData;
        private System.Windows.Forms.CheckBox chkOkidenki;
        private System.Windows.Forms.CheckBox chkHT;
        private System.Windows.Forms.CheckBox chkKyocera;
        private System.Windows.Forms.CheckBox chkFX;
        private System.Windows.Forms.CheckBox chkTSB;
        private System.Windows.Forms.CheckBox chkJCM;
        private System.Windows.Forms.CheckBox chkRiso;
        private System.Windows.Forms.Button btnClearAll;
        private System.Windows.Forms.Button btnActionMain;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtMonth;
    }
}

