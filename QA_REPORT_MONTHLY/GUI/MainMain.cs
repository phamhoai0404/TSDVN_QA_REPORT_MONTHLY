﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QA_REPORT_MONTHLY.FUNCTION;
using QA_REPORT_MONTHLY.MODEL;


namespace GUI
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }
        ActionInput1 inputAction1 = new ActionInput1();
        List<DataFirst> listData = new List<DataFirst>();

        #region Action Style
        /// <summary>
        /// Thuc hien set nut hanh dong trang thai
        /// </summary>
        /// <param name="action"></param>
        /// CreatedBy: HoaiPT(?/?/2022)
        private void actionButton(bool action)
        {
            if (action == true)
            {
                this.picExecute.Visible = false;
                this.picDone.Visible = true;
                this.pnlMainMain.Enabled = true;

                this.updateLable("Sẵn sàng thực hiện");
            }
            else
            {
                this.pnlMainMain.Enabled = false;

                this.picDone.Visible = false;
                this.picExecute.Visible = true;
            }
            this.picExecute.Update();
            this.picDone.Update();
        }
        /// <summary>
        /// Thuc hien update label 
        /// </summary>
        /// <param name="nameText">Ten label muon cap nhat</param>
        /// CreatedBy: HoaiPT(?/?/2022)
        private void updateLable(string nameText)
        {
            this.lblDisplay.Text = nameText;
            this.lblDisplay.Update();
        }
        #endregion
        private void frmMain_Load(object sender, EventArgs e)
        {
            try
            {

                Dictionary<string, object> getConfig = new Dictionary<string, object>();
                string result = MyFunction2.GetDataConfig(@"CONFIG\config_qa_report_monthly.xlsx", ref getConfig);
                if (!result.Equals(RESULT.OK))
                {
                    MessageBox.Show(result, "Error Load Form", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                    return;
                }
                this.GetConfig(getConfig);

                this.btnClearAll.PerformClick();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra vui lòng liên hệ bộ phận IT để được hỗ trợ!" + ex.Message, "Error Load Form", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.actionButton(true);
            }
        }
        private void GetConfig(Dictionary<string, object> getConfig)
        {
            DataConfig.CONFIG_SOURCE_FILE_DATA = getConfig["SourceFileData"].ToString();
            DataConfig.CONFIG_SOURCE_FILE_ERROR = getConfig["SourceFileError"].ToString();
            DataConfig.CONFIG_FILE_ERROR_SHEETNAME = getConfig["FileError_SheetName"].ToString();

            DataConfig.CONFIG_MONTH = DateTime.Now.ToString("MM");
        }

        private void btnActionMain_Click(object sender, EventArgs e)
        {
            try
            {
                this.actionButton(false);
                this.updateLable("Thực hiện validate");

                this.GetDataInput();//Thuc hien lay du lieu

                //Thuc hien validate gia tri nhap vao
                string resultValue = Action1.ValidateInputAction1(ref this.inputAction1);
                if (!resultValue.Equals(RESULT.OK))
                {
                    MessageBox.Show(resultValue, "Validate Action 1", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string sheetName = this.inputAction1.monthString + "." + DateTime.Now.ToString("yyyy");

                this.listData.Clear();
                resultValue = Action1.OpenFileExcel(this.inputAction1, sheetName, ref listData);
                if (!resultValue.Equals(RESULT.OK))
                {
                    MessageBox.Show(resultValue, "Get Data File", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                





















            }
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Có lỗi xảy ra vui lòng liên hệ bộ phận IT để được hỗ trợ!" + ex.Message, "Run Program", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            finally
            {
                this.actionButton(true);
            }
        }
        #region ClearAll
        private void GetDataInput()
        {
            this.inputAction1.monthString = this.txtMonth.Text;
            this.inputAction1.fileData = this.txtFileData.Text;
            this.inputAction1.fileError = this.txtFileError.Text;
        }
        private void btnClearAll_Click(object sender, EventArgs e)
        {
            this.SetDataFirst();
            this.SetAllCheck1();
        }
        private void SetAllCheck1()
        {
            this.chkTSB.Checked = true;
            this.chkFX.Checked = true;
            this.chkKyocera.Checked = true;
            this.chkHT.Checked = true;
            this.chkOkidenki.Checked = true;
            this.chkRiso.Checked = true;
            this.chkJCM.Checked = true;
        }
        private void SetDataFirst()
        {
            this.txtFileData.Text = DataConfig.CONFIG_SOURCE_FILE_DATA;
            this.txtFileError.Text = DataConfig.CONFIG_SOURCE_FILE_ERROR;
            this.txtMonth.Text = DataConfig.CONFIG_MONTH;
        }
        #endregion

        #region SelectFile
        private void ClickSelectFile(string typeClick)
        {
            string result = MyFunction2.SelectFile();
            switch (result)
            {
                case RESULT.OK:
                    return;
                case RESULT.ERROR_SELECT_FILE:
                    MessageBox.Show(RESULT.ERROR_SELECT_FILE, "Error Select File", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
            }

            switch (typeClick)
            {
                case MdlComment.CLICK_FILE_DATA:
                    this.txtFileData.Text = result;
                    break;
                case MdlComment.CLICK_FILE_ERROR:
                    this.txtFileError.Text = result;
                    break;
            }

        }

        private void btnSelectFileData_Click(object sender, EventArgs e)
        {
            this.ClickSelectFile(MdlComment.CLICK_FILE_DATA);
        }

        private void btnSelectFileError_Click(object sender, EventArgs e)
        {
            this.ClickSelectFile(MdlComment.CLICK_FILE_ERROR);
        }
        #endregion

        private void txtFileData_TextChanged(object sender, EventArgs e)
        {

        }
    }
}