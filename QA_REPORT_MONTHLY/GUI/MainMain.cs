using System;
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
        ActionInput2 valueInput2 = new ActionInput2();
        List<DataFirst> listData = new List<DataFirst>();
        List<DataError> listError = new List<DataError>();

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
                this.SetDataFirst2();//Thuc hien set du lieu cho Action 2
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
            DataConfig.CONFIG_FILE_ERROR_PASSWORD = getConfig["FileError_Password"].ToString();
            DataConfig.CONFIG_FILE_TEMPLATE = getConfig["FileTemplate"].ToString();

            //DataConfig.CONFIG_MONTH = DateTime.Now.ToString("MM");
            DataConfig.CONFIG_MONTH = "02";

            DataConfig.CONFIG_2_COLUMM_MODEL = getConfig["2ColumnMode"].ToString();

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

                if (this.chkTSB.Checked == false &&
                   this.chkFX.Checked == false &&
                   this.chkKyocera.Checked == false &&
                   this.chkHT.Checked == false &&
                   this.chkOkidenki.Checked == false &&
                   this.chkRiso.Checked == false &&
                   this.chkJCM.Checked == false)
                {
                    MessageBox.Show("Bạn không thực hiện chọn báo cáo nào!", "Not Select Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }



                string sheetName = this.inputAction1.monthString + "." + DateTime.Now.ToString("yyyy");

                this.updateLable("Lấy dữ liệu file data...");
                this.listData.Clear();
                resultValue = Action1.OpenFileExcelData(this.inputAction1, sheetName, ref listData);
                if (!resultValue.Equals(RESULT.OK))
                {
                    MessageBox.Show(resultValue, "Get Data File", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                this.updateLable("Lấy dữ liệu file lỗi.....");
                this.listError.Clear();
                resultValue = Action1.OpenFileExcelError(this.inputAction1, ref this.listError);
                if (!resultValue.Equals(RESULT.OK))
                {
                    MessageBox.Show(resultValue, "Get File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                this.updateLable("Ghép khách hàng cho dữ liệu lỗi");

                resultValue = Action1.ActionFileError(this.listData, ref this.listError);
                if (!resultValue.Equals(RESULT.OK))
                {
                    MessageBox.Show(resultValue, "Get File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                //Thuc hien  tao file
                resultValue = ActionWrite.CreateFile();
                if (!resultValue.Equals(RESULT.OK))
                {
                    MessageBox.Show(resultValue, "Create File", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string hoa = DateTime.Now.ToString("hh:mm:ss");

                if (this.chkTSB.Checked == true)
                {
                    this.updateLable("Thực hiện lấy dữ liệu TSB");
                    List<DataTSB> listTSB = new List<DataTSB>();
                    resultValue = Action1.GetTSB(listData, listError, ref listTSB);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Action TSB", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    this.updateLable("Thực hiện ghi dữ liệu TSB");
                    resultValue = ActionWrite.WriteTSB1(listTSB);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Write TSB", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                if (this.chkKyocera.Checked == true)
                {
                    this.updateLable("Thực hiện lấy dữ liệu Kyocera");
                    List<DataKyocera> listKyocera = new List<DataKyocera>();
                    resultValue = Action1.GetKyocera(listData, listError, ref listKyocera);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Action Kyocera", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    this.updateLable("Thực hiện ghi dữ liệu Kyocera");
                    resultValue = ActionWrite.WriteKyocera1(listKyocera);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Write Kyocera", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                if (this.chkFX.Checked == true)
                {
                    this.updateLable("Thực hiện lấy dữ liệu Kyocera");
                    List<DataFX> listFX = new List<DataFX>();
                    resultValue = Action1.GetFX(listData, listError, ref listFX);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Action FX", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    this.updateLable("Thực hiện ghi dữ liệu FX");
                    resultValue = ActionWrite.WriteFX(listFX);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Write FX", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                if (this.chkHT.Checked == true)
                {
                    this.updateLable("Thực hiện lấy dữ liệu Hitachi");

                    DataHT valueHT = new DataHT();
                    resultValue = Action1.GetHT(listData, listError, ref valueHT);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Action HT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    this.updateLable("Thực hiện ghi dữ liệu HT");
                    resultValue = ActionWrite.WriteHT(valueHT, this.txtMonth.Text);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Write FX", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                if (this.chkOkidenki.Checked == true)
                {
                    this.updateLable("Thực hiện lấy dữ liệu OKIDENKI");

                    DataOkidenki valueOkidenki = new DataOkidenki();
                    resultValue = Action1.GetOkidenki(listData, listError, ref valueOkidenki);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Action OKIDENKI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    this.updateLable("Thực hiện ghi dữ liệu OKIDENKI");
                    resultValue = ActionWrite.WriteOkidenki(valueOkidenki, this.txtMonth.Text);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Write OKIDENKI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                if (this.chkRiso.Checked == true)
                {
                    this.updateLable("Thực hiện lấy dữ liệu RISO");

                    DataRiso valueRiso = new DataRiso();
                    resultValue = Action1.GetRISO(listData, listError, ref valueRiso);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Action RISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    this.updateLable("Thực hiện ghi dữ liệu RISO");
                    resultValue = ActionWrite.WriteRiso(valueRiso, this.txtMonth.Text);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Write RISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }




                if (this.chkJCM.Checked == true)
                {
                    this.updateLable("Thực hiện lấy dữ liệu JCM");

                    DataJCM valueRiso = new DataJCM();
                    resultValue = Action1.GetJCM(listData, listError, ref valueRiso);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Action JCM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    this.updateLable("Thực hiện ghi dữ liệu JCM");
                    resultValue = ActionWrite.WriteJCM(valueRiso, this.txtMonth.Text);
                    if (!resultValue.Equals(RESULT.OK))
                    {
                        MessageBox.Show(resultValue, "Get Write JCM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }























            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra vui lòng liên hệ bộ phận IT để được hỗ trợ!" + ex.Message, "Run Program", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                MyFunction2.Skill_Process("Excel");
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
            this.SetDataFirst1();
            this.SetAllCheck1();
        }
        private void btn2RefreshAll_Click(object sender, EventArgs e)
        {
            this.txt2RowEnd.Clear();
            this.txt2RowStart.Clear();

            this.SetDataFirst2();
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
        private void SetDataFirst1()
        {
            this.txtFileData.Text = DataConfig.CONFIG_SOURCE_FILE_DATA;
            this.txtFileError.Text = DataConfig.CONFIG_SOURCE_FILE_ERROR;
            this.txtMonth.Text = DataConfig.CONFIG_MONTH;
        }
        private void SetDataFirst2()
        {
            
            this.txt2ColModel.Text = DataConfig.CONFIG_2_COLUMM_MODEL;

            //Du lieu test o day xoa di nha
            this.txt2FileData.Text = @"P:\96. Share Data\99. Other\13. IT\HOAI\QA_REPORT\2023.01_Kyocera様月報 - CUT.xlsx";
            this.txt2RowEnd.Text = "488";
            this.txt2RowStart.Text = "411";
            this.txt2SheetName.Text = "部品コード";
            this.tabMain.SelectedIndex = 1;
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


        private void GetDataInput_2()
        {
            this.valueInput2.rowEndString = this.txt2RowEnd.Text;
            this.valueInput2.rowStartString = this.txt2RowStart.Text;
            this.valueInput2.fileData = this.txt2FileData.Text;
            this.valueInput2.colModel = this.txt2ColModel.Text;
            this.valueInput2.sheetName = this.txt2SheetName.Text;

            
        }
        private void btn2ActionMain_Click(object sender, EventArgs e)
        {
            try
            {
                this.actionButton(false);
                this.updateLable("Thực hiện validate dữ liệu");
                this.GetDataInput_2();//Thuc hien lay du lieu
                
                string resultTemp = Action2.ValidateInputAction2(ref this.valueInput2);
                if (!resultTemp.Equals(RESULT.OK))
                {
                    MessageBox.Show(resultTemp, "Validate Input Action 2", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                this.updateLable("Lấy dữ liệu file Kyocera...");
                List<DataKyocera> listKyocrea = new List<DataKyocera>();
                resultTemp = Action2.GetKyoceraOld(this.valueInput2, ref listKyocrea);
                if (!resultTemp.Equals(RESULT.OK))
                {
                    MessageBox.Show(resultTemp, "Get Data Action 2", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                this.updateLable("Xử lý dữ liệu Kyocera....");
                resultTemp = Action2.ExecuteKyocera(ref listKyocrea);
                if (!resultTemp.Equals(RESULT.OK))
                {
                    MessageBox.Show(resultTemp, "Execute Data Kyocera", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                resultTemp = Action2.WriteKyocera_2(listKyocrea);

                string k = "0";



            }
            finally
            {
                this.actionButton(true);
            }
        }

        
    }
}
