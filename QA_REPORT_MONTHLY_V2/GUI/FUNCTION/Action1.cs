using Microsoft.Office.Interop.Excel;
using QA_REPORT_MONTHLY.MODEL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace QA_REPORT_MONTHLY.FUNCTION
{
    public class Action1
    {
        public static string ValidateInputAction1(ref ActionInput1 input)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(input.monthString) ||
                    string.IsNullOrWhiteSpace(input.fileData) ||
                    string.IsNullOrWhiteSpace(input.fileError))
                {
                    return RESULT.ERROR_NOT_NULL;
                }
                input.monthString = input.monthString.Trim();
                input.fileData = input.fileData.Trim();
                input.fileError = input.fileError.Trim();

                int tempMonth;
                if (!int.TryParse(input.monthString, out tempMonth))
                {
                    return RESULT.ERROR_NOT_NUMBER;
                }

                if (!(tempMonth >= 1 && tempMonth <= 12))
                {
                    return RESULT.ERROR_NOT_MONTH;
                }

                input.monthInt = tempMonth;
                input.monthString = String.Format("{0:00}", tempMonth);
                if (!File.Exists(input.fileData))
                {
                    return string.Format(RESULT.ERROR_NOT_FILE, input.fileData);
                }
                if (!File.Exists(input.fileError))
                {
                    return string.Format(RESULT.ERROR_NOT_FILE, input.fileError);
                }

                return RESULT.OK;
            }
            catch (Exception ex)
            {

                return string.Format(RESULT.ERROR_015_CATCH, "ValidateInputAction1", ex.Message);
            }
        }

        /// <summary>
        /// Hanh dong lay du lieu cua file Data (G2)
        /// </summary>
        /// <param name="valueInput"></param>
        /// <param name="sheetName"></param>
        /// <param name="listData"></param>
        /// <returns></returns>
        /// CreatedBy: HoaiPT(06/03/2023)
        public static string OpenFileExcelData(ActionInput1 valueInput, string sheetName, ref List<DataFirst> listData)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            //try
            //{
                app = new Excel.Application();
                wb = app.Workbooks.Open(valueInput.fileData, ReadOnly: true);

                //Kiem tra su ton tai cua sheetname thang
                bool existSheetName = false;
                foreach (Excel._Worksheet sheet in wb.Worksheets)
                {
                    if (sheet.Name.Equals(sheetName))
                    {
                        existSheetName = true;
                        break;
                    }
                }
                if (existSheetName == false)
                {
                    return string.Format(RESULT.ERROR_SHEETNAME, sheetName);
                }
                ws = wb.Sheets[sheetName];

                int rowCurrent = 4;//Du lieu lay bat dau tu dong thu 4
                string check = ws.Cells[rowCurrent, "D"].Value;
                DataFirst temp = new DataFirst();
                while (!string.IsNullOrWhiteSpace(check))//Kiem  tra xem dong WO co ton  tai hay khong neu khong ton tai thi dung du lieu
                {
                    temp.wo = check;//Du lieu WO
                    temp.model = ws.Cells[rowCurrent, "B"].Value;//Du lieu model
                    try
                    {
                        temp.cusDetail = ws.Cells[rowCurrent, "C"].Value;
                    }
                    catch (Exception exs)//Neu nhay vao catch thi chung to chua co chi tiet cua ma khach hang
                    {
                        return string.Format(RESULT.ERROR_COLUMN_C, rowCurrent, exs.Message);
                    }

                    temp.qty = Convert.ToInt64(ws.Cells[rowCurrent, "F"].Value);//Lay so luong nhap vao cua WO

                    try
                    {
                        temp.cusCode = ws.Cells[rowCurrent, "G"].Value;//khach hang o cot G neu = 0 thi khong phai khach hang nhay vao catch
                    }
                    catch (Exception exs)
                    {
                        return string.Format(RESULT.ERROR_COLUMN_G, rowCurrent, exs.Message);
                    }

                    listData.Add(new DataFirst(temp));//Thuc hien luu du lieu

                    rowCurrent++;//Tang dong len
                    check = ws.Cells[rowCurrent, "D"].Value;//Gan vao check
                }

                wb.Close(false);
                app.Quit();
                return RESULT.OK;
            //}
            //catch (Exception ex)
            //{
            //    return string.Format(RESULT.ERROR_015_CATCH, "OpenFileExcelData 1",  ex.Message);
            //}
            //finally
            //{

            //    if (ws != null)
            //    {
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            //    }
            //    if (wb != null)
            //    {
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            //    }
            //    if (app != null)
            //    {
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            //    }
            //}
        }
        public static string OpenFileExcelError(ActionInput1 valueInput, ref List<DataError> listError)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(valueInput.fileError);
                ws = wb.Sheets[DataConfig.CONFIG_FILE_ERROR_SHEETNAME];//Thuc hien gan vao ws

                //Thuc hien lay dong cuoi cung co du lieu
                if (ws.AutoFilter != null)//Loai bo loc neu co
                {
                    ws.Unprotect(DataConfig.CONFIG_FILE_ERROR_PASSWORD);
                    ws.AutoFilterMode = false;
                }
                int lastRow = ws.Cells[ws.Rows.Count, "P"].End(Excel.XlDirection.xlUp).Row;//Thuc hien lay dong cuoi cung

                DataError err = new DataError();//Check du lieu bat dau tu dong 11
                for (int i = 11; i <= lastRow; i++)
                {
                    //Neu dong du lieu Model trong thi duyet sang dong khac
                    err.model = ws.Cells[i, "P"].value;
                    if (string.IsNullOrWhiteSpace(err.model))
                    {
                        continue;
                    }

                    //Duyet bo phan mac loi
                    err.dept = Convert.ToString(ws.Cells[i, "AA"].value);
                    if (string.IsNullOrWhiteSpace(err.dept))
                    {
                        continue;
                    }
                    if (err.dept.Contains("+") || err.dept.Contains("Các"))
                    {
                        continue;
                    }

                    //Thuc hien lay ten loi
                    err.nameError = Convert.ToString(ws.Cells[i, "T"].value);
                    if (string.IsNullOrWhiteSpace(err.nameError))
                    {
                        continue;
                    }

                    //Lay so luong loi
                    int tempQty;
                    if (!int.TryParse(Convert.ToString(ws.Cells[i, "Z"].value), out tempQty))
                    {
                        continue;//Neu khong phai la so thi chuyen sang dong tiep theo
                    }
                    if (tempQty <= 0)
                    {
                        continue;//Neu so luong <=0 thi chuyen sang dong tiep theo
                    }
                    err.qty = tempQty;

                    err.wo = Convert.ToString(ws.Cells[i, "Q"].value);
                    if (string.IsNullOrWhiteSpace(err.wo))
                    {
                        continue;
                    }

                    if (err.nameError.Equals(MdlComment.TYPE_ERROR_THUA_THIEU_LK))
                    {
                        Comment comment = ws.Cells[i, "T"].Comment;
                        if (comment == null)
                        {
                            return string.Format(RESULT.ERROR_FILE_ERROR_NOT_COMMENT, i);
                        }
                        else
                        {
                            string temp = comment.Text();
                            switch (temp)
                            {
                                case string s when s.IndexOf(MdlComment.TYPE_ERROR_CHILD_THIEU, StringComparison.OrdinalIgnoreCase) >= 0:
                                    err.typeThua = false;
                                    break;
                                case string s when s.IndexOf(MdlComment.TYPE_ERROR_CHILD_THUA, StringComparison.OrdinalIgnoreCase) >= 0:
                                    err.typeThua = true;
                                    break;
                                default:
                                    return string.Format(RESULT.ERROR_FILE_ERROR_COMMENT_NOT_RULE, i, comment.Text());

                            }
                        }
                    }

                    listError.Add(new DataError(err));
                }

                wb.Close(false);
                app.Quit();
                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "OpenFileExcelError", ex.Message);
            }
            finally
            {
                if (ws != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                }
                if (wb != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                }
                if (app != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                }
            }
        }
        /// <summary>
        /// Thuc hien ghep KHAC HANG cho tat cac loi
        /// </summary>
        /// <param name="listData"></param>
        /// <param name="listError"></param>
        /// <returns></returns>
        /// CreatedBy: HoaiPT(06/03/2023)
        public static string ActionFileError(List<DataFirst> listData, ref List<DataError> listError)
        {
            try
            {
                foreach (var itemErr in listError.ToArray())
                {
                    int indexCheck = -1;
                    for (int i = 0; i < listData.Count(); i++)
                    {
                        if (listData[i].wo.Equals(itemErr.wo))
                        {
                            indexCheck = i;
                            break;
                        }
                    }

                    //Neu khong ton tai WO thi dung lai thong bao loi
                    if (indexCheck == -1)
                    {
                        return string.Format(RESULT.ERROR_FILE_ERROR_WO, itemErr.wo, itemErr.ToString());
                    }

                    itemErr.cusCode = listData[indexCheck].cusCode;
                }
                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "ActionFileError 1", ex.Message);
            }
        }

        public static string GetTSB(List<DataFirst> listData, List<DataError> listError, ref List<DataTSB> listTSB)
        {
            try
            {
                var listChildData = listData.Where(x => x.cusCode.Equals("TSB")).ToList();//Lay du lieu trong file ket qua
                var listChildErr = listError.Where(x => x.cusCode.Equals("TSB")).ToList();//Lay du lieu trong file loi

                //Phan lay du lieu cua model xong roi
                foreach (var item in listChildData.ToArray())
                {
                    string tempModel = item.model.Substring(0, 9);//Thuc hien model khong lay ver
                    var check = listTSB.FirstOrDefault(p => p.item.Equals(tempModel));//Kiem tra xem model da ton tai trong listTSB
                    if (check != null)
                    {
                        continue;
                    }//Neu ton tai roi thi chuyen sang item khac

                    //Neu chua  ton tai thi thuc hien lay du lieu va add vao trong list TSB
                    long qtySum = listChildData.Where(p => p.model.Substring(0, 9) == tempModel).Sum(p => p.qty);
                    string tempCus = item.cusDetail.Substring(0, item.cusDetail.IndexOf(")", 2) + 1);
                    listTSB.Add(new DataTSB(tempModel, tempCus, qtySum));
                }

                foreach (var item in listChildErr)//Duyet cac item ton tai trong trong ma loi
                {
                    string tempModel = item.model.Substring(0, 9);
                    bool check = false;
                    foreach (var itemTSB in listTSB)
                    {
                        if (tempModel.Equals(itemTSB.item))
                        {
                            switch (item.nameError)
                            {
                                case string s when s.Equals(MdlComment.TYPE_ERROR_BAC_CAU):
                                    itemTSB.qty4BrightMake += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_BONG_VO_LK):
                                    itemTSB.qty12Peel += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_DI_VAT):
                                    itemTSB.qty10OjectForeign += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_GIA_HAN):
                                    itemTSB.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_KENH_NGHIENG):
                                    itemTSB.qty3Warp += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_KHONG_HAN):
                                    itemTSB.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_NGUOC_HUONG):
                                    itemTSB.qty8Reverse += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_THIEU_THIEC):
                                    itemTSB.qty5TinSmall += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_THUA_THIEU_LK):
                                    if (item.typeThua == true)
                                    {
                                        itemTSB.qty11ItemMiss += item.qty;
                                    }
                                    else
                                    {
                                        itemTSB.qty6ItemLack += item.qty;
                                    }
                                    break;

                                default:
                                    itemTSB.qty13Other += item.qty;
                                    break;
                            }
                            check = true;//Neu ton tai thi dung lai chuyen sang giai doan khac va check = true;
                            break;
                        }
                    }
                    if (check == false)
                    {
                        return string.Format(RESULT.ERROR_FILE_ERROR_MODEL, tempModel);
                    }

                }

                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "GetTSB", ex.Message);
            }
        }
        /// <summary>
        /// Thuc hien lay du lieu cua Kyocera
        /// </summary>
        /// <param name="listData"></param>
        /// <param name="listError"></param>
        /// <param name="listKyocera"></param>
        /// <returns></returns>
        /// CreatedBy: HoaiPT(06/03/2023)
        public static string GetKyocera(List<DataFirst> listData, List<DataError> listError, ref List<DataKyocera> listKyocera)
        {
            try
            {
                var listChildData = listData.Where(x => x.cusCode.Equals("KYOCERA")).ToList();
                var listChildErr = listError.Where(x => x.cusCode.Equals("KYOCERA")).ToList();

                //Phan lay du lieu cua model xong roi
                foreach (var item in listChildData.ToArray())
                {
                    string tempModel = item.model;
                    var check = listKyocera.FirstOrDefault(p => p.item.Equals(tempModel));
                    if (check != null)
                    {
                        continue;
                    }

                    long qtySum = listChildData.Where(p => p.model == tempModel).Sum(p => p.qty);

                    string tempCus = item.cusDetail.Substring(0, item.cusDetail.IndexOf("-"));//Phan nay cat khac so voi cac loai cat khac nhau
                    listKyocera.Add(new DataKyocera(tempModel, tempCus, qtySum));
                }

                foreach (var item in listChildErr)
                {
                    string tempModel = item.model;
                    bool check = false;
                    foreach (var itemKyocera in listKyocera)
                    {
                        if (tempModel.Equals(itemKyocera.item))
                        {
                            switch (item.nameError)
                            {
                                case string s when s.Equals(MdlComment.TYPE_ERROR_BAC_CAU):
                                    itemKyocera.qty4BrightMake += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_BONG_VO_LK):
                                    itemKyocera.qty12Peel += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_DI_VAT):
                                    itemKyocera.qty10OjectForeign += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_GIA_HAN):
                                    itemKyocera.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_KENH_NGHIENG):
                                    itemKyocera.qty3Warp += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_KHONG_HAN):
                                    itemKyocera.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_NGUOC_HUONG):
                                    itemKyocera.qty8Reverse += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_THIEU_THIEC):
                                    itemKyocera.qty5TinSmall += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_THUA_THIEU_LK):
                                    if (item.typeThua == true)
                                    {
                                        itemKyocera.qty11ItemMiss += item.qty;
                                    }
                                    else
                                    {
                                        itemKyocera.qty6ItemLack += item.qty;
                                    }

                                    break;
                                default:
                                    itemKyocera.qty13Other += item.qty;
                                    break;
                            }

                            check = true;
                            break;
                        }
                    }
                    if (check == false)
                    {
                        return string.Format(RESULT.ERROR_FILE_ERROR_MODEL, tempModel);
                    }

                }

                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "GetKyocera", ex.Message);
            }
        }
        public static string GetFX(List<DataFirst> listData, List<DataError> listError, ref List<DataFX> listFX)
        {
            try
            {
                var listChildData = listData.Where(x => x.cusCode.Equals("FX")).ToList();
                var listChildErr = listError.Where(x => x.cusCode.Equals("FX")).ToList();

                //Phan lay du lieu cua model xong roi
                foreach (var item in listChildData.ToArray())
                {
                    string tempModel = item.model.Substring(0, 9);
                    var check = listFX.FirstOrDefault(p => p.item.Substring(0, 9).Equals(tempModel));
                    if (check != null)
                    {
                        continue;
                    }

                    long qtySum = listChildData.Where(p => p.model.Substring(0, 9) == tempModel).Sum(p => p.qty);
                    string tempCus = item.cusDetail.Substring(0, item.cusDetail.IndexOf(")", 2) + 1);
                    listFX.Add(new DataFX(tempModel, tempCus, qtySum));
                }

                foreach (var item in listChildErr)
                {
                    string tempModel = item.model.Substring(0, 9);
                    bool check = false;
                    foreach (var itemFX in listFX)
                    {
                        if (tempModel.Equals(itemFX.item))
                        {
                            switch (item.nameError)
                            {
                                case string s when s.Equals(MdlComment.TYPE_ERROR_BAC_CAU):
                                    itemFX.qty4BrightMake += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_BONG_VO_LK):
                                    itemFX.qty12Peel += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_DI_VAT):
                                    itemFX.qty10OjectForeign += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_GIA_HAN):
                                    itemFX.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_KENH_NGHIENG):
                                    itemFX.qty3Warp += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_KHONG_HAN):
                                    itemFX.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_NGUOC_HUONG):
                                    itemFX.qty8Reverse += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_THIEU_THIEC):
                                    itemFX.qty5TinSmall += item.qty;
                                    break;
                                case string s when s.Equals(MdlComment.TYPE_ERROR_THUA_THIEU_LK):
                                    if (item.typeThua == true)
                                    {
                                        itemFX.qty11ItemMiss += item.qty;
                                    }
                                    else
                                    {
                                        itemFX.qty13Other += item.qty;
                                    }
                                    break;
                                default:
                                    itemFX.qty13Other += item.qty;
                                    break;
                            }

                            check = true;
                            break;
                        }
                    }
                    if (check == false)
                    {
                        return string.Format(RESULT.ERROR_FILE_ERROR_MODEL, tempModel);
                    }

                }

                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "FX", ex.Message);
            }
        }

        /// <summary>
        /// Thuc hien lay du lieu cua Hitachi
        /// </summary>
        /// <param name="listData"></param>
        /// <param name="listError"></param>
        /// <param name="valueHT"></param>
        /// <returns></returns>
        public static string GetHT(List<DataFirst> listData, List<DataError> listError, ref DataHT valueHT)
        {
            try
            {
                valueHT.qtySum = listData.Where(x => x.cusCode.Equals("HITACHI")).Sum(q => q.qty);
                var listChildErr = listError.Where(x => x.cusCode.Equals("HITACHI")).ToList();

                foreach (var item in listChildErr)
                {
                    switch (item.nameError)
                    {
                        case string s when s.Equals(MdlComment.TYPE_ERROR_BAC_CAU):
                            valueHT.qty4BrightMake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_GIA_HAN):
                            valueHT.qty1WeldFake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_KENH_NGHIENG):
                            valueHT.qty3Warp += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_KHONG_HAN):
                            valueHT.qty1WeldFake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_NGUOC_HUONG):
                            valueHT.qty8Reverse += item.qty;
                            break;

                        case string s when s.Equals(MdlComment.TYPE_ERROR_THUA_THIEU_LK):
                            if (item.typeThua == true)
                            {
                                valueHT.qty11ItemMiss += item.qty;
                            }
                            else
                            {
                                valueHT.qty6ItemLack += item.qty;
                            }
                            break;

                        default:
                            valueHT.qty13Other += item.qty;
                            break;
                    }
                }
                return RESULT.OK;

            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "GetHT", ex.Message);
            }
        }
        /// <summary>
        /// Thuc  hien lay du lieu cua Okidenki
        /// </summary>
        /// <param name="listData"></param>
        /// <param name="listError"></param>
        /// <param name="valueOkidenki"></param>
        /// <returns></returns>
        public static string GetOkidenki(List<DataFirst> listData, List<DataError> listError, ref DataOkidenki valueOkidenki)
        {
            try
            {
                valueOkidenki.qtySum = listData.Where(x => x.cusCode.Equals("OKIDENKI")).Sum(q => q.qty);
                var listChildErr = listError.Where(x => x.cusCode.Equals("OKIDENKI")).ToList();

                foreach (var item in listChildErr)
                {
                    switch (item.nameError)
                    {
                        case string s when s.Equals(MdlComment.TYPE_ERROR_BAC_CAU):
                            valueOkidenki.qty4BrightMake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_BONG_VO_LK):
                            valueOkidenki.qty12Peel += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_DI_VAT):
                            valueOkidenki.qty10OjectForeign += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_GIA_HAN):
                            valueOkidenki.qty1WeldFake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_KENH_NGHIENG):
                            valueOkidenki.qty3Warp += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_KHONG_HAN):
                            valueOkidenki.qty1WeldFake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_NGUOC_HUONG):
                            valueOkidenki.qty8Reverse += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_THIEU_THIEC):
                            valueOkidenki.qty5TinSmall += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_THUA_THIEU_LK):
                            if (item.typeThua == true)
                            {
                                valueOkidenki.qty11ItemMiss += item.qty;
                            }
                            else
                            {
                                valueOkidenki.qty6ItemLack += item.qty;
                            }
                            break;

                        default:
                            valueOkidenki.qty13Other += item.qty;
                            break;
                    }
                }
                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "GetOkidenki", ex.Message);
            }
            finally
            {

            }
        }
        /// <summary>
        /// Ghi du lieu cua Riso
        /// </summary>
        /// <param name="listData"></param>
        /// <param name="listError"></param>
        /// <param name="valueRiso"></param>
        /// <returns></returns>
        public static string GetRISO(List<DataFirst> listData, List<DataError> listError, ref DataRiso valueRiso)
        {
            try
            {
                valueRiso.qtySum = listData.Where(x => x.cusCode.Equals("RISO")).Sum(q => q.qty);
                var listChildErr = listError.Where(x => x.cusCode.Equals("RISO")).ToList();

                foreach (var item in listChildErr)
                {
                    switch (item.nameError)
                    {
                        case string s when s.Equals(MdlComment.TYPE_ERROR_BAC_CAU):
                            valueRiso.qty4BrightMake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_BONG_VO_LK):
                            valueRiso.qty12Peel += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_DI_VAT):
                            valueRiso.qty10OjectForeign += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_GIA_HAN):
                            valueRiso.qty1WeldFake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_KENH_NGHIENG):
                            valueRiso.qty3Warp += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_KHONG_HAN):
                            valueRiso.qty1WeldFake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_NGUOC_HUONG):
                            valueRiso.qty8Reverse += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_THIEU_THIEC):
                            valueRiso.qty5TinSmall += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_THUA_THIEU_LK):
                            if (item.typeThua == true)
                            {
                                valueRiso.qty11ItemMiss += item.qty;
                            }
                            else
                            {
                                valueRiso.qty6ItemLack += item.qty;
                            }
                            break;

                        default:
                            valueRiso.qty13Other += item.qty;
                            break;
                    }
                }
                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "GetRiso", ex.Message);
            }
            finally
            {

            }
        }
        /// <summary>
        /// Lay du lieu cua JCM
        /// </summary>
        /// <param name="listData"></param>
        /// <param name="listError"></param>
        /// <param name="valueJCM"></param>
        /// <returns></returns>
        /// CreatedBy: HoaiPT(06/03/2023)
        public static string GetJCM(List<DataFirst> listData, List<DataError> listError, ref DataJCM valueJCM)
        {
            try
            {
                valueJCM.qtySum = listData.Where(x => x.cusCode.Equals("JCM")).Sum(q => q.qty);
                var listChildErr = listError.Where(x => x.cusCode.Equals("JCM")).ToList();

                foreach (var item in listChildErr)
                {
                    switch (item.nameError)
                    {
                        case string s when s.Equals(MdlComment.TYPE_ERROR_BAC_CAU):
                            valueJCM.qty4BrightMake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_BONG_VO_LK):
                            valueJCM.qty12Peel += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_DI_VAT):
                            valueJCM.qty10OjectForeign += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_GIA_HAN):
                            valueJCM.qty1WeldFake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_KHONG_HAN):
                            valueJCM.qty1WeldFake += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_NGUOC_HUONG):
                            valueJCM.qty8Reverse += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_THIEU_THIEC):
                            valueJCM.qty5TinSmall += item.qty;
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_THUA_THIEU_LK):
                            if (item.typeThua == true)
                            {
                                valueJCM.qty11ItemMiss += item.qty;
                            }
                            else
                            {
                                valueJCM.qty6ItemLack += item.qty;
                            }
                            break;
                        case string s when s.Equals(MdlComment.TYPE_ERROR_LECH_LINH_KIEN)://Phan nay khac voai cac phan khac
                            valueJCM.qty14LechLK += item.qty;
                            break;

                        default:
                            valueJCM.qty13Other += item.qty;
                            break;
                    }
                }
                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "GetJCM", ex.Message);
            }
        }
    }
}
