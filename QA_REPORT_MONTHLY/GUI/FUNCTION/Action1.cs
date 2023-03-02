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


        public static string OpenFileExcelData(ActionInput1 valueInput, string sheetName, ref List<DataFirst> listData)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(valueInput.fileData, ReadOnly: true);

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

                int rowCurrent = 4;
                string check = ws.Cells[rowCurrent, "D"].Value;
                DataFirst temp = new DataFirst();
                while (!string.IsNullOrWhiteSpace(check))
                {
                    temp.wo = check;
                    temp.model = ws.Cells[rowCurrent, "B"].Value;
                    try
                    {
                        temp.cusDetail = ws.Cells[rowCurrent, "C"].Value;
                    }
                    catch (Exception exs)
                    {
                        return string.Format(RESULT.ERROR_COLUMN_C, rowCurrent, exs.Message);
                    }

                    temp.qty = Convert.ToInt64(ws.Cells[rowCurrent, "F"].Value);

                    try
                    {
                        temp.cusCode = ws.Cells[rowCurrent, "G"].Value;
                    }
                    catch (Exception exs)
                    {
                        return string.Format(RESULT.ERROR_COLUMN_G, rowCurrent, exs.Message);
                    }

                    listData.Add(new DataFirst(temp));

                    rowCurrent++;
                    check = ws.Cells[rowCurrent, "D"].Value;
                }


                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, ex.Message);
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


        public static string OpenFileExcelError(ActionInput1 valueInput, ref List<DataError> listError)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(valueInput.fileError);
                ws = wb.Sheets[DataConfig.CONFIG_FILE_ERROR_SHEETNAME];

                //Thuc hien lay dong cuoi cung co du lieu
                if (ws.AutoFilter != null)
                {
                    ws.Unprotect(DataConfig.CONFIG_FILE_ERROR_PASSWORD);
                    ws.AutoFilterMode = false;
                }
                int lastRow = ws.Cells[ws.Rows.Count, "P"].End(Excel.XlDirection.xlUp).Row;

                DataError err = new DataError();
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
                    err.nameError= Convert.ToString(ws.Cells[i, "T"].value);
                    if (string.IsNullOrWhiteSpace(err.nameError))
                    {
                        continue;
                    }
                    

                    //Lay so luong loi
                    int tempQty;
                    if (!int.TryParse(Convert.ToString(ws.Cells[i, "Z"].value), out tempQty))
                    {
                        continue;
                    }
                    if (tempQty <= 0)
                    {
                        continue;
                    }
                    err.qty = tempQty;

                    err.wo = Convert.ToString(ws.Cells[i, "Q"].value);
                    if (string.IsNullOrWhiteSpace(err.wo))
                    {
                        continue;
                    }

                    if (err.nameError.Equals("Thừa, thiếu LK"))
                    {
                        Comment comment = ws.Cells[i, "T"].Comment;
                        if(comment == null)
                        {
                            return string.Format(RESULT.ERROR_FILE_ERROR_NOT_COMMENT, i);
                        }
                        else
                        {
                            if(!(comment.Text().Contains("thiếu") || comment.Text().Contains("thừa")))
                            {
                                return string.Format(RESULT.ERROR_FILE_ERROR_COMMENT_NOT_RULE, i, comment.Text());
                            }
                            err.noteNameError = comment.Text();
                        }

                    }

                    listError.Add(new DataError(err));



                }


                return RESULT.OK;
            }
            //catch (Exception ex)
            //{
            //    return string.Format(RESULT.ERROR_015_CATCH, "OpenFileExcelError", ex.Message);
            //}
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
                    if(indexCheck == -1)
                    {
                        return string.Format(RESULT.ERROR_FILE_ERROR_WO, itemErr.wo, itemErr.ToString());
                    }

                    itemErr.cusCode = listData[indexCheck].cusCode;
                }
                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "ActionFileError", ex.Message);
            }
        }

        public static string GetTSB(List<DataFirst> listData, List<DataError> listError, ref List<DataTSB> listTSB)
        {
            try
            {
                var listChildData = listData.Where(x => x.cusCode.Equals("TSB")).ToList();
                var listChildErr = listError.Where(x => x.cusCode.Equals("TSB")).ToList();

                //Phan lay du lieu cua model xong roi
                foreach (var item in listChildData.ToArray())
                {
                    string tempModel = item.model.Substring(0, 9);
                    var check = listTSB.FirstOrDefault(p => p.item.Equals(tempModel));
                    if(check != null)
                    {
                        continue;
                    }

                    long qtySum = listChildData.Where(p => p.model.Substring(0, 9) == tempModel).Sum(p=>p.qty);

                    string tempCus = item.cusDetail.Substring(0, item.cusDetail.IndexOf(")", 2) + 1);
                    listTSB.Add(new DataTSB(tempModel, tempCus, qtySum));
                }

                foreach (var item in listChildErr)
                {
                    string tempModel = item.model.Substring(0, 9);
                    bool check = false;
                    foreach (var itemTSB in listTSB)
                    {
                        if (tempModel.Equals(itemTSB.item))
                        {
                            switch (item.nameError)
                            {
                                case string s when s.Equals("Bắc cầu"):
                                    itemTSB.qty4BrightMake += item.qty;
                                    break;
                                case string s when s.Equals("Bong, vỡ LK"):
                                    itemTSB.qty12Peel += item.qty;
                                    break;
                                case string s when s.Equals("Dị vật"):
                                    itemTSB.qty10OjectForeign += item.qty;
                                    break;
                                case string s when s.Equals("Giả hàn"):
                                    itemTSB.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals("Kênh, Nghiêng"):
                                    itemTSB.qty3Warp += item.qty;
                                    break;
                                case string s when s.Equals("Không hàn"):
                                    itemTSB.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals("Ngược hướng"):
                                    itemTSB.qty8Reverse += item.qty;
                                    break;
                                case string s when s.Equals("Thiếu thiếc"):
                                    itemTSB.qty5TinSmall += item.qty;
                                    break;
                                case string s when s.Equals("Thừa, thiếu LK"):
                                    itemTSB.qty4BrightMake += item.qty;


                                    //cai nay lam sau
                                    break;
                                default:
                                    itemTSB.qty13Other += item.qty;
                                    break;
                            }



                            check = true;
                            break; 
                        }
                    }
                    if(check == false)
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

                    string tempCus = item.cusDetail.Substring(0, item.cusDetail.IndexOf("-"));
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
                                case string s when s.Equals("Bắc cầu"):
                                    itemKyocera.qty4BrightMake += item.qty;
                                    break;
                                case string s when s.Equals("Bong, vỡ LK"):
                                    itemKyocera.qty12Peel += item.qty;
                                    break;
                                case string s when s.Equals("Dị vật"):
                                    itemKyocera.qty10OjectForeign += item.qty;
                                    break;
                                case string s when s.Equals("Giả hàn"):
                                    itemKyocera.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals("Kênh, Nghiêng"):
                                    itemKyocera.qty3Warp += item.qty;
                                    break;
                                case string s when s.Equals("Không hàn"):
                                    itemKyocera.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals("Ngược hướng"):
                                    itemKyocera.qty8Reverse += item.qty;
                                    break;
                                case string s when s.Equals("Thiếu thiếc"):
                                    itemKyocera.qty5TinSmall += item.qty;
                                    break;
                                case string s when s.Equals("Thừa, thiếu LK"):
                                    itemKyocera.qty4BrightMake += item.qty;


                                    //cai nay lam sau
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
                                case string s when s.Equals("Bắc cầu"):
                                    itemFX.qty4BrightMake += item.qty;
                                    break;
                                case string s when s.Equals("Bong, vỡ LK"):
                                    itemFX.qty12Peel += item.qty;
                                    break;
                                case string s when s.Equals("Dị vật"):
                                    itemFX.qty10OjectForeign += item.qty;
                                    break;
                                case string s when s.Equals("Giả hàn"):
                                    itemFX.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals("Kênh, Nghiêng"):
                                    itemFX.qty3Warp += item.qty;
                                    break;
                                case string s when s.Equals("Không hàn"):
                                    itemFX.qty1WeldFake += item.qty;
                                    break;
                                case string s when s.Equals("Ngược hướng"):
                                    itemFX.qty8Reverse += item.qty;
                                    break;
                                case string s when s.Equals("Thiếu thiếc"):
                                    itemFX.qty5TinSmall += item.qty;
                                    break;
                                case string s when s.Equals("Thừa, thiếu LK"):
                                    itemFX.qty4BrightMake += item.qty;

                                    //cai nay lam sau
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
                return string.Format(RESULT.ERROR_015_CATCH, "GetKyocera", ex.Message);
            }
        }
    }
}
