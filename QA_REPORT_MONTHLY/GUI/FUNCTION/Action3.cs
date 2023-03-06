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
    public class Action3
    {

        public static string ValidateInput3(ref ActionInput3 valueInput)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(valueInput.colModel) ||
                    string.IsNullOrWhiteSpace(valueInput.fileData) ||
                    string.IsNullOrWhiteSpace(valueInput.rowEndString) ||
                    string.IsNullOrWhiteSpace(valueInput.rowStartString) ||
                    string.IsNullOrWhiteSpace(valueInput.sheetName) ||
                    string.IsNullOrWhiteSpace(valueInput.fileError) ||
                    string.IsNullOrWhiteSpace(valueInput.fileInput)||
                    string.IsNullOrWhiteSpace(valueInput.monthString)||
                    string.IsNullOrWhiteSpace(valueInput.colWrite))
                {
                    return RESULT.ERROR_NOT_NULL;
                }
                valueInput.Trim();//Thuc hien loai bo ki tu thua

                int tempNumber;
                if (!int.TryParse(valueInput.rowStartString, out tempNumber))
                {
                    return string.Format(RESULT.ERROR_2_INPUT_NOT_NUMBER, "Dòng bắt đầu");
                }
                valueInput.rowStart = tempNumber;

                if (!int.TryParse(valueInput.rowEndString, out tempNumber))
                {
                    return string.Format(RESULT.ERROR_2_INPUT_NOT_NUMBER, "Dòng kết thúc");
                }
                valueInput.rowEnd = tempNumber;
                if (!(valueInput.rowEnd > valueInput.rowStart))
                {
                    return RESULT.ERROR_2_INPUT_NUMBER_RULE;
                }

                int tempMonth;
                if (!int.TryParse(valueInput.monthString, out tempMonth))
                {
                    return RESULT.ERROR_NOT_NUMBER;
                }

                if (!(tempMonth >= 1 && tempMonth <= 12))
                {
                    return RESULT.ERROR_NOT_MONTH;
                }

                valueInput.monthInt = tempMonth;
                valueInput.monthString = String.Format("{0:00}", tempMonth);




                if (!File.Exists(valueInput.fileData))
                {
                    return string.Format(RESULT.ERROR_NOT_FILE, valueInput.fileData);
                }
                if (!File.Exists(valueInput.fileError))
                {
                    return string.Format(RESULT.ERROR_NOT_FILE, valueInput.fileError);
                }
                if (!File.Exists(valueInput.fileInput))
                {
                    return string.Format(RESULT.ERROR_NOT_FILE, valueInput.fileInput);
                }


                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "ValidateInput3", ex.Message);
            }
        }

        public static string CheckSheetName(ActionInput3 vauleInput)
        {

            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(vauleInput.fileInput);

                bool existSheetName = false;
                foreach (Excel._Worksheet sheet in wb.Worksheets)
                {
                    if (sheet.Name.Equals(vauleInput.sheetName))
                    {
                        existSheetName = true;
                        break;
                    }
                }
                if (existSheetName == false)
                {
                    return string.Format(RESULT.ERROR_SHEETNAME, vauleInput.sheetName + "(File ghi dữ liệu)");
                }

                wb.Close(false);
                app.Quit();
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

        public static string OpenFileExcelData3(ActionInput3 valueInput, string sheetName, ref List<DataFirst> listData)
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

                wb.Close(false);
                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "OpenFileExcelData 3", ex.Message);
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

        public static string OpenFileExcelError(ActionInput3 valueInput, ref List<DataError> listError)
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
                    err.nameError = Convert.ToString(ws.Cells[i, "T"].value);
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
                return string.Format(RESULT.ERROR_015_CATCH, "ActionFileError", ex.Message);
            }
        }

        public static string GetTSB_3(List<DataFirst> listData, List<DataError> listError, ref List<DataTSB3> listTSB)
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
                    if (check != null)
                    {
                        continue;
                    }

                    long qtySum = listChildData.Where(p => p.model.Substring(0, 9).Equals(tempModel)).Sum(p => p.qty);
                    int qtyErrorSum = listChildErr.Where(p => p.model.Substring(0, 9).Equals(tempModel)).Sum(p => p.qty);
                   
                    listTSB.Add(new DataTSB3(tempModel, qtySum,qtyErrorSum));
                }

                
                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "GetTSB", ex.Message);
            }
        }

        public static void CheckItemMiss(List<DataTSB3> listTSB, ref string valueResult)
        {
            foreach (var item in listTSB)
            {
                if(item.action == false)
                {
                    valueResult += item.ToString() +"\n";
                }
            }
        }
    }
}
