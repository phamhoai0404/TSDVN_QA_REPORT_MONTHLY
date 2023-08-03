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
    public class Action2
    {
        public static string ValidateInputAction2(ref ActionInput2 valueInput)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(valueInput.colModel) ||
                    string.IsNullOrWhiteSpace(valueInput.fileData) ||
                    string.IsNullOrWhiteSpace(valueInput.rowEndString) ||
                    string.IsNullOrWhiteSpace(valueInput.rowStartString) ||
                    string.IsNullOrWhiteSpace(valueInput.sheetName))
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

                if (!File.Exists(valueInput.fileData))
                {
                    return string.Format(RESULT.ERROR_NOT_FILE, valueInput.fileData);
                }
                if (!(valueInput.rowEnd > valueInput.rowStart))
                {
                    return RESULT.ERROR_2_INPUT_NUMBER_RULE;
                }

                return RESULT.OK;
            }
            catch (Exception ex)
            {

                return string.Format(RESULT.ERROR_015_CATCH, "ValidateInputAction2", ex.Message);
            }
        }

        /// <summary>
        /// Thuc hien lay du lieu cu cua file excel
        /// </summary>
        /// <param name="info"></param>
        /// <param name="listKyocera"></param>
        /// <returns></returns>
        /// CreatedBy: HoaiPT(06/03/2023)
        public static string GetKyoceraOld(ActionInput2 info, ref List<DataKyocera> listKyocera)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(info.fileData, ReadOnly: true);

                bool existSheetName = false;
                foreach (Excel._Worksheet sheet in wb.Worksheets)
                {
                    if (sheet.Name.Equals(info.sheetName))
                    {
                        existSheetName = true;
                        break;
                    }
                }
                if (existSheetName == false)
                {
                    return string.Format(RESULT.ERROR_SHEETNAME, info.sheetName);
                }
                ws = wb.Sheets[info.sheetName];

                int columnIndex = ws.Range[info.colModel + "1"].Column;
                int colCurrent;
                DataKyocera valueTemp;
                for (int i = info.rowStart; i <= info.rowEnd; i++)
                {
                    colCurrent = columnIndex;
                    valueTemp = new DataKyocera();

                    string temp = ws.Cells[i, colCurrent++].value;
                    if (string.IsNullOrWhiteSpace(temp))//Khong duoc de trong dong model
                    {
                        return string.Format(RESULT.ERROR_2_NOT_NULL_MODEL, i);
                    }
                    temp = temp.Trim();//Loại bỏ đi kí tự thừa
                    int indexInt = temp.IndexOf("(");
                    if (indexInt == -1)//Neu = -1 khong ton tai ki tu nay
                    {
                        return string.Format(RESULT.ERROR_2_NOT_OPEN, i);
                    }
                    valueTemp.cus = temp;
                    valueTemp.item = temp.Substring(0, indexInt).Trim();

                    long indexLong;
                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);//Thuc hien column tiếp theo là Sum
                    if (!long.TryParse(temp, out indexLong))
                    {
                        return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Số bản mạch sản xuất", i);
                    }
                    valueTemp.qtySum = indexLong;

                    colCurrent += 3;//3 column khong can lay du lieu

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Hàn giả", i);
                        }
                        valueTemp.qty1WeldFake = indexInt;
                    }

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Sai vị  trí", i);
                        }
                        valueTemp.qty2ErrorPosition = indexInt;
                    }

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Kênh", i);
                        }
                        valueTemp.qty3Warp = indexInt;
                    }

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Bắc cầu", i);
                        }
                        valueTemp.qty4BrightMake = indexInt;
                    }

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Ít thiếc", i);
                        }
                        valueTemp.qty5TinSmall = indexInt;
                    }

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Thiếu linh kiện", i);
                        }
                        valueTemp.qty6ItemLack = indexInt;
                    }

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Lật ngược", i);
                        }
                        valueTemp.qty7ErrorPosition = indexInt;
                    }

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Ngược hướng", i);
                        }
                        valueTemp.qty8Reverse = indexInt;
                    }

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Nhầm linh kiện", i);
                        }
                        valueTemp.qty9DirectionRev = indexInt;
                    }

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Dị vật", i);
                        }
                        valueTemp.qty10OjectForeign = indexInt;
                    }

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Thừa linh kiện", i);
                        }
                        valueTemp.qty11ItemMiss = indexInt;
                    }

                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Bong", i);
                        }
                        valueTemp.qty12Peel = indexInt;
                    }
                    temp = Convert.ToString(ws.Cells[i, colCurrent++].Value);
                    if (!string.IsNullOrEmpty(temp))
                    {
                        if (!int.TryParse(temp, out indexInt))
                        {
                            return string.Format(RESULT.ERROR_2_NOT_NUMBER, "Khác", i);
                        }
                        valueTemp.qty13Other = indexInt;
                    }

                    listKyocera.Add(new DataKyocera(valueTemp));
                }


                wb.Close(false);
                app.Quit();
                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "GetKyoceraOld", ex.Message);
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
        /// Thuc hien xu ly du lieu
        /// </summary>
        /// <param name="listKyocrea"></param>
        /// <returns></returns>
        /// CreatedBy: HoaiPT(06/03/2023)
        public static string ExecuteKyocera(ref List<DataKyocera> listKyocrea)
        {
            try
            {
                for (int i = 0; i < listKyocrea.Count()- 1; i++)
                {
                    for(int j = i +1; j < listKyocrea.Count(); j++)
                    {
                        if (listKyocrea[i].item.Equals(listKyocrea[j].item))
                        {
                            listKyocrea[i].Pus(listKyocrea[j]);//Cong them so luong
                            listKyocrea.RemoveAt(j);//Xoa bo phan tu j
                            j--;//lui lai 
                        }
                    }
                }

                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "ExecuteKyocera", ex.Message);
            }
        }

        internal static string WriteKyocera_2(List<DataKyocera> listKyocrea)
        {
            throw new NotImplementedException();
        }
    }
}
