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


        public static string OpenFileExcel(ActionInput1 valueInput, string sheetName, ref List<DataFirst> listData)
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
    }
}
