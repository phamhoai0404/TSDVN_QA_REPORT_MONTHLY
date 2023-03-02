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
    public class ActionWrite
    {
        public static string CreateFile()
        {
            string currentPath = Directory.GetCurrentDirectory();
            string pathTemplate = currentPath + DataConfig.CONFIG_FILE_TEMPLATE;
            DataConfig.CONFIG_FILE_RESULT = DateTime.Now.ToString("ddMMyyyy_HHmmss");
            DataConfig.CONFIG_FILE_RESULT = currentPath + @"\RESULT\" + DataConfig.CONFIG_FILE_RESULT + ".xlsx";

            Excel.Application excel = new Excel.Application();
            Excel.Workbook originalWorkbook = excel.Workbooks.Open(pathTemplate);

            originalWorkbook.SaveCopyAs(DataConfig.CONFIG_FILE_RESULT);
            originalWorkbook.Close(false);
            excel.Quit();

            return RESULT.OK;
        }

        public static string WriteTSB1(List<DataTSB> listTSB)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;

            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(DataConfig.CONFIG_FILE_RESULT);

                ws = wb.Sheets["TOSIBA_1"];

                int rowCurrent = 3;
                foreach (var item in listTSB)
                {
                    ws.Cells[rowCurrent, "A"].value = item.item + item.cus;
                    ws.Cells[rowCurrent, "C"].value = item.qtySum;
                    ws.Cells[rowCurrent, "H"].value = 80;

                    if (item.qty1WeldFake != 0)
                    {
                        ws.Cells[rowCurrent, 9].value = item.qty1WeldFake;
                    }

                    if (item.qty2ErrorPosition != 0)
                    {
                        ws.Cells[rowCurrent, 10].value = item.qty2ErrorPosition;
                    }

                    if (item.qty3Warp != 0)
                    {
                        ws.Cells[rowCurrent, 11].value = item.qty3Warp;
                    }

                    if (item.qty4BrightMake != 0)
                    {
                        ws.Cells[rowCurrent, 12].value = item.qty4BrightMake;
                    }

                    if (item.qty5TinSmall != 0)
                    {
                        ws.Cells[rowCurrent, 13].value = item.qty5TinSmall;
                    }

                    if (item.qty6ItemLack != 0)
                    {
                        ws.Cells[rowCurrent, 14].value = item.qty6ItemLack;
                    }

                    if (item.qty7ErrorPosition != 0)
                    {
                        ws.Cells[rowCurrent, 15].value = item.qty7ErrorPosition;
                    }

                    if (item.qty8Reverse != 0)
                    {
                        ws.Cells[rowCurrent, 16].value = item.qty8Reverse;
                    }

                    if (item.qty9DirectionRev != 0)
                    {
                        ws.Cells[rowCurrent, 17].value = item.qty9DirectionRev;
                    }

                    if (item.qty10OjectForeign != 0)
                    {
                        ws.Cells[rowCurrent, 18].value = item.qty10OjectForeign;
                    }

                    if (item.qty11ItemMiss != 0)
                    {
                        ws.Cells[rowCurrent, 19].value = item.qty11ItemMiss;
                    }
                    if (item.qty12Peel != 0)
                    {
                        ws.Cells[rowCurrent, 20].value = item.qty12Peel;
                    }
                    if (item.qty13Other != 0)
                    {
                        ws.Cells[rowCurrent, 21].value = item.qty13Other;
                    }


                    rowCurrent++;

                }

                range = ws.Range["A3:W" + (rowCurrent)];
                range.Borders.LineStyle = XlLineStyle.xlContinuous;
                range.Borders.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
                range.Borders.Weight = XlBorderWeight.xlThin;


                rowCurrent = 3;
                for (int i = 0; i <= listTSB.Count; i++)
                {
                    ws.Range["A" + rowCurrent + ":B" + rowCurrent].Merge();

                    ws.Range["C" + rowCurrent + ":D" + rowCurrent].Merge();
                    ws.Range["F" + rowCurrent + ":G" + rowCurrent].Merge();
                    
                    ws.Range["V" + rowCurrent + ":W" + rowCurrent].Merge();

                    ws.Cells[rowCurrent, "F"].Formula = string.Format("=E{0}/C{0}*1000000",rowCurrent);
                    ws.Cells[rowCurrent, "E"].Formula = string.Format("=SUM(I{0}:U{0}", rowCurrent);
                    rowCurrent++;
                }
                rowCurrent--;
                ws.Range["A3:V" + (rowCurrent)].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws.Cells[rowCurrent, "A"].Value = @"合計";
                ws.Cells[rowCurrent, "C"].Formula = @"= SUM(C3:C" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "E"].Formula = @"= SUM(E3:E" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "F"].Formula = @"= SUM(F3:F" + (rowCurrent - 1) + ")";

                ws.Cells[rowCurrent, "H"].Formula = @"= SUM(H3:H" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "I"].Formula = @"= SUM(I3:I" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "J"].Formula = @"= SUM(J3:J" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "K"].Formula = @"= SUM(K3:K" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "L"].Formula = @"= SUM(L3:L" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "M"].Formula = @"= SUM(M3:M" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "N"].Formula = @"= SUM(N3:N" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "O"].Formula = @"= SUM(O3:O" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "P"].Formula = @"= SUM(P3:P" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "Q"].Formula = @"= SUM(Q3:Q" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "R"].Formula = @"= SUM(R3:R" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "S"].Formula = @"= SUM(S3:S" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "T"].Formula = @"= SUM(T3:T" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "U"].Formula = @"= SUM(U3:U" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "F"].Formula = string.Format("=E{0}/C{0}*1000000", rowCurrent);




                wb.Save();

                wb.Close();
                app.Quit();

                return RESULT.OK;
            }
            //catch (Exception ex)
            //{
            //    return string.Format(RESULT.ERROR_015_CATCH, "WriteTSB1", ex.Message);
            //}
            finally
            {
                if (range != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                }
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
        public static string WriteKyocera1(List<DataKyocera> listKyocera)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;

            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(DataConfig.CONFIG_FILE_RESULT);

                ws = wb.Sheets["KYOCERA_1"];

                int rowCurrent = 3;
                foreach (var item in listKyocera)
                {
                    ws.Cells[rowCurrent, "A"].value = item.item + "("+ item.cus + ")";
                    ws.Cells[rowCurrent, "B"].value = item.qtySum;
                    ws.Cells[rowCurrent, "E"].value = 80;

                   // int columnStart = 6;
                    if (item.qty1WeldFake != 0)
                    {
                        ws.Cells[rowCurrent, 6].value = item.qty1WeldFake;
                    }

                    if (item.qty2ErrorPosition != 0)
                    {
                        ws.Cells[rowCurrent, 7].value = item.qty2ErrorPosition;
                    }

                    if (item.qty3Warp != 0)
                    {
                        ws.Cells[rowCurrent, 8].value = item.qty3Warp;
                    }

                    if (item.qty4BrightMake != 0)
                    {
                        ws.Cells[rowCurrent, 9].value = item.qty4BrightMake;
                    }

                    if (item.qty5TinSmall != 0)
                    {
                        ws.Cells[rowCurrent, 10].value = item.qty5TinSmall;
                    }

                    if (item.qty6ItemLack != 0)
                    {
                        ws.Cells[rowCurrent, 11].value = item.qty6ItemLack;
                    }

                    if (item.qty7ErrorPosition != 0)
                    {
                        ws.Cells[rowCurrent, 12].value = item.qty7ErrorPosition;
                    }

                    if (item.qty8Reverse != 0)
                    {
                        ws.Cells[rowCurrent, 13].value = item.qty8Reverse;
                    }

                    if (item.qty9DirectionRev != 0)
                    {
                        ws.Cells[rowCurrent, 14].value = item.qty9DirectionRev;
                    }

                    if (item.qty10OjectForeign != 0)
                    {
                        ws.Cells[rowCurrent, 15].value = item.qty10OjectForeign;
                    }

                    if (item.qty11ItemMiss != 0)
                    {
                        ws.Cells[rowCurrent, 16].value = item.qty11ItemMiss;
                    }
                    if (item.qty12Peel != 0)
                    {
                        ws.Cells[rowCurrent, 17].value = item.qty12Peel;
                    }
                    if (item.qty13Other != 0)
                    {
                        ws.Cells[rowCurrent, 18].value = item.qty13Other;
                    }


                    rowCurrent++;

                }

                range = ws.Range["A3:T" + (rowCurrent)];
                range.Borders.LineStyle = XlLineStyle.xlContinuous;
                range.Borders.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
                range.Borders.Weight = XlBorderWeight.xlThin;


                rowCurrent = 3;
                for (int i = 0; i <= listKyocera.Count; i++)
                {
                    ws.Range["S" + rowCurrent + ":T" + rowCurrent].Merge();

                    ws.Cells[rowCurrent, "D"].Formula = string.Format("=C{0}/B{0}*1000000", rowCurrent);
                    ws.Cells[rowCurrent, "C"].Formula = string.Format("=SUM(F{0}:R{0}", rowCurrent);
                    rowCurrent++;
                }
                rowCurrent--;

                ws.Range["A3:T" + (rowCurrent)].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws.Cells[rowCurrent, "A"].Value = @"合計";

                ws.Cells[rowCurrent, "B"].Formula = @"= SUM(B3:B" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "C"].Formula = @"= SUM(C3:C" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "D"].Formula = @"= SUM(D3:D" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "E"].Formula = @"= SUM(E3:E" + (rowCurrent - 1) + ")";

                ws.Cells[rowCurrent, "F"].Formula = @"= SUM(F3:F" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "G"].Formula = @"= SUM(G3:G" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "H"].Formula = @"= SUM(H3:H" + (rowCurrent - 1) + ")";

                ws.Cells[rowCurrent, "I"].Formula = @"= SUM(I3:I" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "J"].Formula = @"= SUM(J3:J" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "K"].Formula = @"= SUM(K3:K" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "L"].Formula = @"= SUM(L3:L" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "M"].Formula = @"= SUM(M3:M" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "N"].Formula = @"= SUM(N3:N" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "O"].Formula = @"= SUM(O3:O" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "P"].Formula = @"= SUM(P3:P" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "Q"].Formula = @"= SUM(Q3:Q" + (rowCurrent - 1) + ")";
                ws.Cells[rowCurrent, "R"].Formula = @"= SUM(R3:R" + (rowCurrent - 1) + ")";
                

                ws.Cells[rowCurrent, "D"].Formula = string.Format("=C{0}/B{0}*1000000", rowCurrent);




                wb.Save();

                wb.Close();
                app.Quit();

                return RESULT.OK;
            }
            //catch (Exception ex)
            //{
            //    return string.Format(RESULT.ERROR_015_CATCH, "WriteTSB1", ex.Message);
            //}
            finally
            {
                if (range != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                }
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
