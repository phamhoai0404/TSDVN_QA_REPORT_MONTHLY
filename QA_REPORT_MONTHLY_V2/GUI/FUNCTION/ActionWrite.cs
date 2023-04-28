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
            DataConfig.CONFIG_FILE_RESULT = DateTime.Now.ToString("1_ddMMyyyy_HHmmss");
            DataConfig.CONFIG_FILE_RESULT = currentPath + @"\RESULT\" + DataConfig.CONFIG_FILE_RESULT + ".xlsx";

            Excel.Application excel = new Excel.Application();
            Excel.Workbook originalWorkbook = excel.Workbooks.Open(pathTemplate);

            originalWorkbook.SaveCopyAs(DataConfig.CONFIG_FILE_RESULT);
            originalWorkbook.Close(false);
            excel.Quit();

            return RESULT.OK;
        }

        /// <summary>
        /// Thuc hien ghi du lieu cua TSB
        /// </summary>
        /// <param name="listTSB"></param>
        /// <returns></returns>
        /// CreatedBy: HoaiPT(06/03/2023)
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

                int rowCurrent = 3;//Dong du lieu bat dau  tu dong thu 3
                listTSB = listTSB.OrderBy(o => o.item).ToList();//Thuc hien sap xep theo ten model
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

                    rowCurrent++;//Tang len dong tiep theo ma thoi
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

                    ws.Cells[rowCurrent, "F"].Formula = string.Format("=E{0}/C{0}*1000000", rowCurrent);
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
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "WriteTSB1", ex.Message);
            }
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
        /// <summary>
        /// Thuc hien ghi du lieu cua Kyocera
        /// </summary>
        /// <param name="listKyocera"></param>
        /// <returns></returns>
        /// CreatedBy: HoaiPT(06/03/2023)
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
                listKyocera = listKyocera.OrderBy(o => o.item).ToList();
                foreach (var item in listKyocera)
                {
                    ws.Cells[rowCurrent, "A"].value = item.item + "(" + item.cus + ")";
                    ws.Cells[rowCurrent, "B"].value = item.qtySum;
                    ws.Cells[rowCurrent, "E"].value = 80;

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
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "WriteKyocera1", ex.Message);
            }
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

        public static string WriteFX(List<DataFX> listFX)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;

            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(DataConfig.CONFIG_FILE_RESULT);

                ws = wb.Sheets["FX_1"];

                int rowCurrent = 3;
                listFX = listFX.OrderBy(o => o.item).ToList();
                foreach (var item in listFX)
                {
                    ws.Cells[rowCurrent, "A"].value = item.item + item.cus;
                    ws.Cells[rowCurrent, "B"].value = item.qtySum;
                    ws.Cells[rowCurrent, "E"].value = 80;

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

                    if (item.qty8Reverse != 0)
                    {
                        ws.Cells[rowCurrent, 11].value = item.qty8Reverse;
                    }

                    if (item.qty9DirectionRev != 0)
                    {
                        ws.Cells[rowCurrent, 12].value = item.qty9DirectionRev;
                    }

                    if (item.qty10OjectForeign != 0)
                    {
                        ws.Cells[rowCurrent, 13].value = item.qty10OjectForeign;
                    }

                    if (item.qty11ItemMiss != 0)
                    {
                        ws.Cells[rowCurrent, 14].value = item.qty11ItemMiss;
                    }
                    if (item.qty12Peel != 0)
                    {
                        ws.Cells[rowCurrent, 15].value = item.qty12Peel;
                    }
                    if (item.qty13Other != 0)
                    {
                        ws.Cells[rowCurrent, 16].value = item.qty13Other;
                    }


                    rowCurrent++;

                }

                range = ws.Range["A3:R" + (rowCurrent)];
                range.Borders.LineStyle = XlLineStyle.xlContinuous;
                range.Borders.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
                range.Borders.Weight = XlBorderWeight.xlThin;


                rowCurrent = 3;
                for (int i = 0; i <= listFX.Count; i++)
                {
                    ws.Range["Q" + rowCurrent + ":R" + rowCurrent].Merge();

                    ws.Cells[rowCurrent, "D"].Formula = string.Format("=C{0}/B{0}*1000000", rowCurrent);
                    ws.Cells[rowCurrent, "C"].Formula = string.Format("=SUM(F{0}:R{0}", rowCurrent);
                    rowCurrent++;
                }
                rowCurrent--;

                ws.Range["A3:R" + (rowCurrent)].HorizontalAlignment = XlHAlign.xlHAlignCenter;
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

                ws.Cells[rowCurrent, "D"].Formula = string.Format("=C{0}/B{0}*1000000", rowCurrent);

                wb.Save();
                wb.Close();
                app.Quit();

                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "WriteFX", ex.Message);
            }
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

        public static string WriteHT(DataHT valueHT, string month)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;

            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(DataConfig.CONFIG_FILE_RESULT);

                ws = wb.Sheets["HITACHI_1"];

                int rowCurrent = 3;

                ws.Cells[rowCurrent, "A"].value = month + "月";
                ws.Cells[rowCurrent, "C"].value = valueHT.qtySum;
                ws.Cells[rowCurrent, "D"].Formula = string.Format(@"= SUM(G{0}:R{0}", rowCurrent);
                ws.Cells[rowCurrent, "E"].value = 100;
                ws.Cells[rowCurrent, "F"].Formula = string.Format("=IF(D{0}=\"\",\"\",D{0}/C{0}*1000000)", rowCurrent);

                if (valueHT.qty8Reverse != 0)
                {
                    ws.Cells[rowCurrent, "G"].value = valueHT.qty8Reverse;
                }
                if (valueHT.qty6ItemLack != 0)
                {
                    ws.Cells[rowCurrent, "H"].value = valueHT.qty6ItemLack;
                }
                if (valueHT.qty2ErrorPosition != 0)
                {
                    ws.Cells[rowCurrent, "I"].value = valueHT.qty2ErrorPosition;
                }
                if (valueHT.qty11ItemMiss != 0)
                {
                    ws.Cells[rowCurrent, "J"].value = valueHT.qty11ItemMiss;
                }
                if (valueHT.qty1WeldFake != 0)
                {
                    ws.Cells[rowCurrent, "L"].value = valueHT.qty1WeldFake;
                }
                if (valueHT.qty3Warp != 0)
                {
                    ws.Cells[rowCurrent, "M"].value = valueHT.qty3Warp;
                }
                if (valueHT.qty4BrightMake != 0)
                {
                    ws.Cells[rowCurrent, "P"].value = valueHT.qty4BrightMake;
                }
                if (valueHT.qty13Other != 0)
                {
                    ws.Cells[rowCurrent, "R"].value = valueHT.qty13Other;
                }


                wb.Save();

                wb.Close();
                app.Quit();

                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "WriteHT", ex.Message);
            }
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
        /// <summary>
        /// Thuc hien ghi du lieu cua Okidenki
        /// </summary>
        /// <param name="valueOkidenki"></param>
        /// <param name="month"></param>
        /// <returns></returns>
        public static string WriteOkidenki(DataOkidenki valueOkidenki, string month)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;

            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(DataConfig.CONFIG_FILE_RESULT);

                ws = wb.Sheets["OKIDENKI_1"];

                int rowCurrent = 3;

                ws.Cells[rowCurrent, "A"].value = month + "月";
                ws.Cells[rowCurrent, "C"].value = valueOkidenki.qtySum;
                ws.Cells[rowCurrent, "D"].Formula = string.Format(@"= SUM(G{0}:T{0}", rowCurrent);
                ws.Cells[rowCurrent, "E"].Formula = string.Format("=D{0}/C{0}*1000000", rowCurrent);
                ws.Cells[rowCurrent, "F"].value = 80;

                if (valueOkidenki.qty1WeldFake != 0)
                {
                    ws.Cells[rowCurrent, "G"].value = valueOkidenki.qty1WeldFake;
                }
                if (valueOkidenki.qty2ErrorPosition != 0)
                {
                    ws.Cells[rowCurrent, "H"].value = valueOkidenki.qty2ErrorPosition;
                }
                if (valueOkidenki.qty3Warp != 0)
                {
                    ws.Cells[rowCurrent, "I"].value = valueOkidenki.qty3Warp;
                }
                if (valueOkidenki.qty4BrightMake != 0)
                {
                    ws.Cells[rowCurrent, "J"].value = valueOkidenki.qty4BrightMake;
                }
                if (valueOkidenki.qty5TinSmall != 0)
                {
                    ws.Cells[rowCurrent, "K"].value = valueOkidenki.qty5TinSmall;
                }
                if (valueOkidenki.qty6ItemLack != 0)
                {
                    ws.Cells[rowCurrent, "L"].value = valueOkidenki.qty6ItemLack;
                }

                if (valueOkidenki.qty8Reverse != 0)
                {
                    ws.Cells[rowCurrent, "N"].value = valueOkidenki.qty8Reverse;
                }
                if (valueOkidenki.qty9DirectionRev != 0)
                {
                    ws.Cells[rowCurrent, "O"].value = valueOkidenki.qty9DirectionRev;
                }
                if (valueOkidenki.qty10OjectForeign != 0)
                {
                    ws.Cells[rowCurrent, "P"].value = valueOkidenki.qty10OjectForeign;
                }
                if (valueOkidenki.qty11ItemMiss != 0)
                {
                    ws.Cells[rowCurrent, "Q"].value = valueOkidenki.qty11ItemMiss;
                }
                if (valueOkidenki.qty12Peel != 0)
                {
                    ws.Cells[rowCurrent, "R"].value = valueOkidenki.qty12Peel;
                }
                if (valueOkidenki.qty13Other != 0)
                {
                    ws.Cells[rowCurrent, "T"].value = valueOkidenki.qty13Other;
                }

                wb.Save();
                wb.Close();
                app.Quit();

                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "WriteOkidenki", ex.Message);
            }
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
        
        /// <summary>
        /// Thuc hien ghi du lieu cua riso
        /// </summary>
        /// <param name="valueRiso"></param>
        /// <param name="month"></param>
        /// <returns></returns>
        public static string WriteRiso(DataRiso valueRiso, string month)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;

            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(DataConfig.CONFIG_FILE_RESULT);

                ws = wb.Sheets["RISO_1"];

                int rowCurrent = 3;

                ws.Cells[rowCurrent, "A"].value = month + "月";
                ws.Cells[rowCurrent, "C"].value = valueRiso.qtySum;
                ws.Cells[rowCurrent, "D"].Formula = string.Format(@"= SUM(G{0}:S{0}", rowCurrent);
                ws.Cells[rowCurrent, "E"].Formula = string.Format("=D{0}/C{0}*1000000", rowCurrent);
                ws.Cells[rowCurrent, "F"].value = 80;

                if (valueRiso.qty1WeldFake != 0)
                {
                    ws.Cells[rowCurrent, "G"].value = valueRiso.qty1WeldFake;
                }
                if (valueRiso.qty2ErrorPosition != 0)
                {
                    ws.Cells[rowCurrent, "H"].value = valueRiso.qty2ErrorPosition;
                }
                if (valueRiso.qty3Warp != 0)
                {
                    ws.Cells[rowCurrent, "I"].value = valueRiso.qty3Warp;
                }
                if (valueRiso.qty4BrightMake != 0)
                {
                    ws.Cells[rowCurrent, "J"].value = valueRiso.qty4BrightMake;
                }
                if (valueRiso.qty5TinSmall != 0)
                {
                    ws.Cells[rowCurrent, "K"].value = valueRiso.qty5TinSmall;
                }
                if (valueRiso.qty6ItemLack != 0)
                {
                    ws.Cells[rowCurrent, "L"].value = valueRiso.qty6ItemLack;
                }

                if (valueRiso.qty8Reverse != 0)
                {
                    ws.Cells[rowCurrent, "N"].value = valueRiso.qty8Reverse;
                }
                if (valueRiso.qty9DirectionRev != 0)
                {
                    ws.Cells[rowCurrent, "O"].value = valueRiso.qty9DirectionRev;
                }
                if (valueRiso.qty10OjectForeign != 0)
                {
                    ws.Cells[rowCurrent, "P"].value = valueRiso.qty10OjectForeign;
                }
                if (valueRiso.qty11ItemMiss != 0)
                {
                    ws.Cells[rowCurrent, "Q"].value = valueRiso.qty11ItemMiss;
                }
                if (valueRiso.qty12Peel != 0)
                {
                    ws.Cells[rowCurrent, "R"].value = valueRiso.qty12Peel;
                }
                if (valueRiso.qty13Other != 0)
                {
                    ws.Cells[rowCurrent, "S"].value = valueRiso.qty13Other;
                }


                wb.Save();
                wb.Close();
                app.Quit();

                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "WriteRiso", ex.Message);
            }
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
        public static string WriteJCM(DataJCM valueJCM, string month)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;

            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(DataConfig.CONFIG_FILE_RESULT);

                ws = wb.Sheets["JCM_1"];

                int rowCurrent = 3;

                ws.Cells[rowCurrent, "A"].value = month + "月";
                ws.Cells[rowCurrent, "C"].value = valueJCM.qtySum;
                ws.Cells[rowCurrent, "D"].Formula = string.Format(@"= SUM(G{0}:S{0}", rowCurrent);
                ws.Cells[rowCurrent, "E"].Formula = string.Format("=D{0}/C{0}*1000000", rowCurrent);
                ws.Cells[rowCurrent, "F"].value = 80;

                if (valueJCM.qty1WeldFake != 0)
                {
                    ws.Cells[rowCurrent, "G"].value = valueJCM.qty1WeldFake;
                }
                if (valueJCM.qty14LechLK != 0)
                {
                    ws.Cells[rowCurrent, "H"].value = valueJCM.qty14LechLK;
                }

                if (valueJCM.qty4BrightMake != 0)
                {
                    ws.Cells[rowCurrent, "J"].value = valueJCM.qty4BrightMake;
                }

                if (valueJCM.qty5TinSmall != 0)
                {
                    ws.Cells[rowCurrent, "K"].value = valueJCM.qty5TinSmall;
                }
                if (valueJCM.qty6ItemLack != 0)
                {
                    ws.Cells[rowCurrent, "L"].value = valueJCM.qty6ItemLack;
                }
                if (valueJCM.qty8Reverse != 0)
                {
                    ws.Cells[rowCurrent, "M"].value = valueJCM.qty8Reverse;
                }

                if (valueJCM.qty11ItemMiss != 0)
                {
                    ws.Cells[rowCurrent, "N"].value = valueJCM.qty11ItemMiss;
                }

                if (valueJCM.qty10OjectForeign != 0)
                {
                    ws.Cells[rowCurrent, "P"].value = valueJCM.qty10OjectForeign;
                }
                if (valueJCM.qty12Peel != 0)
                {
                    ws.Cells[rowCurrent, "Q"].value = valueJCM.qty12Peel;
                }

                if (valueJCM.qty13Other != 0)
                {
                    ws.Cells[rowCurrent, "S"].value = valueJCM.qty13Other;
                }




                wb.Save();

                wb.Close();
                app.Quit();

                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "WriteRiso", ex.Message);
            }
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

        /// <summary>
        /// Thuc hien tao file tu teamplate
        /// </summary>
        /// <returns></returns>
        public static string CreateFile_2()
        {
            string currentPath = Directory.GetCurrentDirectory();
            string pathTemplate = currentPath + DataConfig.CONFIG_2_FILE_TEMPLATE;
            DataConfig.CONFIG_FILE_RESULT = DateTime.Now.ToString("2_ddMMyyyy_HHmmss");
            DataConfig.CONFIG_FILE_RESULT = currentPath + @"\RESULT\" + DataConfig.CONFIG_FILE_RESULT + ".xlsx";

            Excel.Application excel = new Excel.Application();
            Excel.Workbook originalWorkbook = excel.Workbooks.Open(pathTemplate);

            originalWorkbook.SaveCopyAs(DataConfig.CONFIG_FILE_RESULT);
            originalWorkbook.Close(false);
            excel.Quit();

            return RESULT.OK;
        }
        /// <summary>
        /// Thuc hien ghi du lieu cua Kyocrea
        /// </summary>
        /// <param name="listKyocrea"></param>
        /// <returns></returns>
        public static string WriteKyocera_2(List<DataKyocera> listKyocrea)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;

            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(DataConfig.CONFIG_FILE_RESULT);

                ws = wb.Sheets["KYOCERA_2"];

                listKyocrea = listKyocrea.OrderBy(o => o.item).ToList();
                int rowCurrent = 3;
                foreach (var item in listKyocrea)
                {
                    ws.Cells[rowCurrent, "A"].value = item.cus;
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
                for (int i = 0; i <= listKyocrea.Count; i++)
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
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "WriteKyocera_2", ex.Message);
            }
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


        public static string CreateFile_3(string fileName)
        {
            string currentPath = Directory.GetCurrentDirectory();
            string pathTemplate = fileName;
            DataConfig.CONFIG_FILE_RESULT = DateTime.Now.ToString("3_ddMMyyyy_HHmmss");
            DataConfig.CONFIG_FILE_RESULT = currentPath + @"\RESULT\" + DataConfig.CONFIG_FILE_RESULT + ".xlsx";

            Excel.Application excel = new Excel.Application();
            Excel.Workbook originalWorkbook = excel.Workbooks.Open(pathTemplate);

            originalWorkbook.SaveCopyAs(DataConfig.CONFIG_FILE_RESULT);
            originalWorkbook.Close(false);
            excel.Quit();

            return RESULT.OK;
        }
        public static string WriteTSB_3(ActionInput3 valueInput, ref List<DataTSB3>  listDataTSB)
        {
            CreateFile_3(valueInput.fileInput);//Thuc hien tao file

            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;


            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(DataConfig.CONFIG_FILE_RESULT);

                ws = wb.Sheets[valueInput.sheetName];

                string colModel = valueInput.colModel;
                string colWrite = valueInput.colWrite;
                for (int i = valueInput.rowStart; i <= valueInput.rowEnd; i = i + 2)
                {
                    string tempModel = ws.Cells[i, colModel].value;
                    if (string.IsNullOrWhiteSpace(tempModel))
                    {
                        return string.Format(RESULT.ERROR_2_NOT_NULL_MODEL, i);
                    }
                    var tempObject = new DataTSB3();
                    foreach (var currentRow in listDataTSB)
                    {
                        if (currentRow.item.Equals(tempModel)){
                            currentRow.action = true;
                            tempObject = currentRow;
                            break;
                        }
                    }
                    if(tempObject.item== null)//Neu khong co thi chuyen sang doi tuong khac
                    {
                        continue;
                    }

                    ws.Cells[i, colWrite].value = tempObject.qtySum;
                    ws.Cells[i + 1, colWrite].value = tempObject.qtyErrorSum;

                    //Dinh dang lai file
                    ws.Cells[i, colWrite].Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone; ;
                    ws.Cells[i + 1, colWrite].Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                }



                wb.Save();

                wb.Close();
                app.Quit();

                return RESULT.OK;
            }
            catch (Exception ex)
            {
                return string.Format(RESULT.ERROR_015_CATCH, "WriteKyocera_2", ex.Message);
            }
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
