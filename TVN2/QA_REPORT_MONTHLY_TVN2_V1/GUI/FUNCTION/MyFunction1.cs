using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_REPORT_MONTHLY.FUNCTION
{
    public class MyFunction1
    {
        public static DataTable getDataExcel(string filePath, string sheetName)
        {
            try
            {
                //Tao chuoi ket noi voi Excel
                string connectExcel = string.Empty;
                connectExcel = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=YES'";
                connectExcel = string.Format(connectExcel, filePath);

                //Phan thuc hien doc du lieu tu file excel
                DataTable dtExcel = new DataTable();
                using (OleDbConnection connExcel = new OleDbConnection(connectExcel))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;
                            OpenExcel(connExcel);

                            DataTable dtExcelSchema;
                            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            //sheetName = dtExcelSchema.Rows[1]["TABLE_NAME"].ToString();
                            cmdExcel.CommandText = "SELECT *from [" + sheetName + "$]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dtExcel);

                            CloseExcel(connExcel);
                        }
                    }
                }
                return dtExcel;
            }
            catch (Exception)
            {
                return null;
            }

            


        }
        private static void OpenExcel(OleDbConnection conn)
        {
            if(conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
        }
        private static void CloseExcel(OleDbConnection conn)
        {
            if (conn != null && conn.State == ConnectionState.Open)
                conn.Close();
        }
    }
}
