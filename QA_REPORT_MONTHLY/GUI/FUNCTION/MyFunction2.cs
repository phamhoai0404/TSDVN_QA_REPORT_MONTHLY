using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QA_REPORT_MONTHLY.MODEL;

namespace QA_REPORT_MONTHLY.FUNCTION
{
    public class MyFunction2
    {
        public static string GetDataConfig(string pathFile, ref Dictionary<string, object> getConfig)
        {
            try
            {
                DataTable temp = new DataTable();
                temp = MyFunction1.getDataExcel(pathFile, "Sheet1");


                foreach (DataRow currentRow in temp.Rows)
                {
                    if (!string.IsNullOrEmpty(currentRow[0].ToString().Trim()))
                    {
                        getConfig[currentRow[0].ToString()] = currentRow[1].ToString();
                    }
                    else
                    {
                        break;
                    }
                }
              

                return RESULT.OK;
            }
            catch (Exception ex)
            {

                return string.Format(RESULT.ERROR_015_CATCH, "GetDataConfig", ex.Message);
            }
        }

        /// <summary>
        /// Thuc hien select file 
        /// </summary>
        /// <returns>
        /// Tra ve ket qua la dia chi file; hoac khong chon file nao; hoac nhay vao catch
        /// </returns>
        /// CreatedBy: HoaiPT(01/02/2023)
        public static string SelectFile()
        {
            try
            {
                using (var ofd = new System.Windows.Forms.OpenFileDialog())
                {
                    ofd.Filter = MdlComment.TYPE_FILE_SELECT;
                    if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {

                        return ofd.FileName;
                    }
                }
                return RESULT.OK;
            }
            catch (Exception)
            {
                return RESULT.ERROR_SELECT_FILE;
            }


        }

        /// <summary>
        /// Thuc hien dong tat ca ten .exe la nameProcess (nhu dong o task manager)
        /// </summary>
        /// <param name="nameProcess">Ten muon Skill</param>
        /// <returns></returns>
        /// CreatedBy: HoaiPT(Su dung tu lau nhung h moi chinh thuc dua ra function: 22/11/2022)
        public static bool Skill_Process(string nameProcess)
        {
            try
            {
                foreach (var process in Process.GetProcessesByName(nameProcess))
                {
                    process.Kill();
                }
                return true;
            }
            catch
            {
                return false;
            }

        }
    }
}
