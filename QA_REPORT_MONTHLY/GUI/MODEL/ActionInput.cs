using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_REPORT_MONTHLY.MODEL
{
    public class ActionInput1
    {
        public string monthString { get; set; }
        public int monthInt { get; set; }
        public string fileData { get; set; }
        public string fileError { get; set; }
        

        public override string ToString()
        {
            return monthString.ToString()  + "," + fileData.ToString() + ","  + fileError.ToString();
        }

    }
    public class ActionInput2
    {
        public string colModel { get; set; }
        public int rowStart { get; set; }
        public int rowEnd { get; set; }
        public string sheetName { get; set; }

        public string rowStartString { get; set; }
        public string rowEndString { get; set; }
        public string fileData { get; set; }
       
        public void Trim()
        {
            this.rowEndString = this.rowEndString.Trim();
            this.rowStartString = this.rowStartString.Trim();
            this.fileData = this.fileData.Trim();
            this.colModel = this.colModel.Trim();
            this.sheetName = this.sheetName.Trim();
        }
    }
}
