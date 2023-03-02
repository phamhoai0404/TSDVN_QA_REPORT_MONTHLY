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
}
