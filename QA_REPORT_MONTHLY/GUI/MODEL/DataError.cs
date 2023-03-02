using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_REPORT_MONTHLY.MODEL
{
    public class DataError
    {
        public string model { get; set; }
        public string wo { get; set; }
        public string nameError { get; set; }
        public string noteNameError { get; set; }

        public string dept { get; set; }

        public string cusCode { get; set; }

        public int qty { get; set; }

        public DataError()
        {
        }
        public DataError(DataError s)
        {
            this.nameError = s.nameError;
            this.model = s.model;
            this.wo = s.wo;
            this.cusCode = s.cusCode;
            this.qty = s.qty;
            this.dept = s.dept;
            this.noteNameError = s.noteNameError;
        }
        public override string ToString()
        {
            return model + ";" + wo + ";" + qty + ";" + dept + ";" + nameError;
        }

        //public int qty1WeldFake { get; set; }
        //public int qty2ErrorPosition { get; set; }
        //public int qty3Warp{ get; set; }
        //public int qty4BrightMake { get; set; }
        //public int qty5TinSmall { get; set; }
        //public int qty6ItemLack { get; set; }
        //public int qty7ErrorPosition { get; set; }
        //public int qty8Reverse { get; set; }
        //public int qty9DirectionRev{ get; set; }
        //public int qty10ItemWrong { get; set; }
        //public int qty11OjectForeign { get; set; }
        //public int qty12ItemMiss { get; set; }
        //public int qty13Peel{ get; set; }
        //public int qty14Other { get; set; }
    }
}
