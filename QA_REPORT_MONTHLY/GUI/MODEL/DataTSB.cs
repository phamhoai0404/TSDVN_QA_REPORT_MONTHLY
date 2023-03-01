using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_REPORT_MONTHLY.MODEL
{
    public class DataTSB
    {
        public string item { get; set; }
        public string cus { get; set; }

        public long qtySum { get; set; }

        public int qty1WeldFake { get; set; }
        public int qty2ErrorPosition { get; set; }
        public int qty3Warp { get; set; }
        public int qty4BrightMake { get; set; }
        public int qty5TinSmall { get; set; }
        public int qty6ItemLack { get; set; }
        public int qty7ErrorPosition { get; set; }
        public int qty8Reverse { get; set; }
        public int qty9DirectionRev { get; set; }
        public int qty10ItemWrong { get; set; }
        public int qty11OjectForeign { get; set; }
        public int qty12ItemMiss { get; set; }
        public int qty13Peel { get; set; }
        public int qty14Other { get; set; }

        public DataTSB()
        {

        }
        public DataTSB(string item, string cus, long qtySum)
        {
            this.item = item;
            this.cus = cus;
            this.qtySum = qtySum;
        }

    }
}
