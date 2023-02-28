using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_REPORT_MONTHLY.MODEL
{
    public class DataFirst
    {
        public string  model { get; set; }
        public string cusDetail { get; set; }
        public string wo { get; set; }
        public long qty { get; set; }
        public string cusCode { get; set; }
        public string cusName { get; set; }

        public DataFirst()
        {

        }
        public DataFirst(DataFirst s)
        {
            this.model = s.model;
            this.cusCode = s.cusCode;
            this.cusDetail = s.cusDetail;

            this.qty = s.qty;
            this.wo = s.wo;

            
            this.cusName = s.cusName;

        }
    }

}
