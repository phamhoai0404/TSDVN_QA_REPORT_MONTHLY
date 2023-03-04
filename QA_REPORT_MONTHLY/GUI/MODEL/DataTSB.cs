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
        public int qty10OjectForeign { get; set; }
        public int qty11ItemMiss { get; set; }
        public int qty12Peel { get; set; }
        public int qty13Other { get; set; }

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
    public class DataFX
    {
        public string item { get; set; }
        public string cus { get; set; }

        public long qtySum { get; set; }

        public int qty1WeldFake { get; set; }
        public int qty2ErrorPosition { get; set; }
        public int qty3Warp { get; set; }
        public int qty4BrightMake { get; set; }
        public int qty5TinSmall { get; set; }

        public int qty8Reverse { get; set; }
        public int qty9DirectionRev { get; set; }
        public int qty10OjectForeign { get; set; }
        public int qty11ItemMiss { get; set; }
        public int qty12Peel { get; set; }
        public int qty13Other { get; set; }

        public DataFX()
        {

        }
        public DataFX(string item, string cus, long qtySum)
        {
            this.item = item;
            this.cus = cus;
            this.qtySum = qtySum;
        }

    }

    public class DataKyocera
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
        public int qty10OjectForeign { get; set; }
        public int qty11ItemMiss { get; set; }
        public int qty12Peel { get; set; }
        public int qty13Other { get; set; }

        public DataKyocera()
        {

        }
        public DataKyocera(DataKyocera s)
        {
            this.item = s.item;
            this.cus = s.cus;

            this.qtySum = s.qtySum;

            this.qty1WeldFake = s.qty1WeldFake;
            this.qty2ErrorPosition = s.qty2ErrorPosition;
            this.qty3Warp = s.qty3Warp;
            this.qty4BrightMake = s.qty4BrightMake;
            this.qty5TinSmall = s.qty5TinSmall;
            this.qty6ItemLack = s.qty6ItemLack;
            this.qty7ErrorPosition = s.qty7ErrorPosition;
            this.qty8Reverse = s.qty8Reverse;
            this.qty9DirectionRev = s.qty9DirectionRev;
            this.qty10OjectForeign = s.qty10OjectForeign;
            this.qty11ItemMiss = s.qty11ItemMiss;
            this.qty12Peel = s.qty12Peel;
            this.qty13Other = s.qty13Other;
        }


        public void Pus(DataKyocera s)
        {
            this.qtySum += s.qtySum;

            this.qty1WeldFake += s.qty1WeldFake;
            this.qty2ErrorPosition += s.qty2ErrorPosition;
            this.qty3Warp += s.qty3Warp;
            this.qty4BrightMake += s.qty4BrightMake;
            this.qty5TinSmall += s.qty5TinSmall;
            this.qty6ItemLack += s.qty6ItemLack;
            this.qty7ErrorPosition += s.qty7ErrorPosition;
            this.qty8Reverse += s.qty8Reverse;
            this.qty9DirectionRev += s.qty9DirectionRev;
            this.qty10OjectForeign += s.qty10OjectForeign;
            this.qty11ItemMiss += s.qty11ItemMiss;
            this.qty12Peel += s.qty12Peel;
            this.qty13Other += s.qty13Other;
        }
        public DataKyocera(string item, string cus, long qtySum)
        {
            this.item = item;
            this.cus = cus;
            this.qtySum = qtySum;
        }

    }

    public class DataHT
    {
        public long qtySum { get; set; }

        public int qty1WeldFake { get; set; }
        public int qty2ErrorPosition { get; set; }
        public int qty3Warp { get; set; }
        public int qty4BrightMake { get; set; }
        //public int qty5TinSmall { get; set; }
        public int qty6ItemLack { get; set; }
        //public int qty7ErrorPosition { get; set; }
        public int qty8Reverse { get; set; }
        //public int qty9DirectionRev { get; set; }
        //public int qty10OjectForeign { get; set; }
        public int qty11ItemMiss { get; set; }
        //public int qty12Peel { get; set; }
        public int qty13Other { get; set; }
    }
    public class DataOkidenki
    {
        public long qtySum { get; set; }

        public int qty1WeldFake { get; set; }
        public int qty2ErrorPosition { get; set; }
        public int qty3Warp { get; set; }
        public int qty4BrightMake { get; set; }
        public int qty5TinSmall { get; set; }
        public int qty6ItemLack { get; set; }
        // public int qty7ErrorPosition { get; set; }
        public int qty8Reverse { get; set; }
        public int qty9DirectionRev { get; set; }
        public int qty10OjectForeign { get; set; }
        public int qty11ItemMiss { get; set; }
        public int qty12Peel { get; set; }
        public int qty13Other { get; set; }

    }
    public class DataRiso
    {
        public long qtySum { get; set; }

        public int qty1WeldFake { get; set; }
        public int qty2ErrorPosition { get; set; }
        public int qty3Warp { get; set; }
        public int qty4BrightMake { get; set; }
        public int qty5TinSmall { get; set; }
        public int qty6ItemLack { get; set; }
        // public int qty7ErrorPosition { get; set; }
        public int qty8Reverse { get; set; }
        public int qty9DirectionRev { get; set; }
        public int qty10OjectForeign { get; set; }
        public int qty11ItemMiss { get; set; }
        public int qty12Peel { get; set; }
        public int qty13Other { get; set; }

    }
    public class DataJCM
    {
        public long qtySum { get; set; }

        public int qty1WeldFake { get; set; }
        //public int qty2ErrorPosition { get; set; }
        //public int qty3Warp { get; set; }
        public int qty4BrightMake { get; set; }
        public int qty5TinSmall { get; set; }
        public int qty6ItemLack { get; set; }
        //// public int qty7ErrorPosition { get; set; }
        public int qty8Reverse { get; set; }
        //public int qty9DirectionRev { get; set; }
        public int qty10OjectForeign { get; set; }
        public int qty11ItemMiss { get; set; }
        public int qty12Peel { get; set; }
        public int qty13Other { get; set; }

        public int qty14LechLK { get; set; }

    }
}
