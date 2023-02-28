using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA_REPORT_MONTHLY.MODEL
{
    public static class RESULT
    {
        public const string OK = "OK";
        public const string ERROR_015_CATCH = "Lỗi {0} Catch - {1}";

        public const string ERROR_SELECT_FILE = "Có lỗi xảy ra trong quá trình chọn file!";


        public const string ERROR_NOT_NULL = "Không được để trống các trường dữ liệu!";
        public const string ERROR_NOT_NUMBER = "Tháng nhập vào không phải là số nguyên!";
        public const string ERROR_NOT_MONTH = "Bạn nhập vào không đúng định dạng tháng >=1 và <=12 !";

        public const string ERROR_NOT_FILE = "Không tồn tại file trong hệ thống! File: {0}";
        public const string ERROR_SHEETNAME = "Không tồn tại SheetName: {0} - Trong hệ thống!";

        public const string ERROR_COLUMN_G = "Bạn chưa nhập khách hàng ở địa chỉ Cột G dòng: {0} - Lỗi catch: {1}";
        public const string ERROR_COLUMN_C = "Bạn chưa nhập Chi  tiết Mã khách hàng ở địa chỉ Cột C dòng: {0} - Lỗi catch: {1}";


    }
}
