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

        public const string ERROR_FILE_ERROR_WO = "Trong file lỗi tồn tại WO: {0}  trong file dữ liệu!  Chi tiết dòng dữ liệu lỗi: {1}";

        public const string ERROR_FILE_ERROR_MODEL = "Không tồn tại model: {0} - trong dữ liệu mà chỉ tồn tại trên file lỗi!";

        public const string ERROR_FILE_ERROR_NOT_COMMENT = "Không có comment khi thuộc loại Thừa thiếu linh kiện ở dòng: {0}";
        public const string ERROR_FILE_ERROR_COMMENT_NOT_RULE = "Dòng: {0} - thuộc linh kiện thiếu thừa nhưng note không theo quy tắc (chứa từ thiếu thừa) - note: {1}";
        //const string k = "10";

         public const string ERROR_2_INPUT_NOT_NUMBER = "{0} cần nhập vào là số !";
         public const string ERROR_2_INPUT_NUMBER_RULE = "Dòng kết thúc > dòng bắt đầu !";


        public const string ERROR_2_NOT_NULL_MODEL = "Không được để trống model dữ liệu model ở dòng: {0}";
        public const string ERROR_2_NOT_OPEN = "Model không có tên khách hàng ở dòng: {0}";
        public const string ERROR_2_NOT_NUMBER = "Cột {0} không phải là số ở dòng: {1}";
    }
}
