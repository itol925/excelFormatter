using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelFormatter {
    class Record {
        public string declarant;    // 申请人
        public string compensate;   // 补偿
        public string Approval;     // 审批意见
        public List<Bug> bugs = new List<Bug>();
    }
    class Bug {
        public string date;         // 日期
        public string bugDesc;      // bug
        public string suggest;      // 建议
    }
}
