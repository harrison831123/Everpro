using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.PayRoll.Models
{
    public class TotalReasonCodeReportModel
    {
        //原因碼
        public string reason_code { get; set; }
        //原因碼名稱
        public string remark { get; set; }
        //筆數
        public string cnt { get; set; }
        //調整金額
        public int amount { get; set; }
        //名字
        public string nmember { get; set; }
        //部門
        public string nunit { get; set; }
    }
}
