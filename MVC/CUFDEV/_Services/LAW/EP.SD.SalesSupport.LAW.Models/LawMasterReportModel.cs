using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.LAW.Models
{
    public class LawMasterReportModel
    {
        public string DateNow { get; set; }
        public string DateHave { get; set; }
        /// <summary>結欠金額</summary>
        public decimal LawDueMoney { get; set; }

        /// <summary>結欠總金額</summary>
        public int LawTotalDue { get; set; }

        /// <summary>累計清償金額</summary>
        public int LawTotalRepayment { get; set; }

        /// <summary>清償金額</summary>
        public int LawRepaymentMoney { get; set; }

        /// <summary>達成率</summary>
        public string LawRepayPercent { get; set; }

        public string pstr { get; set; }

        public string repayperstr { get; set; }

        public string year { get; set; }

        public string Lawyear { get; set; }

        public int sumtype { get; set; }

        public string LawYearType { get; set; }
    }
}
