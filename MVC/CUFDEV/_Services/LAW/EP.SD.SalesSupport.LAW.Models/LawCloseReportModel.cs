using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.LAW.Models
{
    public class LawCloseReportModel
    {
        //序號
        public string SN { get; set; }

        /// <summary>照會單號</summary>
        public string LawNoteNo { get; set; }

        /// <summary>團隊名稱</summary>
        public string VmName { get; set; }

        /// <summary>體系名稱</summary>
        public string SmName { get; set; }

        /// <summary>處名稱</summary>
        public string CenterName { get; set; }

        /// <summary>實駐名稱</summary>
        public string WcCenterName { get; set; }

        /// <summary>業務員姓名</summary>
        public string LawDueName { get; set; }

        /// <summary>業務員id</summary>
        public string LawDueAgentId { get; set; }

        /// <summary>工作年月</summary>
        public string ProductionYm { get; set; }

        /// <summary>結欠金額</summary>
        public decimal LawDueMoney { get; set; }

        /// <summary>清償本金</summary>
        public int LawRepaymentCapital { get; set; }

        /// <summary>第一次電催內容</summary>
        public string LawPhoneCall1Desc { get; set; }

        /// <summary>第二次電催內容</summary>
        public string LawPhoneCall2Desc { get; set; }

        /// <summary>存證信函備註</summary>
        public string LawEvidencedesc { get; set; }

        public string LawLitigationProgress { get; set; }

        public string LawDoProgress { get; set; }

        /// <summary>承辦單位名稱</summary>
        public string LawDoUnitName { get; set; }
    }
}
