using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


//需求單號：20240715004 新增佣酬發放調整報表欄位，生效日、年齡、招攬人姓名及ID等欄位 vicky

namespace EB.SL.MerSal.Models
{
    public class MerSalCheckReportModel
    {
        public string ProductionYM { get; set; }

        public short Sequence { get; set; }

        public string CompanyCode { get; set; }

        public string FileSeq { get; set; }

        public string AmountType { get; set; }

        public string PolicyNo2 { get; set; }

        public string PlanCode { get; set; }

        public short CollectYear { get; set; }

        public string ModxSequence { get; set; }

        public string PoIssueDate { get; set; }
        public string Age { get; set; }
        public string ReasonCodeC { get; set; }
        public string ModePrem { get; set; }
        public string CommPremC { get; set; }
        public string Amount { get; set; }
        public string CommPrem { get; set; }
        public string AgentName1 { get; set; }
        public string AgentCode1 { get; set; }
        public string AgentName2 { get; set; }
        public string AgentCode2 { get; set; }
        public string ReceiptDate { get; set; }
        public string PayYm { get; set; }

        //public string PayType { get; set; }
        public string PayTypeName { get; set; }
        public string RptInclideFlag { get; set; }
        public string AgCnt { get; set; }
        public string PolicyStatusName { get; set; }
        public string NotPayYMS { get; set; }
        public string Memo { get; set; }
        public string Remark { get; set; }





    }
}
