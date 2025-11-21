using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.MerSal.Models
{
    public class MerSalCutReportModel
    {
        public string IDD { get; set; }
        public string ProductionYM { get; set; }
        public string Sequence { get; set; }

        public string CompanyCode { get; set; }


        public string AgentCode1 { get; set; }
        public string Names1 { get; set; }
        public string AgentCode2 { get; set; }
        public string Names2 { get; set; }

        public string ReasonCode { get; set; }

        public string PolicyNo2 { get; set; }


        public string PlanCode { get; set; }
        public string CollectYear { get; set; }
        public string ModxSequence { get; set; }


        public string InsuredName { get; set; }
        public string Age { get; set; }
        public string PoIssueDate { get; set; }


        public string ModePrem { get; set; }
        public string CommPremC { get; set; }


        public string PayMonth { get; set; }
        public string PaySeq { get; set; }
        public string CheckIndName { get; set; }



    }
}
