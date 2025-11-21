using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.MerSal.Models
{
	public class MerSalCheckSPruleReportModel
	{

        public string ProductionYmS { get; set; }
        public string ProductionYmE { get; set; }
        public string CompanyCode { get; set; }
        public string FileSeq { get; set; }
        public string AmountTypeName { get; set; }
        public string ChkCodeName { get; set; }
        public string ActName { get; set; }
        public string Rule01 { get; set; }
        public string Rule02 { get; set; }
        public string Rule03 { get; set; }
        public string Rule04 { get; set; }
        public string Remark { get; set; }
        public string CreateDatetime { get; set; }
        //public string CreateUserCode { get; set; }
        public string CreateUserName { get; set; }
        public string UpdateDatetime { get; set; }
        //public string UpdateUserCode { get; set; }
        public string UpdateUserName { get; set; }
        public string IsDelete { get; set; }

    }
}
