using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;


namespace EP.SD.SalesSupport.LAW.Models
{
    public class LawMonthRepaymentReportDetail : IModel
    {
        /// <summary>
        /// 流水號
        /// </summary>
        [DataMember]
        [DisplayName("流水號")]
        [Column("law_id", IsKey = true, IsIdentity = true)]
        public int LawId { get; set; }

        /// <summary>照會單號</summary>
        [DataMember]
        [DisplayName("照會單號")]
        [Column("law_note_no")]
        public string LawNoteNo { get; set; }

        /// <summary>業務員id</summary>
        [DataMember]
        [DisplayName("業務員id")]
        [Column("law_due_agentid")]
        public string LawDueAgentId { get; set; }

        /// <summary>清償金額</summary>
        [DataMember]
        [DisplayName("law_repayment_money")]
        [Column("law_repayment_money")]
        public int LawRepaymentMoney { get; set; }

        [DataMember]
        [Column("vm_code")]
        public string vmcode { get; set; }

        [DataMember]
        [Column("vm_name")]
        public string vmname { get; set; }

        [DataMember]
        [Column("sm_code")]
        public string smcode { get; set; }

        [DataMember]
        [Column("sm_name")]
        public string smname { get; set; }

        [DataMember]
        [Column("name")]
        public string name { get; set; }

        [NonColumn]
        public string chkm { get; set; }

        [NonColumn]
        public string selyear { get; set; }

        [NonColumn]
        public string selmonth { get; set; }

        [NonColumn]
        public int LawSumRepaymentMoney { get; set; }
        
    }
}
