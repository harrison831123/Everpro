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
    public class LawOtherDescDetail : IModel
    {

        /// <summary>
        /// 流水號
        /// </summary>
        [DataMember]
        [DisplayName("law_other_id")]
        [Column("law_other_id", IsKey = true, IsIdentity = true)]
        public int LawOtherId { get; set; }

        /// <summary>
        /// 主檔流水號
        /// </summary>
        [DataMember]
        [DisplayName("law_id")]
        [Column("law_id")]
        public int LawId { get; set; }

        /// <summary>照會單號</summary>
        [DataMember]
        [DisplayName("law_note_no")]
        [Column("law_note_no")]
        public string LawNoteNo { get; set; }

        /// <summary>
        /// 流水號
        /// </summary>
        [DataMember]
        [DisplayName("law_repayment_id")]
        [Column("law_repayment_id", IsKey = true, IsIdentity = true)]
        public int LawRepaymentId { get; set; }


        /// <summary>清償金額</summary>
        [DataMember]
        [DisplayName("law_repayment_money")]
        [Column("law_repayment_money")]
        public int LawRepaymentMoney { get; set; }

        /// <summary>業務員id</summary>
        [DataMember]
        [DisplayName("law_due_agentid")]
        [Column("law_due_agentid")]
        public string LawDueAgentId { get; set; }


        /// <summary>其他備註說明</summary>
        [DataMember]
        [DisplayName("law_other_desc")]
        [Column("law_other_desc")]
        public string LawOtherdesc { get; set; }

        /// <summary>建檔人員ID</summary>
        [DataMember]
        [DisplayName("OtherDescCreatorID")]
        [Column("OtherDescCreatorID")]
        public string OtherDescCreatorID { get; set; }

        /// <summary>建檔日期</summary>
        [DataMember]
        [DisplayName("create_date")]
        [Column("create_date")]
        public DateTime CreateDate { get; set; }

        /// <summary>續佣扣抵</summary>
        [DataMember]
        [DisplayName("law_comm_deduction")]
        [Column("law_comm_deduction")]
        public int LawCommDeduction { get; set; }

    }
}
