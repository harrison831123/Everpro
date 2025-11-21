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
    [DataContract]
    [Table("law_vm_sm_report")]
    public class LawVmSmDetail : IModel
    {
        /// <summary>年</summary>
        [DataMember]
        [Column("law_year")]
        public string LawYear { get; set; }

        /// <summary>團隊</summary>
        [DataMember]
        [Column("vm_name")]
        public string VmName { get; set; }

        /// <summary>體系</summary>
        [DataMember]
        [Column("sm_name")]
        public string SmName { get; set; }

        /// <summary>金額</summary>
        [DataMember]
        [Column("due_money")]
        public string DueMoney { get; set; }

        /// <summary>金額</summary>
        [DataMember]
        [Column("repay_money")]
        public string RepayMoney { get; set; }

        [NonColumn]
        public string pstr { get; set; }

        [NonColumn]
        public decimal dv { get; set; }
    }
}
