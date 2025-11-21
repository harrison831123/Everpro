using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;
using static EP.SD.SalesZone.AGUPG.Models.Enumerations;

namespace EP.SD.SalesZone.AGUPG.Service
{
    [DataContract]
    public class HrUpg25QueryCondition
    {
        /// <summary>
        /// 業務員ID
        /// </summary>
        [DataMember]
        public string AgentCode { get; set; }

        /// <summary>
        /// 業績區間 年度季度
        /// </summary>
        [DataMember]
        public string YYYYSeason { get; set; }

        /// <summary>
        /// 明細類別
        /// </summary>
        [DataMember]
        public string DetailType { get; set; }

        /// <summary>
        /// 身份別
        /// </summary>
        [DataMember]
        public AGUPGUserType UserType { get; set; }
    }
}
