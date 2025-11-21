using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.PayRoll.Service.Contracts
{
    public class AgentBonusAdjustCondition : IModel
    {
        [DataMember]
        [Column("production_ym")]
        public string ProductionYM { get; set; }

        /// <summary>
        /// 序號
        /// </summary>
        [DataMember]
        [Column("sequence")]
        public short Sequence { get; set; }

        /// <summary>
        /// 資料建檔部門代碼
        /// </summary>    
        [DataMember]
        [Column("create_unit")]
        public string CreateUnit { get; set; }

        /// <summary>
        /// 資料建檔人員
        /// </summary>
        [NonColumn]
        public string CreateUserCode { get; set; }

        /// <summary>
        /// 建表人員
        /// </summary>
        [NonColumn]
        public string CreateReportUserName { get; set; }

        /// <summary>
        /// 建表人員部門
        /// </summary>
        [NonColumn]
        public string CreateReportUnitName { get; set; }

        /// <summary>
        /// 上傳者名稱
        /// </summary>
        [DataMember]
        [Column("nmember")]
        public string nmember { get; set; }
    }
}
