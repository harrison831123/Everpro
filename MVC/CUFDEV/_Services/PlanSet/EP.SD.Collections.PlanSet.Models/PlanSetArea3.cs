using Microsoft.CUF.Framework.Data;
using System.ComponentModel;

namespace EP.SD.Collections.PlanSet.Models
{
    /// <summary>
    /// 需求單號：20241210003 永達競賽獎勵(業支部) 2024.12 BY VITA
    /// </summary>
    public class PlanSetArea3 : IModel
    {
        /// <summary>
        ///年期
        /// </summary>
        [DisplayName("年期")]
        [Column("plan_year_cond", IsKey = true, IsIdentity = true)]
        public string plan_year_cond { get; set; }

        /// <summary>
        ///要保申請起日
        /// </summary>
        [DisplayName("要保申請起日")]
        [Column("set_start_date", IsKey = true, IsIdentity = true)]
        public string set_start_date { get; set; }

        /// <summary>
        ///要保申請迄日
        /// </summary>
        [DisplayName("要保申請迄日")]
        [Column("set_end_date", IsKey = true, IsIdentity = true)]
        public string set_end_date { get; set; }

        /// <summary>
        ///競賽FYC
        /// </summary>
        [DisplayName("競賽FYC")]
        [Column("IsCredited_txt", IsKey = true, IsIdentity = true)]
        public string IsCredited_txt { get; set; }
    }
}
