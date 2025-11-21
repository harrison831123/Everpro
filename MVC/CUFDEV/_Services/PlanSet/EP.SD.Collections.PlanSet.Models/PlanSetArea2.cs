using Microsoft.CUF.Framework.Data;
using System.ComponentModel;

namespace EP.SD.Collections.PlanSet.Models
{
    /// <summary>
    /// 需求單號：20241210003 保公獎勵內容(業行部) 2024.12 BY VITA
    /// </summary>
    public class PlanSetArea2 : IModel
    {
        /// <summary>
        ///年期
        /// </summary>
        [DisplayName("年期")]
        [Column("plan_year_cond", IsKey = true, IsIdentity = true)]
        public string plan_year_cond { get; set; }

        /// <summary>
        ///獎勵生效起日
        /// </summary>
        [DisplayName("獎勵生效起日")]
        [Column("reward_start_date", IsKey = true, IsIdentity = true)]
        public string reward_start_date { get; set; }

        /// <summary>
        ///獎勵生效迄日
        /// </summary>
        [DisplayName("獎勵生效迄日")]
        [Column("reward_end_date", IsKey = true, IsIdentity = true)]
        public string reward_end_date { get; set; }

        /// <summary>
        /// 年齡
        /// </summary>
        [DisplayName("年齡")]
        [Column("age_cond", IsKey = true, IsIdentity = true)]
        public string age_cond { get; set; }

        /// <summary>
        /// 其他條件
        /// </summary>
        [DisplayName("其他條件")]
        [Column("other_cond", IsKey = true, IsIdentity = true)]
        public string other_cond { get; set; }

        /// <summary>
        ///首年首次FYC*獎金率
        /// </summary>
        [DisplayName("首年首次FYC*獎金率")]
        [Column("rate_value_txt", IsKey = true, IsIdentity = true)]
        public string rate_value_txt { get; set; }
    }
}