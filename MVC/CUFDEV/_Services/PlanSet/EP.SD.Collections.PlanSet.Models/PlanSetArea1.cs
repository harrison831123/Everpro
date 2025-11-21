using Microsoft.CUF.Framework.Data;
using System.ComponentModel;

namespace EP.SD.Collections.PlanSet.Models
{
    /// <summary>
    /// 需求單號：20241210003 初年度業績換算率(業行部) 2024.12 BY VITA
    /// </summary>
    public class PlanSetArea1 : IModel
    {
        /// <summary>
        ///保險公司
        /// </summary>
        [DisplayName("保險公司")]
        [Column("company_name", IsKey = true, IsIdentity = true)]
        public string company_name { get; set; }

        /// <summary>
        ///險種名稱
        /// </summary>
        [DisplayName("險種名稱")]
        [Column("plan_title", IsKey = true, IsIdentity = true)]
        public string plan_title { get; set; }

        /// <summary>
        ///險種代碼
        /// </summary>
        [DisplayName("險種代碼")]
        [Column("plan_code", IsKey = true, IsIdentity = true)]
        public string plan_code { get; set; }

        /// <summary>
        ///年期
        /// </summary>
        [DisplayName("年期")]
        [Column("plan_year_cond", IsKey = true, IsIdentity = true)]
        public string plan_year_cond { get; set; }

        /// <summary>
        ///起賣日
        /// </summary>
        [DisplayName("起賣日")]
        [Column("plan_start_date", IsKey = true, IsIdentity = true)]
        public string plan_start_date{ get; set; }

        /// <summary>
        ///停賣日
        /// </summary>
        [DisplayName("停賣日")]
        [Column("plan_end_date", IsKey = true, IsIdentity = true)]
        public string plan_end_date { get; set; }

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
        ///初年度業績換算率
        /// </summary>
        [DisplayName("初年度業績換算率")]
        [Column("rate_value_txt", IsKey = true, IsIdentity = true)]
        public string rate_value_txt { get; set; }
    }
}