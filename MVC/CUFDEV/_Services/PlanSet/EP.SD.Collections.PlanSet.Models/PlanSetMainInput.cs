using Microsoft.CUF.Framework.Data;
using System.ComponentModel;
namespace EP.SD.Collections.PlanSet.Models
{
    /// <summary>
    /// 需求單號：20241210003 使用者輸入查詢條件 2024.12 BY VITA
    /// </summary>
    public class PlanSetMainInput : IModel
    {
        /// <summary>
        /// 保險公司
        /// </summary>
        [DisplayName("保險公司")]
        [Column("company_name")]
        public string company_name { get; set; }

        /// <summary>
        /// 險種代碼
        /// </summary>
        [DisplayName("險種代碼")]
        [Column("plan_code")]
        public string plan_code { get; set; }

        /// <summary>
        /// 年期
        /// </summary>
        [DisplayName("年期")]
        [Column("collect")]
        public string collect { get; set; }
    }
}