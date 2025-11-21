using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.Collections.PlanSet.Models
{
    /// <summary>
    /// 需求單號：20241210003 頁面呈現的title資訊 2024.12 BY VITA
    /// </summary>
    public class PlanTitle : IModel
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
    }
}
