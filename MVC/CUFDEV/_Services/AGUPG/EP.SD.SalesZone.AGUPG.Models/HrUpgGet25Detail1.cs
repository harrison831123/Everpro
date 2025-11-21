using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace EP.SD.SalesZone.AGUPG.Models
{
    /// <summary>
    /// 轄下第一代區級
    /// </summary>
    public class HrUpgGet25Detail1 : IModel
    {
        /// <summary>
        /// 年度
        /// </summary>
        [DisplayName("年度")]
        [Column("YYYY")]
        public string YYYY { get; set; }

        /// <summary>
        /// 季別
        /// </summary>
        [DisplayName("season")]
        [Column("season")]
        public string Season { get; set; }

        /// <summary>
        /// 業績區間
        /// </summary>
        [DisplayName("FYCrange")]
        [Column("FYCrange")]
        public string FYCrange { get; set; }

        /// <summary>
        /// 輔導人員代號
        /// </summary>
        [DisplayName("director_id")]
        [Column("director_id")]
        public string DirectorId { get; set; }

        /// <summary>
        /// 業務員代號
        /// </summary>
        [DisplayName("agent_code")]
        [Column("agent_code")]
        public string AgentCode { get; set; }

        /// <summary>
        /// 業務員姓名
        /// </summary>
        [DisplayName("agent_name")]
        [Column("agent_name")]
        public string AgentName { get; set; }

        /// <summary>
        /// 本季第1個月
        /// </summary>
        [DisplayName("本季第1個月")]
        [Column("pfyc_last1")]
        public string PfycLast1 { get; set; }

        /// <summary>
        /// 本季第2個月
        /// </summary>
        [DisplayName("本季第1個月")]
        [Column("pfyc_last2")]
        public string PfycLast2 { get; set; }

        /// <summary>
        /// 本季第3個月
        /// </summary>
        [DisplayName("本季第1個月")]
        [Column("pfyc_last3")]
        public string PfycLast3 { get; set; }

        /// <summary>
        /// 合計
        /// </summary>
        [DisplayName("合計")]
        [Column("pFyc_S1")]
        public string PFycS1 { get; set; }

        /// <summary>
        /// 符合條件
        /// </summary>
        [DisplayName("符合條件")]
        [Column("RightEmpOM01")]
        public string RightEmpOM01 { get; set; }

    }
}
