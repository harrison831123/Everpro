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
    /// 連續四季明細
    /// </summary>
    public class HrUpgGet25Detail2 : IModel
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
        [DisplayName("季別")]
        [Column("season")]
        public string Season { get; set; }

        /// <summary>
        /// 業績季別
        /// </summary>
        [DisplayName("業績季別")]
        [Column("Season1")]
        public string Season1 { get; set; }

        /// <summary>
        /// 業務員代號
        /// </summary>
        [DisplayName("業務員代號")]
        [Column("agent_code")]
        public string AgentCode { get; set; }

        /// <summary>
        /// 業務員姓名
        /// </summary>
        [DisplayName("業務員姓名")]
        [Column("agent_name")]
        public string AgentName { get; set; }

        /// <summary>
        /// 評估季別
        /// </summary>
        [DisplayName("評估季別")]
        [Column("UpgSeason")]
        public string UpgSeason { get; set; }

        /// <summary>
        /// 【條件一】
        /// 率先入圍當年度MDRT
        /// A標/B標
        /// </summary>
        [DisplayName("【條件一】")]
        [Column("AssRst_MdrtAB")]
        public string AssRstMdrtAB { get; set; }

        /// <summary>
        /// 【條件二】
        /// 前一季達FYC 20萬
        /// 當季視同通過評估
        /// </summary>
        [DisplayName("【條件二】")]
        [Column("AssRst_LSfycP20_A")]
        public string AssRstLSfycP20A { get; set; }

        /// <summary>
        /// 【條件三】
        ///  當季FYC合計
        ///  (已含加計及扣除申訴)
        /// </summary>
        [DisplayName("【條件三】")]
        [Column("AssRst_fycP_A")]
        public string AssRstFycPA { get; set; }

        /// <summary>
        /// 【條件四】
        ///  工作日誌達一日兩訪
        /// </summary>
        [DisplayName("【條件四】")]
        [Column("WorkNote")]
        public string WorkNote { get; set; }

        /// <summary>
        /// 核定說明
        /// </summary>
        [DisplayName("核定說明")]
        [Column("AssRstFinal")]
        public string AssRstFinal { get; set; }

    }
}
