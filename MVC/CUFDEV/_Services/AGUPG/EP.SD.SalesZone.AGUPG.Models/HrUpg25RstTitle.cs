using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesZone.AGUPG.Models
{
    public class HrUpg25RstTitle : IModel
    {
        /// <summary>
        /// 下拉選單顯示文字
        /// </summary>
        [DisplayName("下拉選單顯示文字")]
        [Column("now_data")]
        public string NowData { get; set; }

        /// <summary>
        /// 業務員代碼
        /// </summary>
        [DisplayName("業務員代碼")]
        [Column("agent_code")]
        public string AgentCode { get; set; }

        /// <summary>
        /// 業務員姓名
        /// </summary>
        [DisplayName("業務員姓名")]
        [Column("agent_name")]
        public string AgentName { get; set; }

        /// <summary>
        /// 職等名稱
        /// </summary>
        [DisplayName("職等名稱")]
        [Column("agLevelOccpName")]
        public string AgLevelOccpName { get; set; }

        /// <summary>
        /// 簽約日期
        /// </summary>
        [DisplayName("簽約日期")]
        [Column("register_date")]
        public string RegisterDate { get; set; }

        /// <summary>
        /// 原始登錄日
        /// </summary>
        [DisplayName("原始登錄日")]
        [Column("min_record_date")]
        public string MinRecordDate { get; set; }

        /// <summary>
        /// 第一代輔導人代碼
        /// </summary>
        [DisplayName("第一代輔導人代碼")]
        [Column("director_id")]
        public string DirectorId { get; set; }

        /// <summary>
        /// 第一代輔導人姓名
        /// </summary>
        [DisplayName("第一代輔導人姓名")]
        [Column("director_name")]
        public string DirectorName { get; set; }

        /// <summary>
        /// 處單位
        /// </summary>
        [DisplayName("處單位")]
        [Column("center_name")]
        public string CenterName { get; set; }

        /// <summary>
        /// 通訊處
        /// </summary>
        [DisplayName("通訊處")]
        [Column("wc_center_name")]
        public string WcCenterName { get; set; }

        /// <summary>
        /// 加油
        /// </summary>
        [DisplayName("加油")]
        [Column("Rst_UpgradeMemo")]
        public string RstUpgradeMemo { get; set; }

        /// <summary>
        /// 適用晉級工作月
        /// </summary>
        [DisplayName("適用晉級工作月")]
        [Column("UpgMonth")]
        public string UpgMonth { get; set; }

        /// <summary>
        /// 申請晉級日期
        /// </summary>
        [DisplayName("申請晉級日期")]
        [Column("UpgApplyDate")]
        public string UpgApplyDate { get; set; }

        /// <summary>
        /// 查詢日期
        /// </summary>
        [DisplayName("查詢日期")]
        [NonColumn]
        public string QueryDate { get; set; }

   
    }
}
