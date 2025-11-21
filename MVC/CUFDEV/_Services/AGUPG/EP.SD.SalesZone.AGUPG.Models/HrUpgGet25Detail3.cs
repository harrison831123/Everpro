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
    /// 加計被推介人業績
    /// </summary>
    public class HrUpgGet25Detail3 : IModel
    {
        /// <summary>
        /// 年度
        /// </summary>
        [DisplayName("年度")]
        [Column("YYYY")]
        public string YYYY { get; set; }

        /// <summary>
        /// 報表類型
        /// </summary>
        [DisplayName("報表類型")]
        [Column("RptType")]
        public string RptType { get; set; }

        /// <summary>
        /// 通訊處
        /// </summary>
        [DisplayName("通訊處")]
        [Column("wc_center_name")]
        public string WcCenterName { get; set; }

        /// <summary>
        /// 處單位
        /// </summary>
        [DisplayName("處單位")]
        [Column("center_name")]
        public string CenterName { get; set; }

        /// <summary>
        /// 最初推介人代號
        /// </summary>
        [DisplayName("orig_introducer_id")]
        [Column("orig_introducer_id")]
        public string OrigIntroducerId { get; set; }

        /// <summary>
        /// 最初推介人姓名
        /// </summary>
        [DisplayName("最初推介人姓名")]
        [Column("orig_introducer_name")]
        public string OrigIntroducerName { get; set; }

        /// <summary>
        /// 被推介人代號
        /// </summary>
        [DisplayName("被推介人代號")]
        [Column("agent_code")]
        public string AgentCode { get; set; }

        /// <summary>
        /// 被推介人姓名
        /// </summary>
        [DisplayName("被推介人姓名")]
        [Column("ag_name")]
        public string AgName { get; set; }

        /// <summary>
        /// 最新關檔月份稱謂代號
        /// </summary>
        [DisplayName("最新關檔月份稱謂代號")]
        [Column("ag_level")]
        public string AgLevel { get; set; }

        /// <summary>
        /// 異地推介
        /// </summary>
        [DisplayName("異地推介")]
        [Column("isIntroduceRemote")]
        public string IsIntroduceRemote { get; set; }

        /// <summary>
        /// 符合加計條件
        /// </summary>
        [DisplayName("符合加計條件")]
        [Column("PlusStatus")]
        public string PlusStatus { get; set; }

        /// <summary>
        /// 一月
        /// </summary>
        [DisplayName("一月")]
        [Column("Month1")]
        public string Month1 { get; set; }

        /// <summary>
        /// 二月
        /// </summary>
        [DisplayName("二月")]
        [Column("Month2")]
        public string Month2 { get; set; }

        /// <summary>
        /// 三月
        /// </summary>
        [DisplayName("三月")]
        [Column("Month3")]
        public string Month3 { get; set; }

        /// <summary>
        /// 四月
        /// </summary>
        [DisplayName("四月")]
        [Column("Month4")]
        public string Month4 { get; set; }

        /// <summary>
        /// 五月
        /// </summary>
        [DisplayName("五月")]
        [Column("Month5")]
        public string Month5 { get; set; }

        /// <summary>
        /// 六月
        /// </summary>
        [DisplayName("六月")]
        [Column("Month6")]
        public string Month6 { get; set; }

        /// <summary>
        /// 七月
        /// </summary>
        [DisplayName("七月")]
        [Column("Month7")]
        public string Month7 { get; set; }

        /// <summary>
        /// 八月
        /// </summary>
        [DisplayName("八月")]
        [Column("Month8")]
        public string Month8 { get; set; }

        /// <summary>
        /// 九月
        /// </summary>
        [DisplayName("九月")]
        [Column("Month9")]
        public string Month9 { get; set; }

        /// <summary>
        /// 十月
        /// </summary>
        [DisplayName("十月")]
        [Column("Month10")]
        public string Month10 { get; set; }

        /// <summary>
        /// 十一月
        /// </summary>
        [DisplayName("十一月")]
        [Column("Month11")]
        public string Month11 { get; set; }

        /// <summary>
        /// 十二月
        /// </summary>
        [DisplayName("十二月")]
        [Column("Month12")]
        public string Month12 { get; set; }
    }
}
