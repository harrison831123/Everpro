using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesZone.AGUPG.Models
{
    public class HrUpg25RstGrid2 : IModel
    {
        /// <summary>
        /// 流水號
        /// </summary>
        [DisplayName("iden")]
        [Column("iden")]
        public string Iden { get; set; }

        /// <summary>
        /// IDD
        /// </summary>
        [DisplayName("IDD")]
        [Column("IDD")]
        public string IDD { get; set; }

        /// <summary>
        /// 年度
        /// </summary>
        [DisplayName("YYYY")]
        [Column("YYYY")]
        public string YYYY { get; set; }

        /// <summary>
        /// 季度
        /// </summary>
        [DisplayName("Season")]
        [Column("Season")]
        public string Season { get; set; }

        /// <summary>
        /// 業務員代碼
        /// </summary>
        [DisplayName("agent_code")]
        [Column("agent_code")]
        public string AgentCode { get; set; }

        /// <summary>
        /// show_type
        /// </summary>
        [DisplayName("show_type")]
        [Column("show_type")]
        public string ShowType { get; set; }

        /// <summary>
        /// 空白
        /// </summary>
        [DisplayName("show_col1")]
        [Column("show_col1")]
        public string ShowCol1 { get; set; }

        /// <summary>
        /// 適用標準
        /// </summary>
        [DisplayName("show_col2")]
        [Column("show_col2")]
        public string ShowCol2 { get; set; }

        /// <summary>
        /// 我的數字
        /// </summary>
        [DisplayName("show_col3")]
        [Column("show_col3")]
        public string ShowCol3 { get; set; }

        /// <summary>
        /// 檢核結果(已達標、未達標)
        /// </summary>
        [DisplayName("show_col4")]
        [Column("show_col4")]
        public string ShowCol4 { get; set; }

        /// <summary>
        /// 標準值
        /// </summary>
        [DisplayName("show_col5")]
        [Column("show_col5")]
        public string ShowCol5 { get; set; }

        /// <summary>
        /// create_datetime
        /// </summary>
        [DisplayName("create_datetime")]
        [Column("create_datetime")]
        public DateTime CreateDatetime { get; set; }

    }
}
