using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesZone.AGUPG.Models
{
    public class HrUpg25RstViewModel : IModel
    {
        /// <summary>
        /// 業務員代碼
        /// </summary>
        public string AgentCode { get; set; }

        /// <summary>
        /// 職等
        /// </summary>
        public string AgLevel { get; set; }

        /// <summary>
        /// 權限判斷
        /// </summary>
        public bool IsAdmin { get; set; }

        /// <summary>
        /// 年度季度
        /// </summary>
        public string YYYYSeason { get; set; }

        public List<HrUpg25Item> hrUpg25List { get; set; }

        public class HrUpg25Item
        {
            /// <summary>
            /// 下拉選單文字
            /// </summary>
            [DisplayName("NowData")]
            [Column("NowData")]
            public string NowData { get; set; }

            /// <summary>
            /// 年度季度
            /// </summary>
            [DisplayName("YYYYSeason")]
            [Column("YYYYSeason")]
            public string YYYYSeason { get; set; }
        }
    }
}
