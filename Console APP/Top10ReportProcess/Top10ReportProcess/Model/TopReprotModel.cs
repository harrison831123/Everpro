//20250326001-保險公司受理前十大商品年報排程 20250507 by Harrison
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Top10ReportProcess.Model
{
    public class TopReprotModel
    {
        /// <summary>
        /// 排名
        /// </summary>
        public string Rank { get; set; }

        /// <summary>
        /// 保險公司
        /// </summary>
        public string Company { get; set; }

        /// <summary>
        /// 險種
        /// </summary>
        public string PlanTitle { get; set; }

        /// <summary>
        /// 險種代碼
        /// </summary>
        public string PlanCode { get; set; }

        /// <summary>
        /// 傭金
        /// </summary>
        public string FYC { get; set; }

        /// <summary>
        /// 繳別保費
        /// </summary>
        public string FYP { get; set; }

        /// <summary>
        /// 件數
        /// </summary>
        public string Cnt { get; set; }

        /// <summary>
        /// 比例
        /// </summary>
        public string Proportion { get; set; }

    }
}
