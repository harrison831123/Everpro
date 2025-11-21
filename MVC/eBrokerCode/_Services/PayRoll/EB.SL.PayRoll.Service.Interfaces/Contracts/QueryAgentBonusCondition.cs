using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.PayRoll.Service.Contracts
{
    /// <summary>
    /// 查詢調整紀錄的條件
    /// </summary>
    public class QueryAgentBonusCondition
    {
        /// <summary>
        /// 業績年月
        /// </summary>
        public string ProductionYM { get; set; }

        /// <summary>
        /// 序號
        /// </summary>
        public short Sequence { get; set; }

        /// <summary>
        /// 調整類別
        /// </summary>
        public string AdjType { get; set; }

        /// <summary>
        /// 公司代碼
        /// </summary>
        public string CompanyCode { get; set; }

        /// <summary>
        /// 業務員代碼
        /// </summary>
        public string AgentCode { get; set; }

        /// <summary>
        /// 原因碼
        /// </summary>
        public string ReasonCode { get; set; }

        /// <summary>
        /// 說明內容
        /// </summary>
        public string DescContent { get; set; }

        /// <summary>
        /// 業務員姓名
        /// </summary>
        public string AgentName { get; set; }

        /// <summary>
        /// 保單號碼
        /// </summary>
        public string PolicyNo2 { get; set; }

        /// <summary>
        /// 金額
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// FYC
        /// </summary>
        public decimal FYC { get; set; }

        /// <summary>
        /// FYP
        /// </summary>
        public decimal FYP { get; set; }

        /// <summary>
        /// 險種代碼
        /// </summary>
        public string PlanCode { get; set; }

        /// <summary>
        /// 繳費年期
        /// </summary>
        public short CollectYear { get; set; }
        /// <summary>
        /// 繳別繳次
        /// </summary>
        public string ModxSequence { get; set; }
        /// <summary>
        /// 計佣保費
        /// </summary>
        public decimal CommModePrem { get; set; }
        /// <summary>
        /// 建立人員部門
        /// </summary>
        public string CreateUnit { get; set; }
        /// <summary>
        /// 員編
        /// </summary>
        public string memberid { get; set; }
        /// <summary>
        /// ID
        /// </summary>
        public string CreateUserCode { get; set; }

    }
}
