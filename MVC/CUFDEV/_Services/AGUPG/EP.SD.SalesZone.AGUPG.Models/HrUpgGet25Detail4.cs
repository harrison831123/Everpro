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
    /// 遞延未核實
    /// </summary>
    public class HrUpgGet25Detail4 : IModel
    {
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
        /// half件
        /// </summary>
        [DisplayName("half件")]
        [Column("isHalf")]
        public string IsHalf { get; set; }

        /// <summary>
        /// 受理月份
        /// </summary>
        [DisplayName("受理月份")]
        [Column("PolicyYM")]
        public string PolicyYM { get; set; }

        /// <summary>
        /// 起保日
        /// </summary>
        [DisplayName("起保日")]
        [Column("po_issue_date")]
        public string PoIssueDate { get; set; }

        /// <summary>
        /// 保險公司名稱
        /// </summary>
        [DisplayName("保險公司名稱")]
        [Column("company_name")]
        public string CompanyName { get; set; }

        /// <summary>
        /// 保單號碼
        /// </summary>
        [DisplayName("保單號碼")]
        [Column("policy_no2")]
        public string PolicyNo2 { get; set; }

        /// <summary>
        /// 被保險人姓名
        /// </summary>
        [DisplayName("被保險人姓名")]
        [Column("InsuredName")]
        public string InsuredName { get; set; }

        /// <summary>
        /// 要保人姓名
        /// </summary>
        [DisplayName("要保人姓名")]
        [Column("ProposerName")]
        public string ProposerName { get; set; }

        /// <summary>
        /// 繳別
        /// </summary>
        [DisplayName("繳別")]
        [Column("modx_name")]
        public string ModxName { get; set; }

        /// <summary>
        /// 險種代號
        /// </summary>
        [DisplayName("險種代號")]
        [Column("plan_code")]
        public string PlanCode { get; set; }

        /// <summary>
        /// 繳費方式
        /// </summary>
        [DisplayName("繳費方式")]
        [Column("method_name")]
        public string MethodName { get; set; }

        /// <summary>
        /// 業務員簽收日
        /// </summary>
        [DisplayName("業務員簽收日")]
        [Column("AGSignDate")]
        public string AGSignDate { get; set; }

        /// <summary>
        /// 助理簽收日
        /// </summary>
        [DisplayName("助理簽收日")]
        [Column("AssistSignDate")]
        public string AssistSignDate { get; set; }

        /// <summary>
        /// 保單狀態
        /// </summary>
        [DisplayName("保單狀態")]
        [Column("PolicyStatusName")]
        public string PolicyStatusName { get; set; }

        /// <summary>
        /// 台幣總保費
        /// </summary>
        [DisplayName("台幣總保費")]
        [Column("NTTotalPREM")]
        public string NTTotalPREM { get; set; }

        /// <summary>
        /// 預估 FYC
        /// </summary>
        [DisplayName("預估 FYC")]
        [Column("EstimateFYC")]
        public string EstimateFYC { get; set; }

    }
}
