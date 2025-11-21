using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;
using Microsoft.CUF.Framework;

/// <summary>
/// 202505 by Fion 20250527001_佣酬預估試算
/// </summary>
namespace EP.SD.Collections.PlanSet.Service
{
    [DataContract]
    public class AgentRewardPolicyCondition
    {
        [DataMember]
        [Display(Name = "業務員代碼")]
        public string AgentCode { get; set; }

        [DataMember]
        [Display(Name = "業務員姓名")]
        public string AgentName { get; set; }

        [DataMember]
        [Display(Name = "職級")]
        public string AgLevel { get; set; }

        [DataMember]
        [Display(Name = "職級名稱簡稱")]
        public string AgLevelNameBf { get; set; }

        [DataMember]
        [Display(Name = "職級名稱")]
        public string AgLevelName { get; set; }

        [DataMember]
        [Display(Name = "職級OccpInd")]
        public string AgOccpInd { get; set; }

        [DataMember]
        [Display(Name = "MDRT屆數")]
        public string MDRT { get; set; }

        [DataMember]
        [Display(Name = "試算年月")]
        public string CalYM { get; set; }

        [DataMember]
        [Display(Name = "目標FYC")]
        public int TargetFYC { get; set; }

        [DataMember]
        [Display(Name = "資料日期")]
        public string DataDate { get; set; }

        [DataMember]
        [Display(Name = "ViewMsg")]
        public string ViewMsg { get; set; }

        #region 20250901003_佣酬預估試算優化
        [DataMember]
        [Display(Name = "試算年月_起")]
        public string CalYmS { get; set; }

        [DataMember]
        [Display(Name = "試算年月_迄")]
        public string CalYmE { get; set; }

        [DataMember]
        [Display(Name = "業務員Id")]
        public string AgentId { get; set; }

        [DataMember]
        [Display(Name = "年度累計所得")]
        public int AgentTotalIncome { get; set; }
        #endregion #region 20250901003_佣酬預估試算優化
    }
}
