using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    /// <summary>
    /// 客服業務系統報表查詢條件
    /// </summary>
    [DataContract]
    public class QueryReportCondition
    {
        /// <summary>
        /// 類別
        /// </summary>
        [DataMember]
        [Display(Name = "類別")]
        public Category category { get; set; }
        /// <summary>
        /// 類型
        /// </summary>
        [DataMember]
        [Display(Name = "類型")]
        public string code { get; set; }
        /// <summary>
        /// 保單號碼
        /// </summary>
        [DataMember]
        [Display(Name = "保單號碼")]
        public string PolicyNo { get; set; }
        /// <summary>
        /// 要保人
        /// </summary>
        [DataMember]
        [Display(Name = "要保人姓名")]
        public string owner { get; set; }
        /// <summary>
        /// 照會單號
        /// </summary>
        [DataMember]
        [Display(Name = "照會單號")]
        public string crm_no { get; set; }
        /// <summary>
        /// 受理區間(起日)
        /// </summary>
        [DataMember]
        [Display(Name = "受理區間(起日)")]
        public DateTime? startdate { get; set; }
        /// <summary>
        /// 受理區間(迄日)
        /// </summary>
        [DataMember]
        [Display(Name = "受理區間(迄日)")]
        public DateTime? closedate { get; set; }
        /// <summary>
        /// 承辦人員
        /// </summary>
        [DataMember]
        [Display(Name = "承辦人員")]
        public string worker { get; set; }
        /// <summary>
        /// 結案狀態
        /// </summary>
        [DataMember]
        [Display(Name = "結案狀態")]
        public int close_status { get; set; }
        
    }
}