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
    /// 客服業務系統歷史查詢條件
    /// </summary>
    [DataContract]
    public class HistoryQueryCondition
    {
        /// <summary>
        /// 申訴類型(---------待改)
        /// </summary>
        [DataMember]
        [Display(Name = "申訴類型")]
        public int crmtype { get; set; }
        /// <summary>
        /// 結案狀態(---------待改)
        /// </summary>
        [DataMember]
        [Display(Name = "結案狀態")] 
        public int close_status { get; set; }
        /// <summary>
        /// 受理區間(起日)
        /// </summary>
        [DataMember]
        [Display(Name = "受理區間(起日)")] 
        public string startdate { get; set; }
        /// <summary>
        /// 受理區間(迄日)
        /// </summary>
        [DataMember]
        [Display(Name = "受理區間(迄日)")] 
        public string closedate { get; set; }
        /// <summary>
        /// 保單生效年度(民國)
        /// </summary>
        [DataMember]
        [Display(Name = "保單生效年度")] 
        public string issdate { get; set; }
        /// <summary>
        /// 保險公司
        /// </summary>
        [DataMember]
        [Display(Name = "保險公司")] 
        public string company_code { get; set; }

        /// <summary>
        /// 申訴、服務來源(---------待改)
        /// </summary>

        [DataMember]
        [Display(Name = "申訴、服務來源")] 
        public string source { get; set; }
        /// <summary>
        /// 受理類別
        /// </summary>
        [DataMember]
        [Display(Name = "受理類別")] 
        public string do_type { get; set; }

        /// <summary>
        /// 受理編號
        /// </summary>
        [DataMember]
        [Display(Name = "受理編號")] 
        public string crm_no { get; set; }
        /// <summary>
        /// 保單號碼
        /// </summary>
        [DataMember]
        [Display(Name = "保單號碼")] 
        public string policy_no2 { get; set; }


        /// <summary>
        /// 保戶姓名
        /// </summary>
        [DataMember]
        [Display(Name = "保戶姓名")] 
        public string owner { get; set; }
        /// <summary>
        /// 業務員姓名
        /// </summary>
        [DataMember]
        [Display(Name = "業務員姓名")] 
        public string agent_name { get; set; }
        /// <summary>
        /// 處經理姓名
        /// </summary>
        [DataMember]
        [Display(Name = "處經理姓名")] 
        public string center_manager { get; set; }
        /// <summary>
        /// 承辦人員
        /// </summary>
        [DataMember]
        [Display(Name = "承辦人員")] 
        public string worker { get; set; }

    }
}
