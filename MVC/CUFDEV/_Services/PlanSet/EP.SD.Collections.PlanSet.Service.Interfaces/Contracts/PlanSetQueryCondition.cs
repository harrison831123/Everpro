using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.Collections.PlanSet.Service
{
    /// <summary>
    /// 需求單號：20241210003 保險公司下拉式選單 2024.12 BY VITA
    /// </summary>
    [DataContract]   
    public class PlanSetQueryCondition
    {
        [DataMember]
        [Display(Name = "保險公司")]
        public string CompanyCode { get; set; }
    }
}