using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    public class MaintainInfo
    {
        /// <summary>
        /// 序號
        /// </summary>
        [Display(Name = "序號")]
        public int Serial { get; set; }

        /// <summary>
        /// 類別
        /// </summary>
        [Display(Name = "類別")]
        public string Type { get; set; }

        /// <summary>
        /// 類型
        /// </summary>
        [Display(Name = "類型")]
        public string Kind { get; set; }

        /// <summary>
        /// 照會單號(受理編號)
        /// </summary>
        [Display(Name = "照會單號")]
        public string No { get; set; }

        /// <summary>
        /// 受理日
        /// </summary>
        [Display(Name = "受理日")]
        public string CreateTime { get; set; }

        /// <summary>
        /// 要保人
        /// </summary>
        [Display(Name = "要保人")]
        public string Owner { get; set; }

        /// <summary>
        /// 業務員
        /// </summary>
        [Display(Name = "業務員")]
        public string AgentName { get; set; }

        /// <summary>
        /// 維護紀錄
        /// </summary>
        [Display(Name = "維護紀錄")]
        public string MtnRecord { get; set; }

        /// <summary>
        /// 催辦紀錄
        /// </summary>
        [Display(Name = "催辦紀錄")]
        public string DoSCount { get; set; }

        /// <summary>
        /// 稽催紀錄
        /// </summary>
        [Display(Name = "稽催紀錄")]
        public string AuditCount { get; set; }

        /// <summary>
        /// 狀態
        /// </summary>
        [Display(Name = "狀態")]
        public string Status { get; set; }
    }
}
