using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EP.SD.SalesSupport.CUSCRM.Web
{
    public class ProcessFormViewModel
    {
        /// <summary>
        /// 受理編號
        /// </summary>
        [Display(Name = "受理編號")]
        public string No { get; set; }

        /// <summary>
        /// 承辦人
        /// </summary>
        [Display(Name = "承辦人")]
        public string CaseHandler { get; set; }

        /// <summary>
        /// 受理時間
        /// </summary>
        [Display(Name = "受理時間")]
        public string CreatTime { get; set; }

        /// <summary>
        /// 來源
        /// </summary>
        [Display(Name = "來源")]
        public string Source { get; set; }

        /// <summary>
        /// 要保人
        /// </summary>
        [Display(Name = "要保人")]
        public string Owner { get; set; }

        /// <summary>
        /// 保單基本資料清單
        /// </summary>
        [Display(Name = "保單基本資料清單")]
        public string PolicyHistory { get; set; }

        /// <summary>
        /// 經手人資訊
        /// </summary>
        [Display(Name = "經手人資訊")]
        public string AgentHistory { get; set; }

        /// <summary>
        /// 保戶申訴及服務內容
        /// </summary>
        [Display(Name = "保戶申訴及服務內容")]
        public string ContentTXT { get; set; }

        /// <summary>
        /// 聯絡紀錄
        /// </summary>
        [Display(Name = "聯絡紀錄")]
        public string ContactRecord { get; set; }

        /// <summary>
        /// 是否寄發存證信函
        /// </summary>
        [Display(Name = "是否寄發存證信函")]
        public string DepositLetter { get; set; }

        /// <summary>
        /// 存證信函寄發日期
        /// </summary>
        [Display(Name = "存證信函寄發日期")]
        public string depositLetterDateTime { get; set; }

        /// <summary>
        /// 處理過程
        /// </summary>
        [Display(Name = "處理過程")]
        public string ProcessRecord { get; set; }

        /// <summary>
        /// 結案紀錄
        /// </summary>
        [Display(Name = "結案紀錄")]
        public string CloseRecord { get; set; }
    }
}