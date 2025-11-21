using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    public class MtnHistoryInfo
    {
        /// <summary>
        /// 處理結果選單
        /// </summary>
        [Display(Name = "處理結果選單")]
        public List<CRMEDiscipType> ResultDiscipType { get; set; }

        /// <summary>
        /// 保公決議選單
        /// </summary>
        [Display(Name = "保公決議選單")]
        public List<CRMEDiscipType> CompanyDiscipType { get; set; }

        /// <summary>
        /// 受理日期
        /// </summary>
        [Display(Name = "受理日期")]
        public string CreateDate { get; set; }

        /// <summary>
        /// 歷次催辦日期
        /// </summary>
        [Display(Name = "歷次催辦日期")]
        public string HistoryDoSDate { get; set; }

        /// <summary>
        /// 歷次稽催日期
        /// </summary>
        [Display(Name = "歷次稽催日期")]
        public string HistoryAuditDate { get; set; }

        /// <summary>
        /// 歷次維護內容
        /// </summary>
        [Display(Name = "歷次維護內容")]
        public string HistoryContent{ get; set; }

        /// <summary>
        /// 歷次業連保險公司日期
        /// </summary>
        [Display(Name = "歷次業連保險公司日期")]
        public string HistoryBUSContactCompanyDate { get; set; }

        /// <summary>
        /// 歷次保險公司回覆日
        /// </summary>
        [Display(Name = "歷次保險公司回覆日")]
        public string HistoryReplyCompanyDate { get; set; }

        /// <summary>
        /// 歷次回覆單位黃聯日
        /// </summary>
        [Display(Name = "歷次回覆單位黃聯日")]
        public string HistoryReplyYallowBillDate { get; set; }

        /// <summary>
        /// 歷次客戶申訴件回文日
        /// </summary>
        [Display(Name = "歷次客戶申訴件回文日")]
        public string HistoryCCReplyDate { get; set; }

        /// <summary>
        /// 歷次結案處理結果
        /// </summary>
        [Display(Name = "歷次結案處理結果")]
        public string HistoryCloseLog { get; set; }

        /// <summary>
        /// 維護附件
        /// </summary>
        [Display(Name = "維護附件")]
        public List<CRMEFile> Files { get; set; }

        /// <summary>
        /// 保公決議選項
        /// </summary>
        [Display(Name = "保公決議選項")]
        public string HistoryCompanyResult { get; set; }

    }
}
