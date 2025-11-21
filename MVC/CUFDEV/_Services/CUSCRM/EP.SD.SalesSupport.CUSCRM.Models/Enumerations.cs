using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM
{

    /// <summary>
    /// 通知對象類別
    /// </summary>
    public enum NotifyType
    { 
        /// <summary>受文者</summary>
        To,
        /// <summary>副本受文者</summary>
        CC,
        /// <summary>行專</summary>
        Employee,
        /// <summary>其他</summary>
        Other,
        /// <summary>處代理人</summary>
        OMProxy
    }


    /// <summary>
    /// 案件狀態
    /// </summary>
    public enum ContentStatus
    {
        /// <summary>受理完成待通知</summary>
        WaitNotice,
        /// <summary>已通知</summary>
        Notified,
        /// <summary>處理中</summary>
        Process,
        /// <summary>受理不通知</summary>
        NoNotice,
        /// <summary>結案</summary>
        Close
    }

    /// <summary>
    /// 客服申訴類別
    /// </summary>
    public enum Category
    {
        /// <summary>客服</summary>
        [Display(Name = "客服")]
        Service,

        /// <summary>申訴</summary>
        [Display(Name = "申訴")]
        Complain
    }

    /// <summary>
    /// 啟用的狀態
    /// </summary>
    public enum EnableStatus
    {
        /// <summary>啟用</summary>
        [Display(Name = "啟用")]
        Enabled  = 1,
        /// <summary>停用</summary>
        [Display(Name = "停用")]
        Disabled = 0
    }

    /// <summary>
    /// 資料設定類別
    /// </summary>
    public enum DiscipTypeKind
    {
        /// <summary>服務申訴類型</summary>
        [Display(Name = "服務申訴類型")]
        Type = 1,

        /// <summary>案件類型</summary>
        [Display(Name = "案件類型")]
        CaseCategory = 2,

        /// <summary>案件類別</summary>
        [Display(Name = "案件類別")]
        CaseType = 3,

        /// <summary>資料來源</summary>
        [Display(Name = "資料來源")]
        Source = 4,

        /// <summary>結案狀態</summary>
        [Display(Name = "結案狀態")]
        CloseStatus = 5,

        /// <summary>來電者</summary>
        [Display(Name = "來電者")]
        Caller = 6,

        /// <summary>
        /// 處理結果單選
        /// </summary>
        [Display(Name ="處理結果單選")]
        DiscipProcResultRadio = 7,

        /// <summary>
        /// 處理結果複選
        /// </summary>
        [Display(Name = "處理結果複選")]
        DiscipProcResultCheckBox = 8,

        /// <summary>
        /// 保公決議單選
        /// </summary>
        [Display(Name = "保公決議單選")]
        DiscipCompanyResultRadio = 9
    }

    /// <summary>
    /// 資料設定代碼
    /// </summary>
    public enum DiscipTypeCode
    {
        /// <summary>系統</summary>
        [Display(Name = "SYS-系統")]
        [Value("SYS")]
        SYS,

        /// <summary>客戶服務</summary>
        [Display(Name = "CS-客戶服務")]
        [Value("CS")]
        CS,

        /// <summary>業務員服務</summary>
        [Display(Name = "SS-業務員服務")]
        [Value("SS")]
        SS,

        /// <summary>客戶申訴</summary>
        [Display(Name = "CC-客戶申訴")]
        [Value("CC")]
        CC,

        /// <summary>業務員申訴</summary>
        [Display(Name = "SC-業務員申訴")]
        [Value("SC")]
        SC
    }

    /// <summary>
    /// 維護狀態查詢類別
    /// </summary>
    public enum StatusCategory
    {
        /// <summary>全部</summary>
        [Display(Name = "全部")]
        All,

        /// <summary>照會中</summary>
        [Display(Name = "照會中")]
        Note,

        /// <summary>催辦中</summary>
        [Display(Name = "催辦中")]
        Dos,

        /// <summary>稽催中</summary>
        [Display(Name = "稽催中")]
        Audit,

        /// <summary>結案</summary>
        [Display(Name = "結案")]
        Close,

        /// <summary>立案不通知</summary>
        [Display(Name = "立案不通知")]
        Process
    }
}
