using EP.SD.SalesSupport.LAW.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Model
{
    public class LawDetailModel
    {
        /// <summary>管理人員</summary>
        [Display(Name = "選擇管理人員", ResourceType = typeof(LAWResource))]
        public string SysName { get; set; }

        /// <summary>承辦人員</summary>
        [Display(Name = "選擇經辦人員", ResourceType = typeof(LAWResource))]
        public string DouserName { get; set; }

        /// <summary>分機號碼</summary>
        [Display(Name = "分機號碼", ResourceType = typeof(LAWResource))]
        public string DouserPhoneExt { get; set; }

        /// <summary>預設經辦順位</summary>
        [Display(Name = "預設經辦順位", ResourceType = typeof(LAWResource))]
        public int DouserSort { get; set; }

        /// <summary>承辦單位</summary>
        [Display(Name = "承辦單位", ResourceType = typeof(LAWResource))]
        public string UnitName { get; set; }

        /// <summary>啟用狀態</summary>
        //[Display(Name = "執行人員", ResourceType = typeof(LAWResource))]
        public int StatusType { get; set; }

        /// <summary>建立者姓名</summary>
        [Display(Name = "建立者姓名", ResourceType = typeof(LAWResource))]
        public string CreateName { get; set; }

        /// <summary>建立時間</summary>
        [Display(Name = "建立時間", ResourceType = typeof(LAWResource))]
        public string CreateDate { get; set; }

        /// <summary>設定人員名單給{Jason}</summary>
        public string LawimemberToJson { get; set; }

        /// <summary>1是利率,2是服務費率</summary>
        public string LirType { get; set; }

        /// <summary>利率</summary>
        [Display(Name = "利率", ResourceType = typeof(LAWResource))]
        public decimal InterestRates { get; set; }

        /// <summary>律師服務費率</summary>
        [Display(Name = "律師服務費率", ResourceType = typeof(LAWResource))]
        public decimal LawyerServiceRates { get; set; }

        /// <summary>結案狀態名稱</summary>
        [Display(Name = "結案狀態名稱", ResourceType = typeof(LAWResource))]
        public string CloseTypeName { get; set; }

        /// <summary>結案狀態</summary>
        [Display(Name = "結案狀態", ResourceType = typeof(LAWResource))]
        public string CountType { get; set; }

        /// <summary>契撤變原因</summary>
        [Display(Name = "契撤變原因", ResourceType = typeof(LAWResource))]
        public string EvidTypeName { get; set; }

        /// <summary>契撤變原因狀態</summary>
        [Display(Name = "狀態", ResourceType = typeof(LAWResource))]
        public int EvidStatusType { get; set; }

        /// <summary>年度</summary>
        public string yy { get; set; }

        /// <summary>流水號</summary>
        public int SortId { get; set; }

        /// <summary>區塊排序</summary>
        public int SortVm { get; set; }
    }
}