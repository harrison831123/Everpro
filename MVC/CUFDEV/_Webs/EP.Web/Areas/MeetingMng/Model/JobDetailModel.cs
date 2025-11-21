using EP.PSL.WorkResources.MeetingMng.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EP.PSL.WorkResources.MeetingMng.Web.Areas.MeetingMng.Model
{
    public class JobDetailModel
    {
        /// <summary>會議流水號</summary>

        [Display(Name = "流水號", ResourceType = typeof(MeetingMngResource))]

        public int MTID { get; set; }

        /// <summary>決議事項流水號</summary>

        [Display(Name = "流水號", ResourceType = typeof(MeetingMngResource))]
        public int JBID { get; set; }

        /// <summary>追蹤事項進度表流水號</summary>

        [Display(Name = "流水號", ResourceType = typeof(MeetingMngResource))]
        public int JPID { get; set; }

        /// <summary>判斷從哪個功能新增</summary>

        [DisplayName("JBClass")]
        public int JBClass { get; set; }

        /// <summary>追蹤事項標題</summary>

        [Display(Name = "標題", ResourceType = typeof(MeetingMngResource))]
        public string JBSubject { get; set; }

        /// <summary></summary>

        [DisplayName("JBPrevid")]
        public int JBPrevid { get; set; }

        /// <summary>追蹤說明</summary>

        [Display(Name = "說明", ResourceType = typeof(MeetingMngResource))]
        public string JBDesc { get; set; }

        /// <summary>開始日期</summary>

        [Display(Name = "開始日期", ResourceType = typeof(MeetingMngResource))]
        public DateTime JBStartDate { get; set; }

        /// <summary>結束日期</summary>
        [Display(Name = "結束日期", ResourceType = typeof(MeetingMngResource))]
        public DateTime JBEndDate { get; set; }

        /// <summary>建立者</summary>
        [Display(Name = "建立者", ResourceType = typeof(MeetingMngResource))]
        public string JBCreater { get; set; }

        /// <summary>建立日期</summary>
        [Display(Name = "建立日期", ResourceType = typeof(MeetingMngResource))]
        public DateTime JBCreateDate { get; set; }

        /// <summary>建立者IP</summary>

        [DisplayName("建立者IP")]
        public string JBCreateIP { get; set; }

        /// <summary>建立者姓名</summary>

        [DisplayName("建立者姓名")]
        public string Creater { get; set; }

        /// <summary>建立者部門</summary>

        [DisplayName("建立者部門")]
        public string CreaterNunit { get; set; }

        /// <summary></summary>

        [DisplayName("JBAssignment")]
        public int JBAssignment { get; set; }

        /// <summary>MTID</summary>

        [DisplayName("JBCount")]
        public int JBCount { get; set; }

        /// <summary></summary>

        [DisplayName("JBPrivate")]
        public int JBPrivate { get; set; }

        /// <summary>成員編號</summary>

        [DisplayName("imember")]
        public string imember { get; set; }

        /// <summary>完成度</summary>
        [Display(Name = "完成度", ResourceType = typeof(MeetingMngResource))]

        public int JPPercentage { get; set; }

        /// <summary>人員姓名&部門</summary>

        [Display(Name = "成員", ResourceType = typeof(MeetingMngResource))]
        public string JBUnitName { get; set; }
    }
}