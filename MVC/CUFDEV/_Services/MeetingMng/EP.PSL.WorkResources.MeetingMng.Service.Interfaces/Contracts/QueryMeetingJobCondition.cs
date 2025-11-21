using EP.PSL.WorkResources.MeetingMng.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.PSL.WorkResources.MeetingMng.Service
{
    public class QueryMeetingJobCondition
    {
        /// <summary>流水號</summary>
        [Display(Name = "流水號", ResourceType = typeof(MeetingMngResource))]
        public int MTID { get; set; }

        /// <summary>追蹤事項</summary>
        [Display(Name = "追蹤事項", ResourceType = typeof(MeetingMngResource))]
        public string JBSubject { get; set; }

        /// <summary>追蹤說明</summary>
        [Display(Name = "追蹤說明", ResourceType = typeof(MeetingMngResource))]
        public string JBDesc { get; set; }

        /// <summary>追蹤事項狀態</summary>
        [DisplayName("追蹤事項狀態")]
        public string MeetingJobReadType { get; set; }

        /// <summary>
        /// 判斷登入人員
        /// </summary>
        [Display(Name = "人員", ResourceType = typeof(MeetingMngResource))]
        public string imember { get; set; }

        /// <summary>
        /// 會議名稱
        /// </summary>
        [Display(Name = "會議名稱", ResourceType = typeof(MeetingMngResource))]
        public string MTName { get; set; }
    }
}
