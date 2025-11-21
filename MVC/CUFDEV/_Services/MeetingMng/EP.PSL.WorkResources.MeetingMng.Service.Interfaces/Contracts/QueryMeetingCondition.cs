using EP.PSL.WorkResources.MeetingMng.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.PSL.WorkResources.MeetingMng.Service
{
    public class QueryMeetingCondition
    {
        /// <summary>流水號</summary>
        [Display(Name = "流水號", ResourceType = typeof(MeetingMngResource))]
        public int MTID { get; set; }

        /// <summary>
        /// 會議名稱
        /// </summary>
        [Display(Name = "會議名稱", ResourceType = typeof(MeetingMngResource))]
        public string MTName { get; set; }

        /// <summary>
        /// 會議說明
        /// </summary>
        [Display(Name = "會議說明", ResourceType = typeof(MeetingMngResource))]
        public string MTDesc { get; set; }

        /// <summary>
        /// 判斷登入人員
        /// </summary>
        [Display(Name = "人員", ResourceType = typeof(MeetingMngResource))]
        public string imember { get; set; }

        /// <summary>
        /// 會議類別 1未召開 2已召開 3我舉辦 4歷史資料
        /// </summary>
        [Display(Name = "類別", ResourceType = typeof(MeetingMngResource))]
        public string MeetingReadType { get; set; }
    }
}
