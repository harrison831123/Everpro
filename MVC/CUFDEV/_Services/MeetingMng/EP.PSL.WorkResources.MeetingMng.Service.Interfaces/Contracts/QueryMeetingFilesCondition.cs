using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.PSL.WorkResources.MeetingMng.Service
{
    public class QueryMeetingFilesCondition
    {
        /// <summary>流水號</summary>
        [Display(Name = "流水號")]
        public int MTID { get; set; }

        /// <summary>檔案說明</summary>
        [DisplayName("MFDesc")]
        public string MFDesc { get; set; }

        /// <summary>檔案名稱</summary>
        [DisplayName("MFFileName")]
        public string MFFileName { get; set; }
    }
}
