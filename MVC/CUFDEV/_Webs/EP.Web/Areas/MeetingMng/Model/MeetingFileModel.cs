using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EP.PSL.WorkResources.MeetingMng.Web.Areas.MeetingMng.Model
{
    public class MeetingFileModel
    {
        /// <summary>會議流水號</summary>
        public int MTID { get; set; }

        /// <summary>會議檔案流水號</summary>
        public int MFID { get; set; }

        /// <summary>檔案名稱</summary>
        public string MFFileName { get; set; }

        /// <summary>Md5檔名</summary>
        public string MFMd5Name { get; set; }
    }
}