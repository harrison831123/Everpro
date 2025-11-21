using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SACTAPI.Models
{
    public class SACTAPILog
    {
        /// <summary>
        /// API 名稱 (GetSignOffAGIDTOKEN、GetAGNameTOKEN、DownLoadAGData)
        /// </summary>
        public string ApiName { get; set; }

        /// <summary>
        /// 推介人 AGID
        /// </summary>
        public string Introducer { get; set; }

        /// <summary>
        /// 輔導人 AGID
        /// </summary>
        public string Director { get; set; }

        /// <summary>
        /// 輔導人登錄證號
        /// </summary>
        public string RegisterNo { get; set; }

        /// <summary>
        /// 被推介人 ID
        /// </summary>
        public string AGID { get; set; }

        /// <summary>
        /// 回傳代碼
        /// </summary>
        public string ResponseCode { get; set; }

        /// <summary>
        /// 訊息摘要 (錯誤訊息或其他資訊)
        /// </summary>
        public string LogMsg { get; set; }

        /// <summary>
        /// Request Log JSON
        /// </summary>
        public string LogRequestData { get; set; }

        /// <summary>
        /// Response Log JSON
        /// </summary>
        public string LogResponseData { get; set; }
    }
}