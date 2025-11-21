using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OnlineSigning.Models
{
    public class NextSignOff
    {
        /// <summary>
        /// 回傳代碼
        /// </summary>
        public string ResponseCode { get; set; }

        /// <summary>
        /// 回傳訊息
        /// </summary>
        public string ResponseMsg { get; set; }

        /// <summary>
        /// 成功物件：簽核人員資料
        /// </summary>
        //public string ResponseObjSignOff { get; set; }

        /// <summary>
        /// 簽約人員
        /// </summary>
        public string AGID { get; set; }

        /// <summary>
        /// 下一位簽核業務員ID
        /// </summary>
        public string NextSignOffID { get; set; }

        /// <summary>
        /// 下一位簽核者稱謂
        /// </summary>
        public string NextSignOffLevelName { get; set; }

        /// <summary>
        /// 是否為最後簽核者
        /// </summary>
        public string NextCode { get; set; }

        /// <summary>
        /// 成功物件：簽約人資料
        /// </summary>
        //public string ResponseObjSign { get; set; }

        /// <summary>
        /// 簽約人的處代碼
        /// </summary>
        public string CenterCode { get; set; }

        /// <summary>
        /// 簽約人的處名稱
        /// </summary>
        public string CenterName { get; set; }

        /// <summary>
        /// 簽約人的通訊處代碼
        /// </summary>
        public string WcCenterCode { get; set; }

        /// <summary>
        /// 簽約人的通訊處名稱
        /// </summary>
        public string WcCenterName { get; set; }

        /// <summary>
        /// 簽約人的籌處代碼
        /// </summary>
        public string UmCode { get; set; }

        /// <summary>
        /// 簽約人的籌處名稱
        /// </summary>
        public string UmName { get; set; }
    }
}