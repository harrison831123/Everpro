using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SACTAPI.Models
{
    /// <summary>
    /// 簽核回傳物件
    /// </summary>
    public class ONextSignOff
    {
        public ONextSignOff()
        {
            responseCode = string.Empty;
            responseMsg = string.Empty;
            responseObj1 = null;
            responseObj2 = null;
        }

        /// <summary>
        /// 回傳代碼
        /// </summary>
        public string responseCode { get; set; }

        /// <summary>
        /// 回傳訊息
        /// </summary>
        public string responseMsg { get; set; }

        /// <summary>
        /// 成功物件：簽核人員資料
        /// </summary>
        public SignoffResultObj responseObj1 { get; set; }

        /// <summary>
        /// 第二個成功物件（簽約人資料），
        /// 僅在輔導人簽核 = Y 時才會回傳
        /// </summary>
        public SignerInfoObj responseObj2 { get; set; }
    }

    /// <summary>
    /// 簽核結果物件
    /// </summary>
    public class SignoffResultObj
    {
        /// <summary>
        /// 簽約人員 ID
        /// </summary>
        public string AGID { get; set; }

        /// <summary>
        /// 下一位簽核業務員 ID
        /// </summary>
        public string NextSignOffID { get; set; }

        /// <summary>
        /// 下一位簽核業務員 姓名
        /// </summary>
        public string NextSignOffName { get; set; }

        /// <summary>
        /// 下一位簽核者稱謂
        /// </summary>
        public string NextSignOffLevelName { get; set; }

        /// <summary>
        /// 是否為最後簽核者
        /// </summary>
        public string NextCode { get; set; }
    }

    /// <summary>
    /// 簽約人資料物件
    /// </summary>
    public class SignerInfoObj
    {
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