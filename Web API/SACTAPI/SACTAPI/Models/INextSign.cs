using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SACTAPI.Models
{
    public class INextSign
    {
        /// <summary>
        /// 保經公司代碼
        /// </summary>
        public string Broker { get; set; }

        /// <summary>
        /// TOKEN
        /// </summary>
        public string Token { get; set; }

        /// <summary>
        /// 廠商代碼
        /// </summary>
        public string Insurer { get; set; }

        /// <summary>
        /// 推介人ID
        /// </summary>
        public string IntroducerID { get; set; }

        /// <summary>
        /// 推介人及輔導人是否為同一人
        /// </summary>
        public string SameYN { get; set; }

        /// <summary>
        /// 輔導人ID
        /// </summary>
        public string DirectorID { get; set; }

        /// <summary>
        /// 簽約人ID
        /// </summary>
        public string AGID { get; set; }

        /// <summary>
        /// 簽約人員的稱謂代碼
        /// </summary>
        public string LevelCode { get; set; }

        /// <summary>
        /// 推介人是否已簽核
        /// </summary>
        public string IntroducerSign { get; set; }

        /// <summary>
        /// 輔導人是否已簽核
        /// </summary>
        public string DirectorSign { get; set; }

        /// <summary>
        /// 輔導人OM是否已簽核
        /// </summary>
        public string OMSign { get; set; }

        /// <summary>
        /// 輔導人的SM是否已簽核
        /// </summary>
        public string SMSign { get; set; }

        /// <summary>
        /// 輔導人的VM是否已簽核
        /// </summary>
        public string VMSign { get; set; }
    }
}