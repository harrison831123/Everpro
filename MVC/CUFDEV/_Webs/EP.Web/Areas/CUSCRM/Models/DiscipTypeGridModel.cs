using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EP.SD.SalesSupport.CUSCRM.Web
{
    /// <summary>
    /// 資料設查詢結果Model
    /// </summary>
    public class DiscipTypeGridModel
    {
        /// <summary>
        /// 自動編號
        /// </summary>
        public int ID { get; set; }

        /// <summary>
        /// 代碼
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 類別
        /// </summary>
        public string Kind { get; set; }

        /// <summary>
        /// 名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 狀態
        /// </summary>
        public string Status { get; set; }

    }
}