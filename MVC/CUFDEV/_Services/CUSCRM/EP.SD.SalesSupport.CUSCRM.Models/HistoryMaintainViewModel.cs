
using EP.Platform.Service;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;


namespace EP.SD.SalesSupport.CUSCRM
{
    public class HistoryMaintainViewModel
    {

        #region 顯示欄位
        [Display(Name = "受理編號")]
        public string crm_no { get; set; }

        /// <summary>
        /// 輸入者
        /// </summary>
        [Display(Name = "輸入者")]
        public string crm_do_createname { get; set; }
        /// <summary>
        /// 日期
        /// </summary>
        [Display(Name = "日期")]
        public string crm_do_createdate { get; set; }
        /// <summary>
        /// 時間
        /// </summary>
        [Display(Name = "時間")]
        public string crm_do_time { get; set; }

        public List<RecordViewModel> maintainlist { get; set; }

        public List<ValueText> filelist { get; set; }

        #endregion
        public class RecordViewModel
        {
            [Display(Name = "受理編號")]
			public string crm_no { get; set; }
            /// <summary>
            /// 摘要
            /// </summary>
            [Display(Name = "摘要")]
            public string crm_do { get; set; }

            /// <summary>
            /// 輸入者
            /// </summary>
            [Display(Name = "輸入者")]
            public string crm_do_createname { get; set; }
            /// <summary>
            /// 日期
            /// </summary>
            [Display(Name = "日期")]
            public string crm_do_createdate { get; set; }
            /// <summary>
            /// 時間
            /// </summary>
            [Display(Name = "時間")]
            public string crm_do_time { get; set; }
        }
    }
}
