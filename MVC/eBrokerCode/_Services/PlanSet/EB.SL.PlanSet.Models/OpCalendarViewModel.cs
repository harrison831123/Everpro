using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EB.SL.PlanSet.Models
{
	public class OpCalendarViewModel : IModel
    {
        /// <summary>
        /// Iden
        /// </summary>
        [DataMember]
        [Column("iden")]
        [Display(Name = "自動識別碼")]
        public string Iden { get; set; }

        /// <summary>
        /// PorductionYm
        /// </summary>
        [DataMember]
        [Column("production_ym")]
        [Display(Name = "業績年月")]
        public string ProductionYM { get; set; }

        /// <summary>
        /// Sequence
        /// </summary>
        [DataMember]
        [Column("sequence")]
        [Display(Name = "序次")]
        public string Sequence { get; set; }

        /// <summary>
        /// HrCloseDate
        /// </summary>
        [DataMember]
        [Column("hr_close_date")]
        [Display(Name = "人事關檔日期")]
        public string HrCloseDate { get; set; }

        /// <summary>
        /// SalRunDate
        /// </summary>
        [DataMember]
        [Column("sal_run_date")]
        [Display(Name = "RUN佣日期")]
        public string SalRunDate { get; set; }

        /// <summary>
        /// SalPayDate
        /// </summary>
        [DataMember]
        [Column("sal_pay_date")]
        [Display(Name = "發佣日期")]
        public string SalPayDate { get; set; }

        /// <summary>
        /// sal_receipt_date
        /// </summary>
        [DataMember]
        [Column("sal_receipt_date")]
        [Display(Name = "簽收回條截止日")]
        public string SalReceiptDate { get; set; }


        /// <summary>
        /// open_query_date
        /// </summary>
        [DataMember]
        [Column("open_query_date")]
        [Display(Name = "佣酬明細開放查詢日")]
        public string OpenQueryDate { get; set; }

        /// <summary>
        /// open_query_date_ann
        /// </summary>
        [DataMember]
        [Column("open_query_date_ann")]
        [Display(Name = "年終明細開放查詢日")]
        public string OpenQueryDateAnn { get; set; }

        /// <summary>
        /// AdjDateTimeStr
        /// </summary>
        [DataMember]
        [Column("adj_datetime_str")]
        [Display(Name = "調整時間(起)")]
        public DateTime? AdjDateTimeStr { get; set; }

        /// <summary>
        /// AdjDateTimeEnd
        /// </summary>
        [DataMember]
        [Column("adj_datetime_end")]
        [Display(Name = "調整時間(迄)")]
        public DateTime? AdjDateTimeEnd { get; set; }

        /// <summary>
        /// Remark
        /// </summary>
        [DataMember]
        [Column("remark")]
        [Display(Name = "備註")]
        public string Remark { get; set; }

        /// <summary>
        /// CreateDateTime
        /// </summary>
        [DataMember]
        [Column("create_datetime")]
        [Display(Name = "資料建檔時間")]
        public string CreateDateTime { get; set; }

        /// <summary>
        /// CreateUserCode
        /// </summary>
        [DataMember]
        [Column("create_user_code")]
        [Display(Name = "資料建檔人員")]
        public string CreateUserCode { get; set; }

        /// <summary>
        /// CreateUserName
        /// </summary>
        [DataMember]
        [Column("create_user_name")]
        [Display(Name = "資料建檔人員")]
        public string CreateUserName { get; set; }

        /// <summary>
        /// UpdateDateTime
        /// </summary>
        [DataMember]
        [Column("update_datetime")]
        [Display(Name = "資料異動時間")]
        public string UpdateDateTime { get; set; }

        /// <summary>
        /// UpdateUserCode
        /// </summary>
        [DataMember]
        [Column("update_user_code")]
        [Display(Name = "資料異動人員")]
        public string UpdateUserCode { get; set; }

        /// <summary>
        /// UpdateUserName
        /// </summary>
        [DataMember]
        [Column("update_user_name")]
        [Display(Name = "資料異動人員")]
        public string UpdateUserName { get; set; }

        /// <summary>
        /// AdjDateTimeStr
        /// </summary>
        [NonColumn]
        public string AdjDateTimeStrView { get; set; }

        /// <summary>
        /// AdjDateTimeEnd
        /// </summary>
        [NonColumn]
        public string AdjDateTimeEndView { get; set; }
    }
}
