using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.MerSal.Models
{
    public class MerSalViewModel : IModel
    {
        /// <summary>
        /// 業績年月
        /// </summary>
        [DisplayName("業績年月")]
        [Column("production_ym")]
        public string ProductionYM { get; set; }

        /// <summary>
        /// 序號
        /// </summary>
        [DisplayName("序號")]
        [Column("sequence")]
        public short Sequence { get; set; }

        /// <summary>
        /// 保險公司
        /// </summary>
        [DisplayName("保險公司")]
        [Column("company_code")]
        public string CompanyCode { get; set; }

        /// <summary>
        /// 檔案序號
        /// </summary>
        [DisplayName("檔案序號")]
        [Column("file_seq")]
        public string FileSeq { get; set; }

        /// <summary>
        /// 執行檢核筆數
        /// </summary>
        [Column("imp_count")]
        public string ImpCount { get; set; }

        /// <summary>
        /// 錯誤筆數
        /// </summary>
        [Column("err_count")]
        public string ErrCount { get; set; }

        /// <summary>
        /// 警告筆數
        /// </summary>
        [Column("war_count")]
        public string WarCount { get; set; }

        /// <summary>
        /// 入正式筆數
        /// </summary>
        [Column("formal_count")]
        public string FormalCount { get; set; }

        /// <summary>
        /// 處理注記	0：轉入MerSalForCheck；1：檢核中；2：檢核結束；3：資料入MerSalCheck；D：已刪除
        /// </summary>
        [Column("process_status_MSRName")]
        public string ProcessStatusMSR { get; set; }

        /// <summary>
        /// 檢核時間起
        /// </summary>
        [Column("data_checktime_s")]
        public string DataChecktimeS { get; set; }

        /// <summary>
        ///  檢核時間訖
        /// </summary>
        [Column("data_checktime_e")]
        public string DataChecktimeE { get; set; }

        /// <summary>
        /// 執行者
        /// </summary>
        [Column("create_user_code")]
        public string CreateUserCode { get; set; }

        /// <summary>
        /// 入剪檔筆數
        /// </summary>
        [Column("cut_count")]
        public string CutCount { get; set; }

        /// <summary>
        /// 業務員資料來源
        /// </summary>
        [Column("ag_from")]
        public string AgFrom { get; set; }

        /// <summary>
        /// 業務員資料來源
        /// </summary>
        [Column("ag_fromName")]
        public string AgFromName { get; set; }

        /// <summary>
        /// 業務員資料來源
        /// </summary>
        [NonColumn]
        public int ReportType { get; set; }

        /// <summary>
        /// 查詢人員
        /// </summary>
        [NonColumn]
        public string QueryUser { get; set; }

        /// <summary>
        /// 查詢日期
        /// </summary>
        [NonColumn]
        public DateTime QueryDate { get; set; }

    }
}
