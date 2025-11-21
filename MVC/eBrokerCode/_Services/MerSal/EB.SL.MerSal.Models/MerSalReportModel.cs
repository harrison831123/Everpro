using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.MerSal.Models
{
    public class MerSalReportModel : IModel
    {
        /// <summary>
        /// 自動識別碼
        /// </summary>
        [Column("iden", IsKey = true, IsIdentity = true)]
        public int iden { get; set; }

        /// <summary>
        /// OriSalRun序號
        /// </summary>
        [Column("run_seq")]
        public string RunSeq { get; set; }

        /// <summary>
        /// 工作月
        /// </summary>
        [Column("production_ym")]
        public string ProductionYm { get; set; }

        /// <summary>
        /// 序號
        /// </summary>
        [Column("sequence")]
        public string Sequence { get; set; }

        /// <summary>
        /// 保險公司
        /// </summary>
        [Column("company_code")]
        public string CompanyCode { get; set; }

        /// <summary>
        /// 檔案序號
        /// </summary>
        [Column("file_seq")]
        public int FileSeq { get; set; }

        /// <summary>
        /// 保單號碼
        /// </summary>
        [Column("policy_no2")]
        public string PolicyNo2 { get; set; }

        /// <summary>
        /// 檢核代碼
        /// </summary>
        [Column("chk_code")]
        public string ChkCode { get; set; }

        /// <summary>
        /// 錯誤代碼
        /// </summary>
        [Column("chk_code_type")]
        public string ChkCodeType { get; set; }

        /// <summary>
        /// 錯誤訊息
        /// </summary>
        [Column("error_type_msg")]
        public string ErrorTypeMsg { get; set; }

        [Column("error_line_no")]
        public string ErrorLineNo { get; set; }

        /// <summary>
        /// 錯誤資料顯示
        /// </summary>
        [Column("error_value")]
        public string ErrorValue { get; set; }

        [Column("bypass_type")]
        public string BypassType { get; set; }

        /// <summary>
        /// 執行時間
        /// </summary>
        [Column("create_datetime")]
        public string CreateDatetime { get; set; }

        /// <summary>
        /// 執行人員
        /// </summary>
        [Column("create_user_code")]
        public string CreateUserCode { get; set; }

        /// <summary>
        /// 執行人員姓名
        /// </summary>
        [Column("nmember")]
        public string nmember { get; set; }
    }
}
