using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.MerSal.Models
{
	public class MerSalCheckSPruleViewModel : IModel
	{
        /// <summary>
        /// 自動識別碼
        /// </summary>
        [DataMember]
        [Column("iden", IsKey = true, IsIdentity = true)]
        public int iden { get; set; }

        /// <summary>
        /// 工作月+序號起
        /// </summary>
        [DisplayName("工作月+序號起")]
        [Column("production_ym_s")]
        public string ProductionYmS { get; set; }

        /// <summary>
        /// 工作月+序號迄
        /// </summary>
        [DisplayName("工作月+序號迄")]
        [Column("production_ym_e")]
        public string ProductionYmE { get; set; }

        /// <summary>
        /// 保險公司代碼
        /// </summary>
        [Column("company_code")]
        public string CompanyCode { get; set; }

        /// <summary>
        /// 保險公司序號
        /// </summary>
        [Column("file_seq")]
        public string FileSeq { get; set; }

        /// <summary>
        /// 佣酬類別
        /// </summary>
        [Column("amount_type")]
        public string AmountType { get; set; }

        /// <summary>
        /// 佣酬類別名稱
        /// </summary>
        [Column("amount_type_name")]
        public string AmountTypeName { get; set; }

        /// <summary>
        /// 特殊檢核設定名稱
        /// </summary>
        [Column("chk_code_name")]
        public string ChkCodeName { get; set; }

        /// <summary>
        /// 特殊檢核
        /// </summary>
        [Column("act_type")]
        public string ActType { get; set; }

        /// <summary>
        /// 特殊檢核顯示中文
        /// </summary>
        [DisplayName("特殊檢核顯示中文")]
        [Column("act_name")]
        public string ActName { get; set; }

        /// <summary>
        /// 規則01
        /// </summary>
        [DisplayName("檢核規則1")]
        [Column("rule_01")]
        public string Rule01 { get; set; }

        /// <summary>
        /// 規則02
        /// </summary>
        [DisplayName("檢核規則2")]
        [Column("rule_02")]
        public string Rule02 { get; set; }

        /// <summary>
        /// 規則03
        /// </summary>
        [DisplayName("檢核規則3")]
        [Column("rule_03")]
        public string Rule03 { get; set; }

        /// <summary>
        /// 規則04
        /// </summary>
        [DisplayName("檢核規則4")]
        [Column("rule_04")]
        public string Rule04 { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        [DisplayName("備註")]
        [Column("remark")]
        public string Remark { get; set; }

        /// <summary>
        /// 建檔時間
        /// </summary>
        [Column("create_datetime")]
        public string CreateDatetime { get; set; }

        /// <summary>
        /// 建檔人員
        /// </summary>
        [Column("create_user_code")]
        public string CreateUserCode { get; set; }

        /// <summary>
        /// 建檔人員名稱
        /// </summary>
        [Column("create_user_name")]
        public string CreateUserName { get; set; }

        /// <summary>
        /// 更新時間
        /// </summary>
        [Column("update_datetime")]
        public string UpdateDatetime { get; set; }

        /// <summary>
        /// 更新人員
        /// </summary>
        [Column("update_user_code")]
        public string UpdateUserCode { get; set; }

        /// <summary>
        /// 更新人員姓名
        /// </summary>
        [Column("update_user_name")]
        public string UpdateUserName { get; set; }


        /// <summary>
        /// 停用狀態
        /// </summary>
        [Column("is_delete")]
        public string IsDelete { get; set; }

        /// <summary>
        /// chk_code+act_type
        /// </summary>
        [NonColumn]
        public string ChkCodeActType { get; set; }

        /// <summary>
        /// 最大業績已關檔年月
        /// </summary>
        [NonColumn]
        public string YmClose { get; set; }

        /// <summary>
        /// 最大業績已關檔序號
        /// </summary>
        [NonColumn]
        public string SeqClose { get; set; }
    }
}
