//需求單號：20250707001 競賽計C商品清單。 2025/07/03 BY Harrison 
using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.Collections.PlanSet.Models
{
    public class PlanSetWarptSet : IModel
    {
        /// <summary>
        /// ID
        /// </summary>
        [DisplayName("ID")]
        [Column("ID", IsKey = true, IsIdentity = true)]
        public int ID { get; set; }

        /// <summary>
        /// 保險公司代碼
        /// </summary>
        [DisplayName("保險公司代碼")]
        [Column("company_code")]
        public string CompanyCode { get; set; }

        /// <summary>
        /// 保險公司名稱
        /// </summary>
        [DisplayName("保險公司名稱")]
        [Column("company_name")]
        public string CompanyName { get; set; }

        /// <summary>
        /// 險種代碼
        /// </summary>
        [DisplayName("險種代碼")]
        [Column("plan_code")]
        public string PlanCode { get; set; }

        /// <summary>
        /// 保公險種代碼
        /// </summary>
        [DisplayName("保公險種代碼")]
        [Column("plan_code_c")]
        public string PlanCodeC { get; set; }

        /// <summary>
        /// 險種名稱
        /// </summary>
        [DisplayName("險種名稱")]
        [Column("plan_title")]
        public string PlanTitle { get; set; }

        /// <summary>
        /// 繳費年期-起
        /// </summary>
        [DisplayName("繳費年期-起")]
        [Column("plan_year_str")]
        public string PlanYearStr { get; set; }

        /// <summary>
        /// 繳費年期-迄
        /// </summary>
        [DisplayName("繳費年期-迄")]
        [Column("plan_year_end")]
        public string PlanYearEnd { get; set; }

        /// <summary>
        /// 繳費年期
        /// </summary>
        [DisplayName("繳費年期")]
        [Column("plan_year_cond")]
        public string PlanYearCond { get; set; }

        /// <summary>
        /// 要保申請日起
        /// </summary>
        [DisplayName("要保申請日起")]
        [Column("set_start_date")]
        public string SetStartDate { get; set; }

        /// <summary>
        /// 要保申請日迄
        /// </summary>
        [DisplayName("要保申請日迄")]
        [Column("set_end_date")]
        public string SetEndDate { get; set; }

        /// <summary>
        /// 是否計入
        /// </summary>
        [DisplayName("是否計入")]
        [Column("IsCredited")]
        public string IsCredited { get; set; }

        /// <summary>
        /// 是否計入(中文)
        /// </summary>
        [DisplayName("是否計入(中文)")]
        [Column("IsCredited_txt")]
        public string IsCreditedTxt { get; set; }

        /// <summary>
        /// 現售商品
        /// </summary>
        [DisplayName("現售商品")]
        [Column("isCurrent")]
        public string IsCurrent { get; set; }

        /// <summary>
        /// 建立時間
        /// </summary>
        [DisplayName("建立時間")]
        [Column("create_datetime")]
        public DateTime CreateDatetime { get; set; }

    }
}
