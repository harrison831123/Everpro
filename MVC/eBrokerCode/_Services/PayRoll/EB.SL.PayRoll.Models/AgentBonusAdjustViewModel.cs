///========================================================================================== 
/// 程式名稱：人工調帳
/// 建立人員：Harrison
/// 建立日期：2022/07
/// 修改記錄：（[需求單號]、[修改內容]、日期、人員）
/// 需求單號:20240122004-因現有VLIFE系統(核心系統)使用已長達20多年，架構老舊，已不敷使用，且為提升資訊安全等級，故計劃執行VLIFE系統改版(新核心系統:eBroker系統)。; 修改內容:上線; 修改日期:20240613; 修改人員:Harrison;
/// 需求單號:20240807001-調整人工調帳系統產出之相關畫面及報表修改等功能。; 修改日期:20240807; 修改人員:Harrison;
///==========================================================================================
using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.PayRoll.Models
{
    public class AgentBonusAdjustViewModel : IModel
    {
        /// <summary>
        /// 自動識別碼
        /// </summary>
        [Column("iden", IsKey = true, IsIdentity = true)]
        [DisplayName("自動識別碼")]
        public int Iden { get; set; }

        /// <summary>
        /// 全域唯一識別碼
        /// </summary>
        [Column("guid")]
        [DisplayName("全域唯一識別碼")]
        public Guid Guid { get; set; }

        /// <summary>
        /// 業績年月
        /// </summary>
        [Required(ErrorMessage = "Required.")]
        [Column("production_ym")]
        [DisplayName("業績年月")]
        public string ProductionYM { get; set; }

        /// <summary>
        /// 序號
        /// </summary>
        [Column("sequence")]
        [DisplayName("序號")]
        public short Sequence { get; set; }

        /// <summary>
        /// 調整類別
        /// </summary>
        [Required(ErrorMessage = "Required.")]
        [Column("adj_type")]
        [DisplayName("型態")]
        public string AdjType { get; set; }

        /// <summary>
        /// 處理序號
        /// </summary>
        [Column("process_no")]
        [DisplayName("處理序號")]
        public string ProcessNo { get; set; }

        /// <summary>
        /// 保險公司代碼
        /// </summary>
        [Column("company_code")]
        [DisplayName("保險公司代碼")]
        public string CompanyCode { get; set; }

        /// <summary>
        /// 保險公司分公司代碼
        /// </summary>
        [Column("company_code_sub")]
        [DisplayName("保險公司分公司代碼")]
        public string CompanyCodeSub { get; set; }

        /// <summary>
        /// 實駐代碼
        /// </summary>
        [Column("wc_code")]
        [DisplayName("實駐代碼")]
        public string WCCode { get; set; }

        /// <summary>
        /// 保單號碼
        /// </summary>
        [Column("policy_no2")]
        [DisplayName("保單號碼")]
        public string PolicyNo2 { get; set; }

        /// <summary>
        /// 險種代碼
        /// </summary>
        [Column("plan_code")]
        [DisplayName("險種代碼")]
        public string PlanCode { get; set; }

        /// <summary>
        /// 繳費年期
        /// </summary>
        [Column("collect_year")]
        [DisplayName("繳費年期")]
        public short CollectYear { get; set; }

        /// <summary>
        /// 繳別繳次
        /// </summary>
        [Column("modx_sequence")]
        [DisplayName("繳別繳次")]
        public string ModxSequence { get; set; }

        /// <summary>
        /// 業務員代碼
        /// </summary>
        [Column("agent_code")]
        [DisplayName("業務員代碼")]
        public string AgentCode { get; set; }

        /// <summary>
        /// 業務員姓名
        /// </summary>
        [Column("agent_name")]
        [DisplayName("業務員姓名")]
        public string AgentName { get; set; }

        /// <summary>
        /// 原因碼
        /// </summary>
        [Column("reason_code")]
        [DisplayName("原因碼")]
        public string ReasonCode { get; set; }

        /// <summary>
        /// FYP
        /// </summary>
        [Column("fyp")]
        [DisplayName("保費")]
        public int FYP { get; set; }

        /// <summary>
        /// FYC
        /// </summary>
        [Column("fyc")]
        [DisplayName("FYC")]
        public int FYC { get; set; }

        /// <summary>
        /// 計佣保費
        /// </summary>
        [Column("comm_mode_prem")]
        [DisplayName("計佣保費")]
        public int CommModePrem { get; set; }

        /// <summary>
        /// 金額
        /// </summary>
        [Column("amount")]
        [DisplayName("金額")]
        public int Amount { get; set; }

        /// <summary>
        /// 程序代碼
        /// </summary>
        [Column("process_code")]
        [DisplayName("程序代碼")]
        public string ProcessCode { get; set; }

        /// <summary>
        /// 業務酬佣說明檔GUID
        /// </summary>
        [Column("AgentBonusDesc_guid")]
        [DisplayName("業務酬佣說明檔GUID")]
        public Guid AgentBonusDescGuid { get; set; }

        /// <summary>
        /// 資料建檔時間
        /// </summary>
        [Column("create_datetime")]
        [DisplayName("資料建檔時間")]
        public DateTime CreateDatetime { get; set; }

        /// <summary>
        /// 資料建檔人員
        /// </summary>
        [Column("create_user_code")]
        [DisplayName("資料建檔人員")]
        public string CreateUserCode { get; set; }

        /// <summary>
        /// 資料建檔人員
        /// </summary>
        [Column("create_unit_name")]
        [DisplayName("資料建檔部門")]
        public string CreateUnitName { get; set; }

        /// <summary>
        /// 說明內容
        /// </summary>
        [Column("desc_content")]
        [DisplayName("說明內容")]
        public string DescContent { get; set; }

        /// <summary>
        /// 資料建檔部門代碼
        /// </summary>
        [Column("create_unit")]
        [Display(Name = "資料建檔部門代碼")]
        public string CreateUnit { get; set; }

        /// <summary>
        /// 名字
        /// </summary>
        [Column("nmember")]
        [DisplayName("名字")]
        public string nmember { get; set; }

        /// <summary>
        /// 業績年
        /// </summary>
        [NonColumn]
        public string ProductionY { get; set; }

        /// <summary>
        /// 業績年月-序號
        /// </summary>
        [NonColumn]
        public string ProductionYMSequence { get; set; }

        /// <summary>
        /// 次薪(起)
        /// </summary>
        [NonColumn]
        public int SequenceStart { get; set; }

        /// <summary>
        /// 次薪(訖)
        /// </summary>
        [NonColumn]
        public int SequenceFinish { get; set; }

        /// <summary>
        /// 時間
        /// </summary>
        [NonColumn]
        public string CreateDatetimeString { get; set; }

        /// <summary>
        /// FYP-千分位
        /// </summary>
        [NonColumn]
        public string FYPTS { get; set; }

        /// <summary>
        /// FYC-千分位
        /// </summary>
        [NonColumn]
        public string FYCTS { get; set; }

        /// <summary>
        /// 金額-千分位
        /// </summary>
        [NonColumn]
        public string AmountTS { get; set; }
    }
}
