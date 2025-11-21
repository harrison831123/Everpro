using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.LAW.Models
{
    public class LawSearchDetail : IModel
    {
        /// <summary>
        /// 流水號
        /// </summary>
        [DataMember]
        [DisplayName("law_id")]
        [Column("law_id", IsKey = true, IsIdentity = true)]
        public int LawId { get; set; }

        /// <summary>照會單號</summary>
        [DataMember]
        [DisplayName("law_note_no")]
        [Column("law_note_no")]
        public string LawNoteNo { get; set; }

        /// <summary>年度</summary>
        [DataMember]
        [DisplayName("law_year")]
        [Column("law_year")]
        public string LawYear { get; set; }

        /// <summary>月份</summary>
        [DataMember]
        [DisplayName("law_month")]
        [Column("law_month")]
        public string LawMonth { get; set; }

        /// <summary>薪次</summary>
        [DataMember]
        [DisplayName("law_pay_sequence")]
        [Column("law_pay_sequence")]
        public int LawPaySequence { get; set; }

        /// <summary>卷宗</summary>
        [DataMember]
        [DisplayName("law_file_no")]
        [Column("law_file_no")]
        public string LawFileNo { get; set; }

        /// <summary>業務員姓名</summary>
        [DataMember]
        [DisplayName("law_due_name")]
        [Column("law_due_name")]
        public string LawDueName { get; set; }

        /// <summary>業務員id</summary>
        [DataMember]
        [DisplayName("law_due_agentid")]
        [Column("law_due_agentid")]
        public string LawDueAgentId { get; set; }

        /// <summary>結欠金額</summary>
        [DataMember]
        [DisplayName("law_due_money")]
        [Column("law_due_money")]
        public decimal LawDueMoney { get; set; }

        /// <summary>利息起算日</summary>
        [DataMember]
        [DisplayName("law_interest_sdate")]
        [Column("law_interest_sdate")]
        public string LawInterestSdate { get; set; }

        /// <summary>利息結算日</summary>
        [DataMember]
        [DisplayName("law_interest_edate")]
        [Column("law_interest_edate")]
        public string LawInterestEdate { get; set; }

        /// <summary>利息天數</summary>
        [DataMember]
        [DisplayName("law_interest_days")]
        [Column("law_interest_days")]
        public int LawInterestDays { get; set; }

        /// <summary>利息代碼</summary>
        [DataMember]
        [DisplayName("law_interest_rates_id")]
        [Column("law_interest_rates_id")]
        public int LawInterestRatesId { get; set; }

        /// <summary>利息金額</summary>
        [DataMember]
        [DisplayName("law_interest_money")]
        [Column("law_interest_money")]
        public int LawInterestMoney { get; set; }

        /// <summary>總金額</summary>
        [DataMember]
        [DisplayName("law_total_money")]
        [Column("law_total_money")]
        public int LawTotalMoney { get; set; }

        /// <summary>萬用帳號</summary>
        [DataMember]
        [DisplayName("law_super_account")]
        [Column("law_super_account")]
        public string LawSuperAccount { get; set; }

        /// <summary>結欠原因</summary>
        [DataMember]
        [DisplayName("law_due_reason")]
        [Column("law_due_reason")]
        public string LawDueReason { get; set; }

        /// <summary>案件狀態
        /// 0:受理完成 1:處理(進行)中 2:結案
        /// </summary>
        [DataMember]
        [DisplayName("law_status_type")]
        [Column("law_status_type")]
        public int LawStatusType { get; set; }

        /// <summary>承辦單位代碼</summary>
        [DataMember]
        [DisplayName("law_do_unit_id")]
        [Column("law_do_unit_id")]
        public int LawDoUnitId { get; set; }

        /// <summary>承辦單位名稱</summary>
        [DataMember]
        [DisplayName("law_do_unit_name")]
        [Column("law_do_unit_name")]
        public string LawDoUnitName { get; set; }

        /// <summary>承辦人員姓名</summary>
        [DataMember]
        [DisplayName("law_douser_name")]
        [Column("law_douser_name")]
        public string LawDouserName { get; set; }

        /// <summary>未結案名稱</summary>
        [DataMember]
        [DisplayName("law_not_close_type_name")]
        [Column("law_not_close_type_name")]
        public string LawNotCloseTypeName { get; set; }

        /// <summary>結案代碼</summary>
        [DataMember]
        [DisplayName("law_close_type")]
        [Column("law_close_type")]
        public string LawCloseType { get; set; }

        /// <summary>結案日期</summary>
        [DataMember]
        [DisplayName("law_close_date")]
        [Column("law_close_date")]
        public DateTime LawCloseDate { get; set; }

        /// <summary>結案人員</summary>
        [DataMember]
        [DisplayName("law_closer_name")]
        [Column("law_closer_name")]
        public string LawCloserName { get; set; }

        /// <summary>案件建檔人員</summary>
        [DataMember]
        [DisplayName("law_content_create_name")]
        [Column("law_content_create_name")]
        public string LawContentCreateName { get; set; }

        /// <summary>案件建檔日期</summary>
        [DataMember]
        [DisplayName("law_content_create_date")]
        [Column("law_content_create_date")]
        public string LawContentCreateDate { get; set; }

        /// <summary>第一次電催內容</summary>
        [DataMember]
        [DisplayName("law_phone_call1_desc")]
        [Column("law_phone_call1_desc")]
        public string LawPhoneCall1Desc { get; set; }

        /// <summary>第一次電催日期</summary>
        [DataMember]
        [DisplayName("law_phone_call1_date")]
        [Column("law_phone_call1_date")]
        public DateTime LawPhoneCall1Date { get; set; }

        /// <summary>第二次電催內容</summary>
        [DataMember]
        [DisplayName("law_phone_call2_desc")]
        [Column("law_phone_call2_desc")]
        public string LawPhoneCall2Desc { get; set; }

        /// <summary>第二次電催日期</summary>
        [DataMember]
        [DisplayName("law_phone_call2_date")]
        [Column("law_phone_call2_date")]
        public DateTime LawPhoneCall2Date { get; set; }

        /// <summary>案件取消	
        /// Null:有效 1:取消
        /// </summary>
        [DataMember]
        [DisplayName("law_content_cancel_type")]
        [Column("law_content_cancel_type")]
        public string LawContentCancelType { get; set; }

        /// <summary>最後更新日期(法追進度)</summary>
        [DataMember]
        [DisplayName("law_content_lastchange_date")]
        [Column("law_content_lastchange_date")]
        public string LawContentLastchangeDate { get; set; }

        /// <summary>案件建檔人員ID</summary>
        [DataMember]
        [DisplayName("LawContentCreatorID")]
        [Column("LawContentCreatorID")]
        public string LawContentCreatorID { get; set; }

        /// <summary>結案人員ID</summary>
        [DataMember]
        [DisplayName("LawContentCloserID")]
        [Column("LawContentCloserID")]
        public string LawContentCloserID { get; set; }

        /// <summary>承辦人員ID</summary>
        [DataMember]
        [DisplayName("LawContentDouserID")]
        [Column("LawContentDouserID")]
        public string LawContentDouserID { get; set; }

        /// <summary>
        /// 流水號
        /// </summary>
        [DataMember]
        [DisplayName("law_agent_content_id")]
        [Column("law_agent_content_id", IsKey = true, IsIdentity = true)]
        public int LawAgentContentId { get; set; }

        /// <summary>工作年月</summary>
        [DataMember]
        [DisplayName("production_ym")]
        [Column("production_ym")]
        public string ProductionYm { get; set; }

        /// <summary>薪次</summary>
        [DataMember]
        [DisplayName("sequence")]
        [Column("sequence")]
        public string Sequence { get; set; }

        /// <summary>團隊代碼</summary>
        [DataMember]
        [DisplayName("vm_code")]
        [Column("vm_code")]
        public string VmCode { get; set; }

        /// <summary>團隊名稱</summary>
        [DataMember]
        [DisplayName("vm_name")]
        [Column("vm_name")]
        public string VmName { get; set; }


        /// <summary>體系代碼</summary>
        [DataMember]
        [DisplayName("sm_code")]
        [Column("sm_code")]
        public string SmCode { get; set; }

        /// <summary>體系名稱</summary>
        [DataMember]
        [DisplayName("sm_name")]
        [Column("sm_name")]
        public string SmName { get; set; }

        /// <summary>實駐代碼</summary>
        [DataMember]
        [DisplayName("wc_center")]
        [Column("wc_center")]
        public string WcCenter { get; set; }

        /// <summary>實駐名稱</summary>
        [DataMember]
        [DisplayName("wc_center_name")]
        [Column("wc_center_name")]
        public string WcCenterName { get; set; }

        /// <summary>處代碼</summary>
        [DataMember]
        [DisplayName("center_code")]
        [Column("center_code")]
        public string CenterCode { get; set; }

        /// <summary>處名稱</summary>
        [DataMember]
        [DisplayName("center_name")]
        [Column("center_name")]
        public string CenterName { get; set; }

        /// <summary>處經理ID</summary>
        [DataMember]
        [DisplayName("administrat_id")]
        [Column("administrat_id")]
        public string AdministratId { get; set; }

        /// <summary>處經理姓名</summary>
        [DataMember]
        [DisplayName("admin_name")]
        [Column("admin_name")]
        public string AdminName { get; set; }

        /// <summary>處經理職級</summary>
        [DataMember]
        [DisplayName("admin_level")]
        [Column("admin_level")]
        public string AdminLevel { get; set; }

        /// <summary>業務員ID</summary>
        [DataMember]
        [DisplayName("agent_code")]
        [Column("agent_code")]
        public string AgentCode { get; set; }

        /// <summary>業務員姓名</summary>
        [DataMember]
        [DisplayName("name")]
        [Column("name")]
        public string Name { get; set; }

        /// <summary>業務員任用代碼</summary>
        [DataMember]
        [DisplayName("ag_status_code")]
        [Column("ag_status_code")]
        public string AgStatusCode { get; set; }

        /// <summary>業務員職級</summary>
        [DataMember]
        [DisplayName("ag_level")]
        [Column("ag_level")]
        public string AgLevel { get; set; }

        /// <summary>職級名稱</summary>
        [DataMember]
        [DisplayName("ag_level_name")]
        [Column("ag_level_name")]
        public string AgLevelName { get; set; }

        /// <summary>登錄日期</summary>
        [DataMember]
        [DisplayName("record_date")]
        [Column("record_date")]
        public string RecordDate { get; set; }

        /// <summary>簽約日期</summary>
        [DataMember]
        [DisplayName("register_date")]
        [Column("register_date")]
        public string RegisterDate { get; set; }

        /// <summary>解約日期</summary>
        [DataMember]
        [DisplayName("ag_status_date")]
        [Column("ag_status_date")]
        public string AgStatusDate { get; set; }

        /// <summary>萬用帳號</summary>
        [DataMember]
        [DisplayName("super_account")]
        [Column("super_account")]
        public string SuperAccount { get; set; }

        /// <summary>建立日期</summary>
        [DataMember]
        [DisplayName("create_date")]
        [Column("create_date")]
        public string CreateDate { get; set; }     

        [NonColumn]
        public string CloseTypeName { get; set; }

        [NonColumn]
        public string LawSearchType { get; set; }
        
        /// <summary>清償本金</summary>
        [DataMember]
        [DisplayName("law_repayment_capital")]
        [Column("law_repayment_capital")]
        public int LawRepaymentCapital { get; set; }

        /// <summary>清償金額</summary>
        [DataMember]
        [DisplayName("law_repayment_money")]
        [Column("law_repayment_money")]
        public int LawRepaymentMoney { get; set; }

        /// <summary>存證信函備註</summary>
        [DataMember]
        [DisplayName("law_evidence_desc")]
        [Column("law_evidence_desc")]
        public string LawEvidencedesc { get; set; }
        
        /// <summary>訴訟程序內容</summary>
        [DataMember]
        [DisplayName("law_litigation_progress")]
        [Column("law_litigation_progress")]
        public string LawLitigationprogress { get; set; }

        /// <summary>執行程序說明</summary>
        [DataMember]
        [DisplayName("law_do_progress")]
        [Column("law_do_progress")]
        public string LawDoprogress { get; set; }

        /// <summary>實駐名稱</summary>
        [DataMember]
        [DisplayName("WcCenterNameCG")]
        [Column("WcCenterNameCG")]
        public string WcCenterNameCG { get; set; }

        /// <summary>處名稱</summary>
        [DataMember]
        [DisplayName("CenterNameCG")]
        [Column("CenterNameCG")]
        public string CenterNameCG { get; set; }

        /// <summary>判斷是否為管理員</summary>
        /// 0 否 1 是
        [NonColumn]
        public int SysType { get; set; }

        /// <summary>使用者</summary>
        [NonColumn]
        public string UserName { get; set; }
    }
}
