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
    public class LawContentDetail
    {
        /// <summary>
        /// 流水號
        /// </summary>
        [DataMember]
        [DisplayName("流水號")]
        [Column("law_id", IsKey = true, IsIdentity = true)]
        public int LawId { get; set; }

        /// <summary>照會單號</summary>
        [DataMember]
        [DisplayName("照會單號")]
        [Column("law_note_no")]
        public string LawNoteNo { get; set; }

        /// <summary>年度</summary>
        [DataMember]
        [DisplayName("年度")]
        [Column("law_year")]
        public string LawYear { get; set; }

        /// <summary>月份</summary>
        [DataMember]
        [DisplayName("月份")]
        [Column("law_month")]
        public string LawMonth { get; set; }

        /// <summary>薪次</summary>
        [DataMember]
        [DisplayName("薪次")]
        [Column("law_pay_sequence")]
        public int LawPaySequence { get; set; }

        /// <summary>卷宗</summary>
        [DataMember]
        [DisplayName("卷宗")]
        [Column("law_file_no")]
        public string LawFileNo { get; set; }

        /// <summary>業務員姓名</summary>
        [DataMember]
        [DisplayName("業務員姓名")]
        [Column("law_due_name")]
        public string LawDueName { get; set; }

        /// <summary>業務員id</summary>
        [DataMember]
        [DisplayName("業務員id")]
        [Column("law_due_agentid")]
        public string LawDueAgentId { get; set; }

        /// <summary>結欠金額</summary>
        [DataMember]
        [DisplayName("結欠金額")]
        [Column("law_due_money")]
        public decimal LawDueMoney { get; set; }

        /// <summary>結欠金額</summary>
        [NonColumn]
        public decimal OldLawDueMoney { get; set; }
        
        /// <summary>利息起算日</summary>
        [DataMember]
        [DisplayName("利息起算日")]
        [Column("law_interest_sdate")]
        public string LawInterestSdate { get; set; }

        /// <summary>利息起算日</summary>
        [NonColumn]
        public string OldLawInterestSdate { get; set; }

        /// <summary>利息結算日</summary>
        [DataMember]
        [DisplayName("利息結算日")]
        [Column("law_interest_edate")]
        public string LawInterestEdate { get; set; }

        /// <summary>利息結算日</summary>
        [NonColumn]
        public string OldLawInterestEdate { get; set; }

        /// <summary>利息天數</summary>
        [DataMember]
        [DisplayName("利息天數")]
        [Column("law_interest_days")]
        public int LawInterestDays { get; set; }

        /// <summary>利息代碼</summary>
        [DataMember]
        [DisplayName("利息代碼")]
        [Column("law_interest_rates_id")]
        public int LawInterestRatesId { get; set; }

        /// <summary>利息代碼</summary>
        [NonColumn]
        public int OldLawInterestRatesId { get; set; }

        /// <summary>利息金額</summary>
        [DataMember]
        [DisplayName("利息金額")]
        [Column("law_interest_money")]
        public int LawInterestMoney { get; set; }

        /// <summary>總金額</summary>
        [DataMember]
        [DisplayName("總金額")]
        [Column("law_total_money")]
        public int LawTotalMoney { get; set; }

        /// <summary>萬用帳號</summary>
        [DataMember]
        [DisplayName("萬用帳號")]
        [Column("law_super_account")]
        public string LawSuperAccount { get; set; }

        /// <summary>結欠原因</summary>
        [DataMember]
        [DisplayName("結欠原因")]
        [Column("law_due_reason")]
        public string LawDueReason { get; set; }

        /// <summary>案件狀態
        /// 0:受理完成 1:處理(進行)中 2:結案
        /// </summary>
        [DataMember]
        [DisplayName("案件狀態")]
        [Column("law_status_type")]
        public int LawStatusType { get; set; }

        /// <summary>承辦單位代碼</summary>
        [DataMember]
        [DisplayName("承辦單位代碼")]
        [Column("law_do_unit_id")]
        public int LawDoUnitId { get; set; }

        /// <summary>承辦單位名稱</summary>
        [DataMember]
        [DisplayName("承辦單位名稱")]
        [Column("law_do_unit_name")]
        public string LawDoUnitName { get; set; }

        /// <summary>承辦人員姓名</summary>
        [DataMember]
        [DisplayName("承辦人員姓名")]
        [Column("law_douser_name")]
        public string LawDouserName { get; set; }

        /// <summary>未結案名稱</summary>
        [DataMember]
        [DisplayName("未結案名稱")]
        [Column("law_not_close_type_name")]
        public string LawNotCloseTypeName { get; set; }

        /// <summary>結案代碼</summary>
        [DataMember]
        [DisplayName("結案代碼")]
        [Column("law_close_type")]
        public string LawCloseType { get; set; }

        /// <summary>結案日期</summary>
        [DataMember]
        [DisplayName("結案日期")]
        [Column("law_close_date")]
        public string LawCloseDate { get; set; }

        /// <summary>結案人員</summary>
        [DataMember]
        [DisplayName("結案人員")]
        [Column("law_closer_name")]
        public string LawCloserName { get; set; }

        /// <summary>案件建檔人員</summary>
        [DataMember]
        [DisplayName("案件建檔人員")]
        [Column("law_content_create_name")]
        public string LawContentCreateName { get; set; }

        /// <summary>案件建檔日期</summary>
        [DataMember]
        [DisplayName("通知日期")]
        [Column("law_content_create_date")]
        public string LawContentCreateDate { get; set; }

        /// <summary>第一次電催內容</summary>
        [DataMember]
        [DisplayName("第一次電催內容")]
        [Column("law_phone_call1_desc")]
        public string LawPhoneCall1Desc { get; set; }

        /// <summary>第一次電催日期</summary>
        [DataMember]
        [DisplayName("第一次電催日期")]
        [Column("law_phone_call1_date")]
        public string LawPhoneCall1Date { get; set; }

        /// <summary>第一次電催內容</summary>
        [NonColumn]
        public string OldLawPhoneCall1Desc { get; set; }

        /// <summary>第一次電催日期</summary>
        [NonColumn]
        public string OldLawPhoneCall1Date { get; set; }

        /// <summary>第二次電催內容</summary>
        [DataMember]
        [DisplayName("第二次電催內容")]
        [Column("law_phone_call2_desc")]
        public string LawPhoneCall2Desc { get; set; }

        /// <summary>第二次電催日期</summary>
        [DataMember]
        [DisplayName("第二次電催日期")]
        [Column("law_phone_call2_date")]
        public string LawPhoneCall2Date { get; set; }

        /// <summary>第二次電催內容</summary>
        [NonColumn]
        public string OldLawPhoneCall2Desc { get; set; }

        /// <summary>第二次電催日期</summary>
        [NonColumn]
        public string OldLawPhoneCall2Date { get; set; }

        /// <summary>案件取消	
        /// Null:有效 1:取消
        /// </summary>
        [DataMember]
        [DisplayName("案件取消")]
        [Column("law_content_cancel_type")]
        public string LawContentCancelType { get; set; }

        /// <summary>最後更新日期(法追進度)</summary>
        [DataMember]
        [DisplayName("最後更新日期(法追進度)")]
        [Column("law_content_lastchange_date")]
        public string LawContentLastchangeDate { get; set; }

        /// <summary>案件建檔人員ID</summary>
        [DataMember]
        [DisplayName("案件建檔人員ID")]
        [Column("LawContentCreatorID")]
        public string LawContentCreatorID { get; set; }

        /// <summary>結案人員ID</summary>
        [DataMember]
        [DisplayName("結案人員ID")]
        [Column("LawContentCloserID")]
        public string LawContentCloserID { get; set; }

        /// <summary>承辦人員ID</summary>
        [DataMember]
        [DisplayName("承辦人員ID")]
        [Column("LawContentDouserID")]
        public string LawContentDouserID { get; set; }

        //存證信函
        [NonColumn]
        public int LawEvidenceId { get; set; }

        [DisplayName("存證信函")]
        [NonColumn]
        public string LawEvidencedesc { get; set; }

        //分案日期&承辦單位
        [NonColumn]
        public int LawDounitLogId { get; set; }

        [DisplayName("分案日期")]
        [NonColumn]
        public string CaseDate { get; set; }

        [NonColumn]
        public string OldCaseDate { get; set; }

        [NonColumn]
        public string UnitName { get; set; }

        [NonColumn]
        public string OldUnitName { get; set; }

        [NonColumn]
        public string dounitlist { get; set; }

        [NonColumn]
        public string AgentID { get; set; }

        [NonColumn]
        public string LawProductionYM { get; set; }

        [DisplayName("訴訟程序")]
        [NonColumn]
        public string LawLitigationprogress { get; set; }

        [DisplayName("執行程序")]
        [NonColumn]
        public string LawDoprogress { get; set; }

        [DisplayName("備註")]
        [NonColumn]
        public string LawDescdesc { get; set; }

        /// <summary>案件狀態其他說明</summary>
        [NonColumn]
        public string LawCloseOtherDesc { get; set; }

        /// <summary>案件狀態其他流水號</summary>
        [NonColumn]
        public int LawCloseOtherId { get; set; }

        /// <summary>0:未結 1:已結</summary>
        [NonColumn]
        public string closetype { get; set; }

        /// <summary>案件狀態其他流水號</summary>
        [NonColumn]
        public string nclose { get; set; }

        /// <summary>結案狀態名稱</summary>
        [NonColumn]
        public string CloseTypeName { get; set; }

        /// <summary>利率</summary>
        [NonColumn]
        public decimal InterestRates { get; set; }

        /// <summary>清償日期</summary>
        [NonColumn]
        [DisplayName("清償日期")]
        public string LawRepaymentDate { get; set; }

        /// <summary>清償金額</summary>
        [NonColumn]
        [DisplayName("清償金額")]
        public int LawRepaymentMoney { get; set; }

        /// <summary>累計清償金額</summary>
        [NonColumn]
        [DisplayName("累計清償金額")]
        public int TotalLawRepaymentMoney { get; set; }

        /// <summary>剩餘結欠金額</summary>
        [NonColumn]
        [DisplayName("剩餘結欠金額")]
        public string RemainLawRepaymentMoney { get; set; }

        /// <summary>清償本金</summary>
        [NonColumn]
        [DisplayName("清償本金")]
        public int LawRepaymentCapital { get; set; }

        /// <summary>清償利息</summary>
        [NonColumn]
        [DisplayName("清償利息")]
        public int LawRepaymentInterest { get; set; }

        /// <summary>法院規費</summary>
        [NonColumn]
        [DisplayName("法院規費")]
        public int LawRepaymentCourt { get; set; }

        /// <summary>其他規費</summary>
        [NonColumn]
        [DisplayName("其他規費")]
        public int LawRepaymentOther { get; set; }

        /// <summary>續佣扣抵</summary>
        [NonColumn]
        [DisplayName("續佣扣抵")]
        public int LawCommDeduction { get; set; }

        /// <summary>總金額</summary>
        [NonColumn]
        [DisplayName("總金額")]
        public int LawRepaymentTotal { get; set; }

        /// <summary>結欠頁面狀態</summary>
        [NonColumn]
        public string PayMentType { get; set; }

        /// <summary>
        /// 結欠流水號
        /// </summary>
        [NonColumn]
        public int LawRepaymentId { get; set; }

        /// <summary>
        /// 律師服務流水號
        /// </summary>
        [NonColumn]
        public int LawLawyerPayId { get; set; }

        /// <summary>給付年度</summary>
        [DisplayName("給付年度")]
        [NonColumn]
        public string LawRewardPayYear { get; set; }

        /// <summary>給付月份</summary>
        [DisplayName("給付月份")]
        [NonColumn]
        public string LawRewardPayMonth { get; set; }

        /// <summary>給付日期</summary>
        [DisplayName("給付日期")]
        [NonColumn]
        public string LawRewardPayYearMonth { get; set; }

        /// <summary>核扣規費</summary>
        [DisplayName("核扣規費")]
        [NonColumn]
        public int LawFees { get; set; }

        /// <summary>服務費率</summary>
        [DisplayName("服務費率")]
        [NonColumn]
        public decimal LawRates { get; set; }

        /// <summary>計算本金</summary>
        [DisplayName("計算本金")]
        [NonColumn]
        public int LawRepaymentMoneyORG { get; set; }

        /// <summary>計算比例</summary>
        [DisplayName("計算比例")]
        [NonColumn]
        public string LawyerServiceRates { get; set; }

        /// <summary>服務費報酬</summary>
        [DisplayName("服務費報酬")]
        [NonColumn]
        public int LawServiceReward { get; set; }

        /// <summary>其他備註說明</summary>
        [Column("law_other_desc")]
        public string LawOtherdesc { get; set; }

    }
}
