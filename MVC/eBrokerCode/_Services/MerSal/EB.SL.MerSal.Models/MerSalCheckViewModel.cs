using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace EB.SL.MerSal.Models
{
    public class MerSalCheckViewModel : IModel
    {
        /// <summary>
        /// 自動識別碼
        /// </summary>
        [Column("Iden", IsKey = true, IsIdentity = true)]
        public int Iden { get; set; }

        /// <summary>
        /// 業績年月
        /// </summary>
        [DisplayName("轉入工作月")]
        [Column("production_ym")]
        public string ProductionYM { get; set; }

        /// <summary>
        /// 序號
        /// </summary>
        [DisplayName("轉入次薪")]
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
        /// 佣酬類別
        /// </summary>
        [DisplayName("佣酬類別")]
        [Column("amount_type")]
        public string AmountType { get; set; }

        /// <summary>
        /// 保單號碼
        /// </summary>
        [Column("policy_no2")]
        [DisplayName("保單號碼")]
        public string PolicyNo2 { get; set; }

        /// <summary>
        /// 保公險種
        /// </summary>
        [Column("plan_code")]
        [DisplayName("險種代碼")]
        public string PlanCode { get; set; }

        /// <summary>
        /// 年期
        /// </summary>
        [Column("collect_year")]
        [DisplayName("年期")]
        public short CollectYear { get; set; }

        /// <summary>
        /// 繳別繳次
        /// </summary>
        [Column("modx_sequence")]
        [DisplayName("繳別繳次")]
        public string ModxSequence { get; set; }

        /// <summary>
        /// 保公原因碼
        /// </summary>
        [Column("reason_code_c")]
        [DisplayName("保公原因碼")]
        public string ReasonCodeC { get; set; }

        /// <summary>
        /// 保費
        /// </summary>
        [Column("mode_prem")]
        [DisplayName("保費")]
        public string ModePrem { get; set; }

        /// <summary>
        /// 保公佣金
        /// </summary>
        [Column("comm_prem_c")]
        [DisplayName("保公佣金")]
        public string CommPremC { get; set; }

        /// <summary>
        /// 永達佣金
        /// </summary>
        [Column("amount")]
        [DisplayName("永達佣金")]
        public string Amount { get; set; }

        /// <summary>
        /// 永達FYC
        /// </summary>
        [Column("comm_prem")]
        [DisplayName("永達FYC")]
        public string CommPrem { get; set; }

        /// <summary>
        /// 預收前端狀態
        /// </summary>
        [Column("policy_status_name")]
        [DisplayName("預收前端狀態")]
        public string PolicyStatusName { get; set; }

        /// <summary>
        /// 簽收回條日
        /// </summary>
        [Column("receipt_date")]
        [DisplayName("簽收回條日")]
        public string ReceiptDate { get; set; }

        /// <summary>
        /// 檢核註記
        /// </summary>
        [Column("check_ind")]
        [DisplayName("檢核註記")]
        public string CheckInd { get; set; }

        /// <summary>
        /// 佣酬發放註記
        /// </summary>
        [Column("pay_type")]
        [DisplayName("發放註記")]
        public string PayType { get; set; }

        /// <summary>
        /// 佣酬發放註記名稱
        /// </summary>
        [Column("pay_type_name")]
        [DisplayName("佣酬發放註記名稱")]
        public string PayTypeName { get; set; }

        /// <summary>
        /// 報表結轉下期註記
        /// </summary>
        [Column("rpt_include_flag")]
        [DisplayName("報表結轉下期註記")]
        public string RptIncludeFlag { get; set; }

        /// <summary>
        /// 系統備註
        /// </summary>
        [Column("remark")]
        [DisplayName("系統備註")]
        public string Remark { get; set; }

        /// <summary>
        /// 人工備註
        /// </summary>
        [Column("memo")]
        [DisplayName("人工備註")]
        public string Memo { get; set; }

        /// <summary>
        /// 佣酬發放勾選顯示註記
        /// </summary>
        [Column("paytype_show")]
        [DisplayName("佣酬發放勾選顯示註記")]
        public string PaytypeShow { get; set; }

        /// <summary>
        /// 發放工作月-序號
        /// </summary>
        [Column("pay_ym")]
        [DisplayName("發放工作月")]
        public string PayYm { get; set; }

        /// <summary>
        /// 發放工作月
        /// </summary>
        [Column("pay_month")]
        [DisplayName("發放工作月")]
        public string PayMonth { get; set; }

        /// <summary>
        /// 發放序號
        /// </summary>
        [Column("pay_seq")]
        [DisplayName("發放序號")]
        public string PaySeq { get; set; }

        /// <summary>
        /// NotPay年月-起
        /// </summary>
        [Column("not_pay_YMSS")]
        [DisplayName("異動工作月起")]
        public string NotPayYMSS { get; set; }

        /// <summary>
        /// NotPay年月-迄
        /// </summary>
        [Column("not_pay_YMSE")]
        [DisplayName("異動工作月迄")]
        public string NotPayYMSE { get; set; }

        /// <summary>
        /// 錯誤代碼
        /// </summary>
        [Column("chk_code")]
        [DisplayName("錯誤代碼")]
        public string ChkCode { get; set; }

        /// <summary>
        /// NotPay年月序號
        /// </summary>
        [Column("not_pay_YMS")]
        [DisplayName("NotPay年月序號")]
        public string NotPayYMS { get; set; }

        /// <summary>
        /// 最後異動人員
        /// </summary>
        [Column("process_user_code")]
        [DisplayName("最後異動人員")]
        public string ProcessUserCode { get; set; }

        /// <summary>
        /// half件
        /// </summary>
        [Column("ag_cnt")]
        [DisplayName("half件")]
        public string AgCnt { get; set; }

        /// <summary>
        /// 目前業績年月[GetProductionYmNow]
        /// </summary>
        public string YMNow { get; set; }

        /// <summary>
        /// 目前人事關檔業績年月[GetAgymIndOne]
        /// </summary>
        public string YmAgymCloseNow { get; set; }

        //需求單號：20240715004 新增佣酬發放調整報表欄位，生效日、年齡、招攬人姓名及ID等欄位 vicky
        /// <summary>
        /// 生效日
        /// </summary>
        [Column("po_issue_date")]
        [DisplayName("生效日")]
        public string PoIssueDate { get; set; }

        /// <summary>
        /// 年齡
        /// </summary>
        [Column("age")]
        [DisplayName("年齡")]
        public string Age { get; set; }

        /// <summary>
        /// 招攬人1姓名
        /// </summary>
        [Column("agent_name1")]
        [DisplayName("招攬人1姓名")]
        public string AgentName1 { get; set; }

        /// <summary>
        /// 招攬人1ID
        /// </summary>
        [Column("agent_code1")]
        [DisplayName("招攬人1ID")]
        public string AgentCode1 { get; set; }

        /// <summary>
        /// 招攬人2姓名
        /// </summary>
        [Column("agent_name2")]
        [DisplayName("招攬人2姓名")]
        public string AgentName2 { get; set; }

        /// <summary>
        /// 招攬人2ID
        /// </summary>
        [Column("agent_code2")]
        [DisplayName("招攬人2ID")]
        public string AgentCode2 { get; set; }

    }
}
