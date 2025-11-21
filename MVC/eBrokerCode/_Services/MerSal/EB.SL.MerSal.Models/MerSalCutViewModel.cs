using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.MerSal.Models
{
    public class MerSalCutViewModel : IModel
    {

        /// <summary>
        /// 資料序號
        /// </summary>
        [Column("IDD")]
        [DisplayName("資料序號")]
        public string IDD { get; set; }

        /// <summary>
        /// 轉入工作月/業績年月
        /// </summary>
        [DisplayName("轉入工作月")]
        [Column("production_ym")]
        public string ProductionYM { get; set; }

        /// <summary>
        /// 序號
        /// </summary>
        [DisplayName("轉入序號")]
        [Column("sequence")]
        public string Sequence { get; set; }

        /// <summary>
        /// 保險公司
        /// </summary>
        [DisplayName("保險公司")]
        [Column("company_code")]
        public string CompanyCode { get; set; }

        /// <summary>
        /// 業務員ID1
        /// </summary>
        [Column("agent_code1")]
        [DisplayName("業務員ID1")]
        public string AgentCode1 { get; set; }

        /// <summary>
        /// 業務員姓名1
        /// </summary>
        [Column("names1")]
        [DisplayName("業務員姓名1")]
        public string Names1 { get; set; }

        /// <summary>
        /// 業務員ID2
        /// </summary>
        [Column("agent_code2")]
        [DisplayName("業務員ID2")]
        public string AgentCode2 { get; set; }

        /// <summary>
        /// 業務員姓名2
        /// </summary>
        [Column("names2")]
        [DisplayName("業務員姓名2")]
        public string Names2 { get; set; }

        /// <summary>
        /// 調整原因碼
        /// </summary>
        [Column("reason_code")]
        [DisplayName("調整原因碼")]
        public string ReasonCode { get; set; }

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
        public string CollectYear { get; set; }

        /// <summary>
        /// 繳別繳次
        /// </summary>
        [Column("modx_sequence")]
        [DisplayName("繳別繳次")]
        public string ModxSequence { get; set; }

        /// <summary>
        /// 被保險人
        /// </summary>
        [Column("insured_name")]
        [DisplayName("被保險人")]
        public string InsuredName { get; set; }

        /// <summary>
        /// 年齡
        /// </summary>
        [Column("age")]
        [DisplayName("年齡")]
        public string Age { get; set; }

        /// <summary>
        /// 
        /// </summary>
        [Column("po_issue_date")]
        [DisplayName("生效日")]
        public string PoIssueDate { get; set; }

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
        /// 大批調整年月
        /// </summary>
        [Column("pay_month")]
        [DisplayName("大批調整年月")] 
        public string PayMonth { get; set; }

        /// <summary>
        /// 大批調整序號
        /// </summary>
        [Column("pay_seq")]
        [DisplayName("大批調整序號")] 
        public string PaySeq { get; set; }

        /// <summary>
        /// 系統保留
        /// </summary>
        [Column("check_ind")]
        [DisplayName("系統保留")]
        public string CheckInd { get; set; }

        /// <summary>
        /// 判斷按下哪個button
        /// </summary>
        public string BtnType { get; set; }

        /// <summary>
        /// 空白欄位
        /// </summary>
        public string EmpCol1 { get; set; }
        public string EmpCol2 { get; set; }
        public string EmpCol3 { get; set; }
        public string EmpCol4 { get; set; }
        public string EmpCol5 { get; set; }
        public string EmpCol6 { get; set; }
        public string EmpCol7 { get; set; }

        /// <summary>
        /// 型態
        /// </summary>
        [Column("check_ind_name")]
        [DisplayName("型態")]
        public string CheckIndName { get; set; }

    }
}
