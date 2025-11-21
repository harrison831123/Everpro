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
    public class LawEvidenceDetail : IModel
    {
        /// <summary>
        /// 流水號
        /// </summary>
        [DataMember]
        [DisplayName("evid_id")]
        [Column("evid_id", IsKey = true, IsIdentity = true)]
        public int EvidId { get; set; }

        /// <summary>
        /// 主檔流水號
        /// </summary>
        [DataMember]
        [DisplayName("law_id")]
        [Column("law_id")]
        public int LawId { get; set; }

        /// <summary>照會單號</summary>
        [DataMember]
        [DisplayName("law_note_no")]
        [Column("law_note_no")]
        public string LawNoteNo { get; set; }

        /// <summary>存證信函字號</summary>
        [DataMember]
        [DisplayName("存證信函字號")]
        [Column("evid_no")]
        public string EvidNo { get; set; }

        /// <summary>寄件人(公司)</summary>
        [DataMember]
        [DisplayName("evid_sender")]
        [Column("evid_sender")]
        public string EvidSender { get; set; }

        /// <summary>寄件(公司)地址</summary>
        [DataMember]
        [DisplayName("evid_sender_add")]
        [Column("evid_sender_add")]
        public string EvidSenderAdd { get; set; }

        /// <summary>業務id</summary>
        [DataMember]
        [DisplayName("evid_agent_id")]
        [Column("evid_agent_id")]
        public string EvidAgentId { get; set; }

        /// <summary>業務員姓名</summary>
        [DataMember]
        [DisplayName("evid_agent_name")]
        [Column("evid_agent_name")]
        public string EvidAgentName { get; set; }

        /// <summary>戶籍地址</summary>
        [DataMember]
        [DisplayName("evid_agent_add")]
        [Column("evid_agent_add")]
        public string EvidAgentAdd { get; set; }

        /// <summary>連絡地址1</summary>
        [DataMember]
        [DisplayName("evid_agent_add1")]
        [Column("evid_agent_add1")]
        public string EvidAgentAdd1 { get; set; }

        /// <summary>連絡地址2</summary>
        [DataMember]
        [DisplayName("evid_agent_add2")]
        [Column("evid_agent_add2")]
        public string EvidAgentAdd2 { get; set; }

        /// <summary>契撤變原因</summary>
        [DataMember]
        [DisplayName("evid_reason")]
        [Column("evid_reason")]
        public string EvidReason { get; set; }

        /// <summary>大寫金額</summary>
        [DataMember]
        [DisplayName("evid_money")]
        [Column("evid_money")]
        public string EvidMoney { get; set; }

        /// <summary>金額</summary>
        [DataMember]
        [DisplayName("金額")]
        [Column("evid_money_num")]
        public int EvidMoneyNum { get; set; }

        /// <summary>萬用帳號</summary>
        [DataMember]
        [DisplayName("萬用帳號")]
        [Column("evid_account")]
        public string EvidAccount { get; set; }

        /// <summary>承辦人</summary>
        [DataMember]
        [DisplayName("evid_user")]
        [Column("evid_user")]
        public string EvidUser { get; set; }

        /// <summary>分機</summary>
        [DataMember]
        [DisplayName("evid_phone")]
        [Column("evid_phone")]
        public string EvidPhone { get; set; }

        /// <summary>建檔者</summary>
        [DataMember]
        [DisplayName("evid_create_name")]
        [Column("evid_create_name")]
        public string EvidCreateName { get; set; }

        /// <summary>建檔日期</summary>
        [DataMember]
        [DisplayName("evid_create_date")]
        [Column("evid_create_date")]
        public DateTime EvidCreateDate { get; set; }

        /// <summary>建檔者ID</summary>
        [DataMember]
        [DisplayName("EvidenceCreatorID")]
        [Column("EvidenceCreatorID")]
        public string EvidenceCreatorID { get; set; }

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

        /// <summary>
        /// 流水號
        /// </summary>
        [DataMember]
        [DisplayName("law_agent_data_id")]
        [Column("law_agent_data_id", IsKey = true, IsIdentity = true)]
        public int LawAgentDataId { get; set; }

        /// <summary>生日</summary>
        [DataMember]
        [DisplayName("birth")]
        [Column("birth")]
        public string Birth { get; set; }

        /// <summary>電話1</summary>
        [DataMember]
        [DisplayName("phone_01")]
        [Column("phone_01")]
        public string Phone01 { get; set; }

        /// <summary>電話2</summary>
        [DataMember]
        [DisplayName("phone_02")]
        [Column("phone_02")]
        public string Phone02 { get; set; }

        /// <summary>手機1</summary>
        [DataMember]
        [DisplayName("cell_01")]
        [Column("cell_01")]
        public string Cell01 { get; set; }

        /// <summary>手機2</summary>
        [DataMember]
        [DisplayName("cell_02")]
        [Column("cell_02")]
        public string Cell02 { get; set; }

        /// <summary>戶籍地址</summary>
        [DataMember]
        [DisplayName("address")]
        [Column("address")]
        public string Address { get; set; }

        /// <summary>聯絡地址1</summary>
        [DataMember]
        [DisplayName("address_01")]
        [Column("address_01")]
        public string Address01 { get; set; }

        /// <summary>聯絡地址2</summary>
        [DataMember]
        [DisplayName("address_02")]
        [Column("address_02")]
        public string Address02 { get; set; }

        /// <summary>頁面狀態</summary>
        [NonColumn]
        public string MyPageStatus { get; set; }

        [NonColumn]
        public string EvidenceLogCreatorID { get; set; }

        [NonColumn]
        public string reason { get; set; }

    }
}
