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
    public class LawAgentDetail : IModel
    {
        /// <summary>
        /// 流水號
        /// </summary>
        [DataMember]
        [DisplayName("law_agent_content_id")]
        [Column("law_agent_content_id", IsKey = true, IsIdentity = true)]
        public int LawAgentContentId { get; set; }

        /// <summary>照會單號</summary>
        [DataMember]
        [DisplayName("law_note_no")]
        [Column("law_note_no")]
        public string LawNoteNo { get; set; }

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
        [DisplayName("團隊")]
        [Column("vm_name")]
        public string VmName { get; set; }


        /// <summary>體系代碼</summary>
        [DataMember]
        [DisplayName("sm_code")]
        [Column("sm_code")]
        public string SmCode { get; set; }

        /// <summary>體系名稱</summary>
        [DataMember]
        [DisplayName("體系")]
        [Column("sm_name")]
        public string SmName { get; set; }

        /// <summary>實駐代碼</summary>
        [DataMember]
        [DisplayName("wc_center")]
        [Column("wc_center")]
        public string WcCenter { get; set; }

        /// <summary>實駐名稱</summary>
        [DataMember]
        [DisplayName("實駐")]
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
        [DisplayName("處經理")]
        [Column("admin_name")]
        public string AdminName { get; set; }

        /// <summary>處經理職級</summary>
        [DataMember]
        [DisplayName("admin_level")]
        [Column("admin_level")]
        public string AdminLevel { get; set; }

        /// <summary>業務員ID</summary>
        [DataMember]
        [DisplayName("結欠人員ID")]
        [Column("agent_code")]
        public string AgentCode { get; set; }

        /// <summary>業務員姓名</summary>
        [DataMember]
        [DisplayName("結欠人員姓名")]
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
        [DisplayName("職級")]
        [Column("ag_level_name")]
        public string AgLevelName { get; set; }

        /// <summary>登錄日期</summary>
        [DataMember]
        [DisplayName("登錄日期")]
        [Column("record_date")]
        public string RecordDate { get; set; }

        /// <summary>簽約日期</summary>
        [DataMember]
        [DisplayName("簽約日期")]
        [Column("register_date")]
        public string RegisterDate { get; set; }

        /// <summary>解約日期</summary>
        [DataMember]
        [DisplayName("解約日期")]
        [Column("ag_status_date")]
        public string AgStatusDate { get; set; }

        /// <summary>萬用帳號</summary>
        [DataMember]
        [DisplayName("萬用帳號")]
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
        [DisplayName("出生日期")]
        [Column("birth")]
        public string Birth { get; set; }

        /// <summary>電話1</summary>
        [DataMember]
        [DisplayName("住家聯絡電話1")]
        [Column("phone_01")]
        public string Phone01 { get; set; }

        /// <summary>電話2</summary>
        [DataMember]
        [DisplayName("住家聯絡電話2")]
        [Column("phone_02")]
        public string Phone02 { get; set; }

        /// <summary>手機1</summary>
        [DataMember]
        [DisplayName("手機號碼1")]
        [Column("cell_01")]
        public string Cell01 { get; set; }

        /// <summary>手機2</summary>
        [DataMember]
        [DisplayName("手機號碼2")]
        [Column("cell_02")]
        public string Cell02 { get; set; }

        /// <summary>戶籍地址</summary>
        [DataMember]
        [DisplayName("戶籍地址")]
        [Column("address")]
        public string Address { get; set; }

        /// <summary>聯絡地址1</summary>
        [DataMember]
        [DisplayName("通訊地址1")]
        [Column("address_01")]
        public string Address01 { get; set; }

        /// <summary>聯絡地址2</summary>
        [DataMember]
        [DisplayName("通訊地址2")]
        [Column("address_02")]
        public string Address02 { get; set; }

        [NonColumn]
        public string NoteStr { get; set; }

        [NonColumn]
        public string StatusStr { get; set; }

        [NonColumn]
        public string oldname { get; set; }

        [NonColumn]
        public string AgentID { get; set; }

        

    }
}
