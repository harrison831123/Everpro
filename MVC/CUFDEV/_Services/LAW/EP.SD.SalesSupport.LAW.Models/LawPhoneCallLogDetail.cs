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
    public class LawPhoneCallLogDetail : IModel
    {
        /// <summary>
        /// 流水號
        /// </summary>
        [DataMember]
        [DisplayName("law_call_log_id")]
        [Column("law_call_log_id", IsKey = true, IsIdentity = true)]
        public int LawCallLogId { get; set; }

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

        /// <summary>電催序號(1-2次)</summary>
        [DataMember]
        [DisplayName("law_phone_call_no")]
        [Column("law_phone_call_no")]
        public string LawPhoneCallNo { get; set; }

        /// <summary>電催日期</summary>
        [DataMember]
        [DisplayName("law_phone_call_date")]
        [Column("law_phone_call_date")]
        public string LawPhoneCallDate { get; set; }

        /// <summary>電催7日期限日期</summary>
        [DataMember]
        [DisplayName("law_phone_call_limited_date")]
        [Column("law_phone_call_limited_date")]
        public string LawPhoneCallLimitedDate { get; set; }

        /// <summary>電催讀取記錄</summary>
        [DataMember]
        [DisplayName("law_phone_call_read_log")]
        [Column("law_phone_call_read_log")]
        public string LawPhoneCallReadLog { get; set; }

        /// <summary>讀取日期</summary>
        [DataMember]
        [DisplayName("law_phone_call_readlog_date")]
        [Column("law_phone_call_readlog_date")]
        public string LawPhoneCallReadlogDate { get; set; }

        /// <summary>建檔人員</summary>
        [DataMember]
        [DisplayName("PhoneCallCreatorID")]
        [Column("PhoneCallCreatorID")]
        public string PhoneCallCreatorID { get; set; }

        /// <summary>電催讀取人員</summary>
        [DataMember]
        [DisplayName("PhoneCallReadID")]
        [Column("PhoneCallReadID")]
        public string PhoneCallReadID { get; set; }

        /// <summary>最後更新日期(法追進度)</summary>
        [DataMember]
        [DisplayName("law_content_lastchange_date")]
        [Column("law_content_lastchange_date")]
        public string LawContentLastchangeDate { get; set; }

        /// <summary>承辦人員ID</summary>
        [DataMember]
        [DisplayName("LawContentDouserID")]
        [Column("LawContentDouserID")]
        public string LawContentDouserID { get; set; }

        /// <summary>承辦人員姓名</summary>
        [DataMember]
        [DisplayName("law_douser_name")]
        [Column("law_douser_name")]
        public string LawDouserName { get; set; }
    }
}
