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
    public class LawNoteDetail : IModel
    {
        /// <summary>
        /// 流水號
        /// </summary>
        [DataMember]
        [DisplayName("law_note_id")]
        [Column("law_note_id", IsKey = true, IsIdentity = true)]
        public int LawNoteId { get; set; }

        /// <summary>照會單號</summary>
        [DataMember]
        [DisplayName("law_note_no")]
        [Column("law_note_no")]
        public string LawNoteNo { get; set; }

        /// <summary>留言編號</summary>
        [DataMember]
        [DisplayName("law_note_km")]
        [Column("law_note_km")]
        public int LawNoteKm { get; set; }

        /// <summary>單位</summary>
        [DataMember]
        [DisplayName("law_note_center")]
        [Column("law_note_center")]
        public string LawNoteCenter { get; set; }

        /// <summary>對象(主管)</summary>
        [DataMember]
        [DisplayName("law_note_to")]
        [Column("law_note_to")]
        public string LawNoteTo { get; set; }

        /// <summary>職稱</summary>
        [DataMember]
        [DisplayName("law_note_level")]
        [Column("law_note_level")]
        public string LawNoteLevel { get; set; }

        /// <summary>業務員姓名</summary>
        [DataMember]
        [DisplayName("業務員")]
        [Column("law_note_name")]
        public string LawNoteName { get; set; }

        /// <summary>部門</summary>
        [DataMember]
        [DisplayName("law_note_dep")]
        [Column("law_note_dep")]
        public string LawNoteDep { get; set; }

        /// <summary>經辦</summary>
        [DataMember]
        [DisplayName("law_note_pro")]
        [Column("law_note_pro")]
        public string LawNotePro { get; set; }

        /// <summary>分機</summary>
        [DataMember]
        [DisplayName("law_note_tel")]
        [Column("law_note_tel")]
        public string LawNoteTel { get; set; }

        /// <summary>發送方式 0:系統；1:人工</summary>
        [DataMember]
        [DisplayName("law_note_type")]
        [Column("law_note_type")]
        public int LawNoteType { get; set; }

        /// <summary>建檔人員</summary>
        [DataMember]
        [DisplayName("law_note_creatername")]
        [Column("law_note_creatername")]
        public string LawNoteCreatername { get; set; }

        /// <summary>建檔日期</summary>
        [DataMember]
        [DisplayName("law_note_createdate")]
        [Column("law_note_createdate")]
        public DateTime LawNoteCreatedate { get; set; }

        /// <summary>建檔人員ID</summary>
        [DataMember]
        [DisplayName("NoteCreatorID")]
        [Column("NoteCreatorID")]
        public string NoteCreatorID { get; set; }

        [NonColumn]
        public string BPMURL { get; set; }

        /// <summary>流水號</summary>
        [NonColumn]
        public int LawDouserId { get; set; }

        /// <summary>承辦人員姓名</summary>
        [NonColumn]
        public string DouserName { get; set; }

        /// <summary>分機號碼</summary>
        [NonColumn]
        public string DouserPhoneExt { get; set; }

        /// <summary>設定成員名單給{Jason}</summary>
        [NonColumn]
        [DisplayName("設定成員名單給{Jason}")]
        public string RecipientToJson { get; set; }

        /// <summary>業務員ID</summary>
        [NonColumn]
        public string AgentID { get; set; }

        /// <summary>
        /// 主檔流水號
        /// </summary>
        [NonColumn]
        public int LawId { get; set; }

    }
}
