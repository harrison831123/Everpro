using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;



namespace EP.SD.SalesSupport.CUSCRM
{
    /// <summary>
	/// 結案紀錄
	/// </summary>
	[DataContract]
    [Table("crm_close_log")]
    public class crm_close_log : IModel
    {
        /// <summary>
		/// 自動編號
		/// </summary>
		[DataMember]
        [Column("id", IsIdentity = true)]
        [Display(Name = "自動編號")]
        public int id { get; set; }
        [DataMember]
        [Column("crm_no")]
        [Display(Name = "受理編號")]
        public string crm_no { get; set; }
        [DataMember]
        [Column("old_crm_close")]
        [Display(Name = "")]
        public string old_crm_close { get; set; }
        [DataMember]
        [Column("new_crm_close")]
        [Display(Name = "")]
        public string new_crm_close { get; set; }
        [DataMember]
        [Column("desc_log")]
        [Display(Name = "")]
        public string desc_log { get; set; }
        [DataMember]
        [Column("creater_orgid")]
        [Display(Name = "")]
        public int creater_orgid { get; set; }
        [DataMember]
        [Column("crm_closeday")]
        [Display(Name = "")]
        public int crm_closeday { get; set; }
        [DataMember]
        [Column("new_createdate")]
        [Display(Name = "")]
        public DateTime new_createdate { get; set; }
        [DataMember]
        [Column("old_closedate")]
        [Display(Name = "")]
        public DateTime old_closedate { get; set; }
        [DataMember]
        [Column("new_closedate")]
        [Display(Name = "")]
        public DateTime new_closedate { get; set; }
        [DataMember]
        [Column("createdate")]
        [Display(Name = "")]
        public DateTime createdate { get; set; }
    }
}
