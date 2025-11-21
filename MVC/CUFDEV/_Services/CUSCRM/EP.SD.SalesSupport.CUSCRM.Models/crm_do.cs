using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;


namespace EP.SD.SalesSupport.CUSCRM
{
    /// <summary>
	/// 聯繫紀錄
	/// </summary>
	[DataContract]
    [Table("crm_do")]
    public class tcrm_do :IModel
    {
		/// <summary>
		/// 自動編號
		/// </summary>
		[DataMember]
		[Column("crm_do_id", IsIdentity = true)]
		[Display(Name = "自動編號")]
		public int crm_do_id { get; set; }
		/// <summary>
		/// 受理編號
		/// </summary>
		[DataMember]
		[Column("crm_no")]
		[Display(Name = "受理編號")]
		public string crm_no { get; set; }
		/// <summary>
		/// 摘要
		/// </summary>
		[DataMember]
		[Column("crm_do")]
		[Display(Name = "摘要")]
		public string crm_do { get; set; }
		/// <summary>
		/// 輸入者id
		/// </summary>
		[DataMember]
		[Column("crm_do_createid")]
		[Display(Name = "輸入者id")]
		public string crm_do_createid { get; set; }

		/// <summary>
		/// 輸入者
		/// </summary>
		[DataMember]
		[Column("crm_do_createname")]
		[Display(Name = "輸入者")]
		public string crm_do_createname { get; set; }
		/// <summary>
		/// 日期
		/// </summary>
		[DataMember]
		[Column("crm_do_createdate")]
		[Display(Name = "日期")]
		public DateTime crm_do_createdate { get; set; }

	}
}
