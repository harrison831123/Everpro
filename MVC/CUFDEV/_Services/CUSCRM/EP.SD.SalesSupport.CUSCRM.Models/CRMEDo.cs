using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 維護記錄
	/// </summary>
	[DataContract]
	[Table("CRMEDo")]
	public class CRMEDo : IModel
	{

		/// <summary>
		/// 自動編號
		/// </summary>
		[DataMember]
		[Column("ID", IsIdentity = true)]
		[Display(Name = "自動編號")]
		public int ID { get; set; }

		/// <summary>
		/// 受理號碼
		/// </summary>
		[DataMember]
		[Column("No")]
		[Display(Name = "受理號碼")]
		public string No { get; set; }

		/// <summary>
		/// 維護記錄
		/// </summary>
		[DataMember]
		[Column("Content")]
		[Display(Name = "維護記錄")]
		public string Content { get; set; }

		/// <summary>
		/// 業連保險公司日期
		/// </summary>
		[DataMember]
		[Column("BUSContactCompanyDate")]
		[Display(Name = "業連保險公司日期")]
		public DateTime? BUSContactCompanyDate { get; set; }

		/// <summary>
		/// 保險公司回覆日
		/// </summary>
		[DataMember]
		[Column("ReplyCompanyDate")]
		[Display(Name = "保險公司回覆日")]
		public DateTime? ReplyCompanyDate { get; set; }

		/// <summary>
		/// 回覆單位黃聯日
		/// </summary>
		[DataMember]
		[Column("ReplyYallowBillDate")]
		[Display(Name = "回覆單位黃聯日")]
		public DateTime? ReplyYallowBillDate { get; set; }

		/// <summary>
		/// 客戶申訴件(CC)的回文日
		/// </summary>
		[DataMember]
		[Column("CCReplyDate")]
		[Display(Name = "客戶申訴件的回文日")]
		public DateTime? CCReplyDate { get; set; }

		/// <summary>
		/// 維護者員編
		/// </summary>
		[DataMember]
		[Column("Creator")]
		[Display(Name = "維護者員編")]
		public string Creator { get; set; }

		/// <summary>
		/// 維護建立時間
		/// </summary>
		[DataMember]
		[Column("CreateTime")]
		[Display(Name = "維護建立時間")]
		public DateTime CreateTime { get; set; }

		/// <summary>
		/// 保公決議選項
		/// </summary>
		[DataMember]
		[Column("CompanyResult")]
		[Display(Name = "保公決議選項")]
		public int? CompanyResult { get; set; }
	}

}
