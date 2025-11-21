using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 催辦紀錄
	/// </summary>
	[DataContract]
	[Table("CRMEDoS")]
	public class CRMEDoS : IModel
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
		/// 催辦記錄
		/// </summary>
		[DataMember]
		[Column("Content")]
		[Display(Name = "催辦記錄")]
		public string Content { get; set; }

		/// <summary>
		/// 催辦者員編
		/// </summary>
		[DataMember]
		[Column("Creator")]
		[Display(Name = "催辦者員編")]
		public string Creator { get; set; }

		/// <summary>
		/// 催辦建立時間
		/// </summary>
		[DataMember]
		[Column("CreateTime")]
		[Display(Name = "催辦建立時間")]
		public DateTime CreateTime { get; set; }

	}

}
