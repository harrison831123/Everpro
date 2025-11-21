using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 結案歷史紀錄
	/// </summary>
	[DataContract]
	[Table("CRMECloseLog")]
	public class CRMECloseLog : IModel
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
		/// 處理結果代碼
		/// </summary>
		[DataMember]
		[Column("ResultCode")]
		[Display(Name = "處理結果代碼")]
		public int? ResultCode { get; set; }

		/// <summary>
		/// 處理結果代碼2
		/// </summary>
		[DataMember]
		[Column("ResultCode2")]
		[Display(Name = "處理結果代碼2")]
		public int? ResultCode2 { get; set; }

		/// <summary>
		/// 結案者員編
		/// </summary>
		[DataMember]
		[Column("Creator")]
		[Display(Name = "結案者員編")]
		public string Creator { get; set; }

		/// <summary>
		/// 結案建檔日期
		/// </summary>
		[DataMember]
		[Column("CreateTime")]
		[Display(Name = "結案建檔日期")]
		public DateTime? CreateTime { get; set; }

	}

}
