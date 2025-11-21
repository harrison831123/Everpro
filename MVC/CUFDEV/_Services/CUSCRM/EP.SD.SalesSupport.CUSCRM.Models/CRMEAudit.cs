using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;


namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 稽催紀錄
	/// </summary>
	[Table("CRMEAudit")]
	public class CRMEAudit : IModel
	{

		/// <summary>
		/// 自動編號
		/// </summary>
		[Column("ID", IsIdentity = true)]
		[Display(Name = "自動編號")]
		public int ID { get; set; }

		/// <summary>
		/// 受理號碼
		/// </summary>
		[Column("No")]
		[Display(Name = "受理號碼")]
		public string No { get; set; }

		/// <summary>
		/// 稽催狀態
		/// </summary>
		[Column("Type")]
		[Display(Name = "稽催狀態")]
		public int? Type { get; set; }

		/// <summary>
		/// 稽催日期
		/// </summary>
		[Column("Date")]
		[Display(Name = "稽催日期")]
		public DateTime? Date { get; set; }

		/// <summary>
		/// 稽催內容
		/// </summary>
		[Column("Content")]
		[Display(Name = "稽催內容")]
		public string Content { get; set; }

		/// <summary>
		/// 稽催人員員編
		/// </summary>
		[Column("Creator")]
		[Display(Name = "稽催人員員編")]
		public string Creator { get; set; }

		/// <summary>
		/// 稽催建立時間
		/// </summary>
		[Column("CreateTime")]
		[Display(Name = "稽催建立時間")]
		public DateTime CreateTime { get; set; }

	}

}
