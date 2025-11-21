using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 資料設定
	/// </summary>
	[Table("CRMEDiscipType")]
	public class CRMEDiscipType : IModel
	{

		/// <summary>
		/// 自動編號
		/// </summary>
		[Column("ID", IsIdentity = true)]
		[Display(Name = "自動編號")]
		public int ID { get; set; }

		/// <summary>
		/// 代碼
		/// </summary>
		[Column("Code")]
		[Display(Name = "代碼")]
		public DiscipTypeCode Code { get; set; }

		/// <summary>
		/// 代碼名稱
		/// </summary>
		[Column("Name")]
		[Display(Name = "代碼名稱")]
		public string Name { get; set; }

		/// <summary>
		/// 狀態
		/// </summary>
		[Column("Status")]
		[Display(Name = "狀態")]
		public EnableStatus? Status { get; set; }

		/// <summary>
		/// 類別
		/// </summary>
		[Column("Kind")]
		[Display(Name = "類別")]
		public DiscipTypeKind? Kind { get; set; }

		/// <summary>
		/// 排序
		/// </summary>
		[Column("Sort")]
		[Display(Name = "排序")]
		public int? Sort { get; set; }

		/// <summary>
		/// 建立人員
		/// </summary>
		[Column("Creator")]
		[Display(Name = "建立人員")]
		public string Creator { get; set; }

		/// <summary>
		/// 建立時間
		/// </summary>
		[Column("CreateTime")]
		[Display(Name = "建立時間")]
		public DateTime CreateTime { get; set; }

		/// <summary>
		/// 修改人員
		/// </summary>
		[Column("Updator")]
		[Display(Name = "修改人員")]
		public string Updator { get; set; }

		/// <summary>
		/// 修改時間
		/// </summary>
		[Column("UpdateTime")]
		[Display(Name = "修改時間")]
		public DateTime? UpdateTime { get; set; }

	}

}
