using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 通知作業查詢Model
	/// </summary>
	public class NotifyViewModel
	{
		/// <summary>
		/// 頁面所需的自動編號
		/// </summary>
		[Column("ViewID")]
		public int ViewID { get; set; }

		/// <summary>
		/// 類別
		/// </summary>
		[Column("Type")]
		public string Type { get; set; }

		/// <summary>
		/// 受理編號
		/// </summary>
		[Column("No")]
		public string No { get; set; }

		/// <summary>
		/// 客戶姓名
		/// </summary>
		[Column("Owner")]
		public string Owner { get; set; }

		/// <summary>
		/// 案件GUID
		/// </summary>
		[Column("CaseGuid")]
		[Display(Name = "案件GUID")]
		public Guid? CaseGuid { get; set; }

		/// <summary>
		/// 業務人員
		/// </summary>
		[Column("ToMemberID")]
		public string ToMemberID { get; set; }

		/// <summary>
		/// 照會日期
		/// </summary>
		[Column("CreateTime")]
		public string CreateTime { get; set; }

		/// <summary>
		/// 催辦日期
		/// </summary>
		[Column("DoSCreateTime")]
		public string DoSCreateTime { get; set; }

		/// <summary>
		/// 稽催日期
		/// </summary>
		[Column("AuditCreateTime")]
		public string AuditCreateTime { get; set; }

		/// <summary>
		/// 已留言對象
		/// </summary>
		[Column("AllToMemberID")]
		public string AllToMemberID { get; set; }

		/// <summary>
		/// 已留言對象
		/// </summary>
		[Column("AllCCMemberID")]
		public string AllCCMemberID { get; set; }

		/// <summary>
		/// 通知附件檔案
		/// </summary>
		[Column("CFID")]
		public string CFID { get; set; }

		/// <summary>附加檔案名稱</summary>
		[NonColumn]
		public string CRMFileName { get; set; }

		/// <summary>
		/// 附件檔
		/// </summary>
		[NonColumn]
		public List<CRMEFile> File { get; set; }

		/// <summary>設定成員名單給{Jason}</summary>
		[NonColumn]
		public string RecipientToJson { get; set; }

		/// <summary>
		/// 附檔名
		/// </summary>
		[NonColumn]
		public string UploadFilesName { get; set; }

		/// <summary>進度</summary>
		[Column("Seq")]
		public string Seq { get; set; }

		/// <summary>狀態</summary>
		[Column("State")]
		public string State { get; set; }

	}
}
