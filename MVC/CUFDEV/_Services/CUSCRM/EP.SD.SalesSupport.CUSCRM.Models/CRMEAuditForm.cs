using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 稽催通知單
	/// </summary>
	[Table("CRMEAuditForm")]
	public class CRMEAuditForm : IModel
	{

		/// <summary>
		/// 自動編號
		/// </summary>
		[Column("ID", IsIdentity = true)]
		[Display(Name = "自動編號")]
		public int ID { get; set; }

		/// <summary>
		/// 稽催記錄編號
		/// </summary>
		[Column("AuditId")]
		[Display(Name = "稽催記錄編號")]
		public long AuditId { get; set; }

		/// <summary>
		/// 受理號碼
		/// </summary>
		[Column("No")]
		[Display(Name = "受理號碼")]
		public string No { get; set; }

		/// <summary>
		/// 通知單類別
		/// </summary>
		[Column("AuditFormType")]
		[Display(Name = "通知單類別")]
		public int? AuditFormType { get; set; }

		/// <summary>
		/// 敬會者員編
		/// </summary>
		[Column("AuditFormTo")]
		[Display(Name = "敬會者員編")]
		public string AuditFormTo { get; set; }

		/// <summary>
		/// 受理日期
		/// </summary>
		[Column("AuditFormStart")]
		[Display(Name = "受理日期")]
		public DateTime? AuditFormStart { get; set; }

		/// <summary>
		/// 回覆日期
		/// </summary>
		[Column("AuditFormEnd")]
		[Display(Name = "回覆日期")]
		public DateTime? AuditFormEnd { get; set; }

		/// <summary>
		/// 行銷過程報告
		/// </summary>
		[Column("AuditFormData1")]
		[Display(Name = "行銷過程報告")]
		public string AuditFormData1 { get; set; }

		/// <summary>
		/// 附件文宣
		/// </summary>
		[Column("AuditFormData2")]
		[Display(Name = "附件文宣")]
		public string AuditFormData2 { get; set; }

		/// <summary>
		/// 經辦
		/// </summary>
		[Column("AuditFormDo")]
		[Display(Name = "經辦")]
		public string AuditFormDo { get; set; }

		/// <summary>
		/// 經辦分機
		/// </summary>
		[Column("AuditFormTel")]
		[Display(Name = "經辦分機")]
		public string AuditFormTel { get; set; }

		/// <summary>
		/// 稽催人員員編
		/// </summary>
		[Column("AuditFormCreator")]
		[Display(Name = "稽催人員員編")]
		public string AuditFormCreator { get; set; }

		/// <summary>
		/// 稽催單建立時間
		/// </summary>
		[Column("AuditFormCreaterDate")]
		[Display(Name = "稽催單建立時間")]
		public DateTime AuditFormCreaterDate { get; set; }

	}

}
