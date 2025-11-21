using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 案件內容
	/// </summary>
	[Table("CRMECaseContent")]
	public class CRMECaseContent : IModel
	{

		/// <summary>
		/// 自動編號
		/// </summary>
		[Column("ID", IsIdentity = true)]
		[Display(Name = "自動編號")]
		public int ID { get; set; }

		/// <summary>
		/// 受理編號
		/// </summary>
		[Column("No")]
		[Display(Name = "受理編號")]
		public string No { get; set; }

		/// <summary>
		/// 案件GUID
		/// </summary>
		[Column("CaseGuid")]
		[Display(Name = "案件GUID")]
		public Guid? CaseGuid { get; set; }

		/// <summary>
		/// 類別
		/// </summary>
		[Column("Type")]
		[Display(Name = "類別")]
		public string Type { get; set; }

		/// <summary>
		/// 資料來源
		/// </summary>
		[Column("SourceID")]
		[Display(Name = "資料來源")]
		public int? SourceID { get; set; }

		/// <summary>
		/// 資料來源內容
		/// </summary>
		[Column("SourceDesc")]
		[Display(Name = "資料來源內容")]
		public string SourceDesc { get; set; }

		/// <summary>
		/// 保險公司代碼
		/// </summary>
		[Column("CompanyCode")]
		[Display(Name = "保險公司代碼")]
		public string CompanyCode { get; set; }

		/// <summary>
		/// 來電者
		/// </summary>
		[Column("CallerID")]
		[Display(Name = "來電者")]
		public int? CallerID { get; set; }

		/// <summary>
		/// 案件類別
		/// </summary>
		[Column("CaseTypeID")]
		[Display(Name = "案件類別")]
		public int? CaseTypeID { get; set; }

		/// <summary>
		/// 案件類型
		/// </summary>
		[Column("CaseCategoryID")]
		[Display(Name = "案件類型")]
		public int? CaseCategoryID { get; set; }

		/// <summary>
		/// 收文日期
		/// </summary>
		[Column("ReceiveDateTime")]
		[Display(Name = "收文日期")]
		public DateTime? ReceiveDateTime { get; set; }

		/// <summary>
		/// 回文期限
		/// </summary>
		[Column("ReplayDDLDateTime")]
		[Display(Name = "回文期限")]
		public DateTime? ReplayDDLDateTime { get; set; }

		/// <summary>
		/// 照會內容
		/// </summary>
		[Column("Content")]
		[Display(Name = "照會內容")]
		public string Content { get; set; }

		/// <summary>
		/// 業務連繫函
		/// </summary>
		[Column("BizContract")]
		[Display(Name = "業務連繫函")]
		public string BizContract { get; set; }

		/// <summary>
		/// 招攬服務報告書
		/// </summary>
		[Column("SolicitingRpt")]
		[Display(Name = "招攬服務報告書")]
		public string SolicitingRpt { get; set; }

		/// <summary>
		/// 申訴人姓名
		/// </summary>
		[Column("Complainant")]
		[Display(Name = "申訴人姓名")]
		public string Complainant { get; set; }

		/// <summary>
		/// 案件狀態
		/// </summary>
		[Column("Status")]
		[Display(Name = "案件狀態")]
		public int? Status { get; set; }

		/// <summary>
		/// 經辦人員
		/// </summary>
		[Column("DoUser")]
		[Display(Name = "經辦人員")]
		public string DoUser { get; set; }

		/// <summary>
		/// 經辦人員分機表
		/// </summary>
		[Column("DoUserTelExt")]
		[Display(Name = "經辦人員分機表")]
		public string DoUserTelExt { get; set; }

		/// <summary>
		/// 處理結果(單選)
		/// </summary>
		[Column("ProcResultRadio")]
		[Display(Name = "處理結果")]
		public int? ProcResultRadio { get; set; }

		/// <summary>
		/// 處理結果(複選)
		/// </summary>
		[Column("ProcResultCheckBox")]
		[Display(Name = "處理結果")]
		public string ProcResultCheckBox { get; set; }

		/// <summary>
		/// 線上回覆
		/// </summary>
		[Column("OnlineReplay")]
		[Display(Name = "線上回覆")]
		public string OnlineReplay { get; set; }

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
		/// 類別
		/// </summary>
		[NonColumn]
		[Display(Name = "類別")]
		public string TypeName { get; set; }

		/// <summary>
		/// 保險公司名稱
		/// </summary>
		[NonColumn]
		[Display(Name = "保險公司")]
		public string CompanyName { get; set; }

		/// <summary>
		/// 資料來源名稱
		/// </summary>
		[NonColumn]
		[Display(Name = "資料來源")]
		public string SourceName { get; set; }

		/// <summary>
		/// 來電者名稱
		/// </summary>
		[NonColumn]
		[Display(Name = "來電者")]
		public string CallerName { get; set; }

		/// <summary>
		/// 案件類別名稱
		/// </summary>
		[NonColumn]
		[Display(Name = "案件類別")]
		public string CaseTypeName { get; set; }

		/// <summary>
		/// 案件類型名稱
		/// </summary>
		[NonColumn]
		[Display(Name = "案件類型")]
		public string CaseCategoryName { get; set; }

		/// <summary>
		/// 經辦人員名稱
		/// </summary>
		[NonColumn]
		[Display(Name = "經辦人員")]
		public string DoUserName { get; set; }

		/// <summary>
		/// 建立人員名稱
		/// </summary>
		[NonColumn]
		[Display(Name = "建立人員")]
		public string CreatorName { get; set; }

	}
}
