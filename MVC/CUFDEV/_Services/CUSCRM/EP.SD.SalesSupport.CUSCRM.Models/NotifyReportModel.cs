using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 照會單Model
	/// </summary>
	public class NotifyReportModel
	{
		/// <summary>
		/// 受理編號
		/// </summary>
		[Display(Name = "受理編號")]
		public string No { get; set; }

		/// <summary>
		/// 受文者
		/// </summary>
		public List<string> To { get; set; }

		/// <summary>
		/// 副本受文者
		/// </summary>
		public List<string> CC { get; set; }

		/// <summary>
		/// 催辦內容集合
		/// </summary>
		public List<string> listContent { get; set; }
		
		/// <summary>
		/// 處代理受文者
		/// </summary>
		public List<string> OM { get; set; }

		/// <summary>
		/// 
		/// </summary>
		public List<NotifyInsuranceViewModel> NotifyInsuranceViews { get; set; }

		/// <summary>
		/// 案件GUID
		/// </summary>
		[Display(Name = "案件GUID")]
		public Guid? CaseGuid { get; set; }

		/// <summary>
		/// 申訴人姓名
		/// </summary>
		[Display(Name = "申訴人姓名")]
		public string Complainant { get; set; }

		/// <summary>
		/// 資料來源內容
		/// </summary>
		[Display(Name = "資料來源內容")]
		public string SourceDesc { get; set; }

		/// <summary>
		/// 資料來源名稱
		/// </summary>
		[Display(Name = "資料來源")]
		public string SourceName { get; set; }

		/// <summary>
		/// 照會內容
		/// </summary>
		[Display(Name = "照會內容")]
		public string Content { get; set; }

		/// <summary>
		/// 業務連繫函
		/// </summary>
		[Display(Name = "業務連繫函")]
		public string BizContract { get; set; }

		/// <summary>
		/// 招攬服務報告書
		/// </summary>
		[Display(Name = "招攬服務報告書")]
		public string SolicitingRpt { get; set; }

		/// <summary>
		/// 經辦人員名稱
		/// </summary>
		[Display(Name = "經辦人員")]
		public string DoUserName { get; set; }

		/// <summary>
		/// 經辦人員分機表
		/// </summary>
		[Display(Name = "經辦人員分機表")]
		public string DoUserTelExt { get; set; }

		/// <summary>
		/// 建立時間
		/// </summary>
		[Display(Name = "建立時間")]
		public string CreateTime { get; set; }

		/// <summary>
		/// 照會時間+5工作日
		/// </summary>
		[Display(Name = "建立時間")]
		public string CreateFiveTime { get; set; }

		/// <summary>
		/// 催辦、稽催時間
		/// </summary>
		[Display(Name = "建立時間")]
		public string DoSAuditTime { get; set; }

		/// <summary>
		/// 催辦、稽催時間+N工作日
		/// </summary>
		[Display(Name = "建立時間")]
		public string DoSAuditAddTime { get; set; }

		/// <summary>
		/// 催辦、稽催次數
		/// </summary>
		public int DoSAuditCount { get; set; }

		/// <summary>
		/// 催辦內容
		/// </summary>
		public string DoSContent { get; set; }
	}
}
