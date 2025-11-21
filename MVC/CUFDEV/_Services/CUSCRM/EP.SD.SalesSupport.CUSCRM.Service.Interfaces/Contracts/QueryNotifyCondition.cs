using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM.Service.Contracts
{
	/// <summary>
	/// 通知作業查詢條件
	/// </summary>
	public class QueryNotifyCondition
	{
		/// <summary>
		/// 受理編號
		/// </summary>
		public string No { get; set; }

		/// <summary>
		/// 案件狀態
		/// </summary>
		public int Status { get; set; }

		/// <summary>
		/// 類型
		/// </summary>
		public string Type { get; set; }

		/// <summary>
		/// 結案
		/// </summary>
		public string CloseCase { get; set; }

		/// <summary>
		/// 類別
		/// </summary>
		public string Category { get; set; }

		/// <summary>
		/// 業務員ID
		/// </summary>
		public string AgentCode { get; set; }

		/// <summary>
		/// 人員ID
		/// </summary>
		public string MemberID { get; set; }

		/// <summary>
		/// 案件GUID
		/// </summary>
		public Guid? CaseGuid { get; set; }

	}
}
