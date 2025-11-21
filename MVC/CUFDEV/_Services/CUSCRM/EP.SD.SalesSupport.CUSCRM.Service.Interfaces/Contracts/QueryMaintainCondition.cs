using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
	/// <summary>
	/// 維護介面查詢條件
	/// </summary>
	public class QueryMaintainCondition
	{
		/// <summary>
		/// 案件類別
		/// </summary>
		[Display(Name = "類別")]
		public string Type { get; set; }

		/// <summary>
		/// 照會單號(受理編號)
		/// </summary>
		[Display(Name = "照會單號")]
		public string No { get; set; }

		/// <summary>
		/// 案件類型
		/// </summary>
		[Display(Name = "類型")]
		public string Kind { get; set; }

		/// <summary>
		/// 受理區間(起)
		/// </summary>
		[Display(Name = "受理區間")]
		public string DateS { get; set; }

		/// <summary>
		/// 受理區間(迄)
		/// </summary>
		[Display(Name = "受理區間")]
		public string DateE { get; set; }

		/// <summary>
		/// 保單號碼
		/// </summary>
		[Display(Name = "保單號碼")]
		public string PolicyNo { get; set; }

		/// <summary>
		/// 要保人姓名
		/// </summary>
		[Display(Name = "要保人姓名")]
		public string Owner { get; set; }

		/// <summary>
		/// 承辦人員
		/// </summary>
		[Display(Name = "承辦人員")]
		public string Creator { get; set; }

		/// <summary>
		/// 結案
		/// </summary>
		[Display(Name = "結案")]
		public bool IsClose { get; set; }

		/// <summary>
		/// 狀態
		/// </summary>
		[Display(Name = "狀態")]
		public StatusCategory Status { get; set; }
	}
}
