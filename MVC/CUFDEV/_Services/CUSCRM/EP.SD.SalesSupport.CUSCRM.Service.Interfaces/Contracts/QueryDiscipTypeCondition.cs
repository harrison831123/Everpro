using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    /// <summary>
    /// 查詢資料設定的條件
    /// </summary>
    public class QueryDiscipTypeCondition
    {
		/// <summary>
		/// 代碼
		/// </summary>
		[Display(Name = "代碼")]
		public DiscipTypeCode? Code { get; set; }

		/// <summary>
		/// 代碼名稱
		/// </summary>
		[Display(Name = "代碼名稱")]
		public string Name { get; set; }

		/// <summary>
		/// 狀態
		/// </summary>
		[Display(Name = "狀態")]
		public EnableStatus? Status { get; set; }

		/// <summary>
		/// 類別
		/// </summary>
		[Display(Name = "類別")]
		public DiscipTypeKind? Kind { get; set; }
	}
}
