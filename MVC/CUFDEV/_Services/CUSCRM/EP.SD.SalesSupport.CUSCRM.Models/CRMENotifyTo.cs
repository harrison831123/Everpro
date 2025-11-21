using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
    /// <summary>
    /// 立案通知對象
    /// </summary>
	[Table("CRMENotifyTo")]
	public class CRMENotifyTo : IModel
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
		/// 通知對象類別
		/// </summary>
		[Column("NotifyType")]
		[Display(Name = "通知對象類別")]
		public NotifyType? NotifyType { get; set; }

		/// <summary>

		/// </summary>
		[Column("MemberID")]

		public string MemberID { get; set; }

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

	}

}
