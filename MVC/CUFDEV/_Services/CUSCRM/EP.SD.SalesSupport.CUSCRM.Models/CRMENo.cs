using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
    /// <summary>
    /// 受理編號
    /// </summary>
    [DataContract]
	[Table("CRMENo")]
	public class CRMENo : IModel
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
		/// 受理類型代碼
		/// </summary>
		[Column("Code")]
		[Display(Name = "受理類型代碼")]
		public string Code { get; set; }

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
