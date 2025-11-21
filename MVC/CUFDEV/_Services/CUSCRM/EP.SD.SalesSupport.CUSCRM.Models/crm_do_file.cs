using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
   
    
	/// <summary>
	/// 聯繫紀錄檔案
	/// </summary>
	[DataContract]
	[Table("crm_do_file")]
	public class crm_do_file : IModel
	{
		/// <summary>
		/// 自動編號
		/// </summary>
		[DataMember]
		[Column("id", IsIdentity = true)]
		[Display(Name = "自動編號")]
		public int id { get; set; }

		/// <summary>
		/// 受理編號
		/// </summary>
		[DataMember]
		[Column("crm_no")]
		[Display(Name = "受理編號")]
		public string crm_no { get; set; }
		/// <summary>
		/// 原始檔名
		/// </summary>
		[DataMember]
		[Column("crm_filename")]
		[Display(Name = "原始檔名")]
		public string crm_filename { get; set; }
		/// <summary>
		/// 原始檔名
		/// </summary>
		[DataMember]
		[Column("crm_md5name")]
		[Display(Name = "系統檔名")]
		public string crm_md5name { get; set; }



	}
}