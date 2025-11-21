using Microsoft.CUF.Framework.Data;
using System;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 維護附件
	/// </summary>
	[DataContract]
	[Table("CRMEFile")]
	public class CRMEFile : IModel
	{

		/// <summary>
		/// 自動編號
		/// </summary>
		[DataMember]
		[Column("ID", IsIdentity = true)]
		[Display(Name = "自動編號")]
		public int ID { get; set; }

		/// <summary>
		/// 附件來源類別
		/// </summary>
		[DataMember]
		[Column("SourceType")]
		[Display(Name = "附件來源類別")]
		public int? SourceType { get; set; }

		/// <summary>
		/// 目錄編號（受理編號）
		/// </summary>
		[DataMember]
		[Column("FolderNo")]
		[Display(Name = "目錄編號（受理編號）")]
		public string FolderNo { get; set; }

		/// <summary>
		/// 目錄編號（Connector）
		/// </summary>
		[DataMember]
		[Column("FileNo")]
		[Display(Name = "目錄編號（Connector）")]
		public string FileNo { get; set; }

		/// <summary>
		/// 上傳檔案名稱
		/// </summary>
		[DataMember]
		[Column("FileName")]
		[Display(Name = "上傳檔案名稱")]
		public string FileName { get; set; }

		/// <summary>
		/// 上傳檔案編碼
		/// </summary>
		[DataMember]
		[Column("FileMD5Name")]
		[Display(Name = "上傳檔案編碼")]
		public string FileMD5Name { get; set; }

		/// <summary>
		/// 上傳者員編
		/// </summary>
		[DataMember]
		[Column("Creator")]
		[Display(Name = "上傳者員編")]
		public string Creator { get; set; }

		/// <summary>
		/// 維護建立時間
		/// </summary>
		[DataMember]
		[Column("CreateTime")]
		[Display(Name = "維護建立時間")]
		public DateTime? CreateTime { get; set; }

		/// <summary>
		/// 讀取狀態
		/// </summary>
		[DataMember]
		[Column("FileRead")]
		[Display(Name = "讀取狀態")]
		public int? FileRead { get; set; }

		/// <summary>
		/// 讀取時間
		/// </summary>
		[DataMember]
		[Column("FileReaddate")]
		[Display(Name = "讀取時間")]
		public DateTime? FileReaddate { get; set; }

	}

}
