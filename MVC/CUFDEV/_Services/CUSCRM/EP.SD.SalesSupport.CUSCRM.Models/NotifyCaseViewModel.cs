using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using Microsoft.CUF.Framework.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM
{
	/// <summary>
	/// 通知作業立案通知Model
	/// </summary>
	public class NotifyCaseViewModel : IModel
	{
		/// <summary>
		/// 受理保單資料
		/// </summary>
		public List<NotifyInsuranceViewModel> NotifyInsurance { get; set; }

		/// <summary>
		/// 處代理人
		/// </summary>
		public List<NotifyInsuranceViewModel> ChinaOM { get; set; }

		/// <summary>
		/// 行專
		/// </summary>
		public List<Tuple<string, string>> NotifyEmployee { get; set; }

		/// <summary>
		/// 受理編號
		/// </summary>
		public string No { get; set; }

		/// <summary>
		/// 序號
		/// </summary>
		[Display(Name = "序號")]
		public int ViewID { get; set; }

		/// <summary>
		/// 經辦人員
		/// </summary>
		[Display(Name = "經辦人員")]
		public string DoUser { get; set; }

		/// <summary>
		/// 受文者Checkbox
		/// </summary>
		public string[] ToCheckbox { get; set; }

		/// <summary>
		/// 副本受文者Checkbox
		/// </summary>
		public string[] CCCheckbox { get; set; }

		/// <summary>
		/// 行專Checkbox
		/// </summary>
		public string[] EPCheckbox { get; set; }

		/// <summary>
		/// 處代理人Checkbox
		/// </summary>
		public string[] OMProxyCheckbox { get; set; }

		/// <summary>
		/// 經辦人員分機表
		/// </summary>
		[Display(Name = "經辦人員分機表")]
		public string DoUserTelExt { get; set; }

		/// <summary>經辦</summary>
		public string RecipientDoUserJson { get; set; }
		
		/// <summary>附加檔案名稱</summary>
		[Display(Name = "附加檔案名稱")]
		public string CRMFileName { get; set; }

		/// <summary>
		/// 附件檔
		/// </summary>
		[Display(Name = "附件檔")]
		public List<CRMEFile> File { get; set; }

		/// <summary>受文者</summary>
		public string RecipientToJson { get; set; }

		/// <summary>副本受文者</summary>
		public string RecipientCCJson { get; set; }

		/// <summary>行專</summary>
		public string RecipientEPJson { get; set; }

		/// <summary>
		/// 附檔名
		/// </summary>
		public string UploadFilesName { get; set; }

		/// <summary>
		/// 行專ORGID
		/// </summary>
		[Column("people_orgid")]
		public string PeopleOrgid { get; set; }

		/// <summary>
		/// 行專
		/// </summary>
		[Column("people_name")]
		public string PeopleName { get; set; }

		/// <summary>
		/// TabUniqueId 用來取得暫存資料夾的名稱
		/// </summary>
		[NonColumn]
		public string TabUniqueId { get; set; }

	}
}
