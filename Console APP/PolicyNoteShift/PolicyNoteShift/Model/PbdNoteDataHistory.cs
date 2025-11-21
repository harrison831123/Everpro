using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PolicyNoteShift.Model
{
	/// <summary>
	/// 保公照會檔處理歷史檔
	/// </summary>
	public class PbdNoteDataHistory
	{
		/// <summary>
		/// 照會類別
		/// </summary>
		public string note_type { get; set; }

		/// <summary>
		/// 要保書受理序號
		/// </summary>
		public string po_serial { get; set; }

		/// <summary>
		/// 保單號碼
		/// </summary>
		public string policy_no { get; set; }

		/// <summary>
		/// 照會日期
		/// </summary>
		public string notice_date { get; set; }

		/// <summary>
		/// 照會回覆期限
		/// </summary>
		public string replay_date { get; set; }

		/// <summary>
		/// 資料內容序號
		/// </summary>
		public string content_seq { get; set; }

		/// <summary>
		/// 業務員登錄字號
		/// </summary>
		public string agent_license_no { get; set; }

		/// <summary>
		/// 附檔PDF檔名
		/// </summary>
		public string note_pdf_name { get; set; }

		/// <summary>
		/// 原始壓縮檔名
		/// </summary>
		public string zipfile_name { get; set; }

		/// <summary>
		/// 保險公司代碼
		/// </summary>
		public string company_code { get; set; }

		/// <summary>
		/// 掃描受理序號
		/// </summary>
		public string policy_serial { get; set; }

		/// <summary>
		/// 轉檔資料檢核處理結果
		/// </summary>
		public string result_flag { get; set; }

		/// <summary>
		/// 轉檔資料異常描述
		/// </summary>
		public string result_desc { get; set; }

		/// <summary>
		/// 轉檔處理時間
		/// </summary>
		public DateTime? batch_datetime { get; set; }

		/// <summary>
		/// 轉入前端照會處理結果
		/// </summary>
		public string trans2ep_flag { get; set; }

		/// <summary>
		/// 轉入前端照會處理時間
		/// </summary>
		public DateTime? trans2ep_datetime { get; set; }

		/// <summary>
		/// 序號
		/// </summary>
		public int Iden { get; set; }

	}




}
