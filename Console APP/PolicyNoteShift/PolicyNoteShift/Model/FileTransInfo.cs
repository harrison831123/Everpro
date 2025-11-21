using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PolicyNoteShift.Model
{
	public class FileTransInfo
	{
		/// <summary> 
		/// 保險公司代碼        
		/// <summary> 
		public string CompanyCode { get; set; }

		/// <summary> 
		/// FuncId        
		/// <summary> 
		public string FuncId { get; set; }

		/// <summary> 
		/// FuncId說明        
		/// <summary> 
		public string FuncDesc { get; set; }

		/// <summary> 
		/// 程式/SP名稱    
		/// <summary> 
		public string ProgramName { get; set; }

		/// <summary> 
		/// 壓縮檔名        
		/// <summary> 
		public string ZipFileName { get; set; }

		/// <summary> 
		/// 壓縮檔密碼       
		/// <summary> 
		public string ZipPw { get; set; }

		/// <summary> 
		/// 檔案名稱
		/// <summary> 
		public string FileName { get; set; }

		/// <summary> 
		/// 檔案類型
		/// <summary> 
		public string FileNameExtension { get; set; }

		/// <summary> 
		/// 檔案編碼
		/// <summary> 
		public string FileEncoding { get; set; }

		/// <summary> 
		///  SP參數名稱        
		/// <summary> 
		public string SpParmatersName { get; set; }

		/// <summary> 
		/// 來源檔路徑        
		/// <summary> 
		public string SourcePath { get; set; }

		/// <summary> 
		///  來源檔Folder       
		/// <summary> 
		public string SourceFolder { get; set; }

		/// <summary> 
		/// 來源檔路徑帳號       
		/// <summary> 
		public string SourceAccount { get; set; }

		/// <summary> 
		/// 來源檔路徑密碼
		/// <summary> 
		public string SourcePassword { get; set; }

		/// <summary> 
		/// 來源檔備份路徑        
		/// <summary> 
		public string BackupPath { get; set; }

		/// <summary> 
		///  來源檔備份Folder       
		/// <summary> 
		public string BackupFolder { get; set; }

		/// <summary> 
		/// 來源檔備份路徑帳號        
		/// <summary> 
		public string BackupAccount { get; set; }

		/// <summary> 
		/// 來源檔備份路徑密碼        
		/// <summary> 
		public string BackupPassword { get; set; }

		/// <summary> 
		/// 產生比對成功檔案        
		/// <summary> 
		public string SuccessFile { get; set; }

		/// <summary> 
		/// 產生比對失敗檔案        
		/// <summary> 
		public string FailFile { get; set; }

		/// <summary> 
		/// 執行時間        
		/// <summary> 
		public string ExecTime { get; set; }

		/// <summary> 
		/// 檔案資料存放表        
		/// <summary> 
		public string UserDefineTableName { get; set; }
	}
}
