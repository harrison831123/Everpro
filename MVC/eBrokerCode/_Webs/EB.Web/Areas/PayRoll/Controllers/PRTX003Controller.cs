using EB.Common;
using EB.EBrokerModels;
using EB.SL.PayRoll.Models;
using EB.SL.PayRoll.Service;
using EB.SL.PayRoll.Web.Areas.PayRoll.Utilities;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace EB.SL.PayRoll.Web.Areas.PayRoll.Controllers
{
	[Program("PRTX003")]
	public class PRTX003Controller : BaseController
	{
		private IPayRollService _service;
		private static string _programID = "PRTX003";
		public PRTX003Controller()
		{
			_service = ServiceHelper.Create<IPayRollService>();
		}
		// GET: PayRoll/PRTX003
		[HasPermission("EB.SL.PayRoll.PRTX003")]
		public ActionResult Index()
		{
			return View();
		}

		/// <summary>
		/// 上傳調整資料
		/// </summary>
		/// <param name="file">上傳的excel</param>
		/// <param name="productionYM">工作年月</param>
		/// <param name="sequence">N次佣</param>
		/// <returns></returns>
		[HttpPost]
		[HasPermission("EB.SL.PayRoll.PRTX003")]
		public ActionResult Upload(HttpPostedFileBase file, string productionYM, string sequence, bool termination)
		{
			List<string> errorList = new List<string>();  //記錄錯誤
			List<string> warnList = new List<string>();  //記錄可匯入系統之錯誤
			string check = string.Empty;
			try
			{


				//檢查檔案大小
				var fileSize = file.ContentLength;
				if (fileSize > 0 &&
					(fileSize / 1048576.0) < Int32.Parse(EB.Web.PlatformHelper.GetVarConfig("UploadLimit").ToString()))
				{
					//重新給檔名
					string extension = Path.GetExtension(file.FileName);
					var splitName = file.FileName.Split(new string[] { extension }, StringSplitOptions.RemoveEmptyEntries);
					string saveFileName = string.Format("{0}-{1}{2}", splitName[0], DateTime.Now.ToString("yyyyMMddHHmmssfff"), extension);
					string dirPath = string.Format("{0}/{1}/{2}/", EB.Web.PlatformHelper.GetVarConfig("PayRollBatchAdjustDir"), productionYM.Replace("/", ""), User.AccountInfo.ID);
					string csvDirPath = "";

					csvDirPath = Path.Combine(dirPath, saveFileName.Replace(extension, ""));
					string savePath = Path.Combine(dirPath, saveFileName);

					if (!Directory.Exists(dirPath))
					{
						Directory.CreateDirectory(dirPath);
					}

					file.SaveAs(savePath);

					//處理csv folder
					if (!Directory.Exists(csvDirPath))
					{
						Directory.CreateDirectory(csvDirPath);
					}

					IWorkbook workbook = null;
					using (FileStream openFile = new FileStream(savePath, FileMode.Open, FileAccess.Read))
					{
						//2007版本  
						if (savePath.IndexOf(".xlsx") > 0)
							workbook = new XSSFWorkbook(openFile);
					}

					//csv轉換處理
					string csvTargetFile = Path.Combine(csvDirPath, saveFileName.Replace(extension, "") + ".csv");
					DataTable csvDT = EPBKS.Common.EPPlus.EPPlusHelper.GetDataTableFromExcel(savePath, workbook.GetSheetName(0), 0);
					CsvHelper.DataTableToCSV(csvDT, csvTargetFile);
					csvDT.TableName = "AgentBonusAdjust";


					#region 設定datable column名稱
					//set column name
					#region  舊順序           
					//型態[0] 保公代碼[1] 業務員代碼[2] 姓名[3] 調整原因碼[4] 保單號碼[5] 金額[6] FYC[7] FYP[8] 說明[9] 險種代碼[10] 年期[11] 繳別繳次[12] 計傭保費[13] 來源資料表鍵值[14]
					#endregion
					//型態[0] 保公代碼[1] 業務員代碼[2] 姓名[3] 調整原因碼[4] 保單號碼[5] 險種代碼[6] 年期[7] 繳別繳次[8] 金額[9] FYC[10] FYP[11]  說明[12] 來源資料表鍵值[13]
					csvDT.Columns[0].ColumnName = "adj_type";
					csvDT.Columns[1].ColumnName = "company_code";
					csvDT.Columns[2].ColumnName = "agent_code";
					csvDT.Columns[3].ColumnName = "agent_name";
					csvDT.Columns[4].ColumnName = "reason_code";
					csvDT.Columns[5].ColumnName = "policy_no2";
					csvDT.Columns[6].ColumnName = "plan_code";
					csvDT.Columns[7].ColumnName = "collect_year";
					csvDT.Columns[8].ColumnName = "modx_sequence";
					csvDT.Columns[9].ColumnName = "amount";
					csvDT.Columns[10].ColumnName = "fyc";
					csvDT.Columns[11].ColumnName = "fyp";
					//csvDT.Columns[12].ColumnName = "comm_mode_prem";
					csvDT.Columns[12].ColumnName = "remark";
					csvDT.Columns[13].ColumnName = "source_table_key";
					//----add column
					csvDT.Columns.Add("guid", typeof(Guid));//[14]
					csvDT.Columns.Add("production_ym", typeof(String));     //[15]
					csvDT.Columns.Add("process_no", typeof(String));        //[16]
					csvDT.Columns.Add("process_code", typeof(String));      //[17]
					csvDT.Columns.Add("AgentBonusDesc_guid", typeof(Guid));  //[18]
					csvDT.Columns.Add("create_datetime", typeof(DateTime));   //[19]
					csvDT.Columns.Add("create_user_code", typeof(String));  //[20]
					csvDT.Columns.Add("update_datetime", typeof(DateTime));   //[21]
					csvDT.Columns.Add("update_user_code", typeof(String));  //[22]
					csvDT.Columns.Add("sequence", typeof(String));  //[23]
					csvDT.Columns.Add("create_unit", typeof(String));  //[24]
					csvDT.Columns.Add("create_unit_name", typeof(String));  //[25]
					csvDT.Columns.Add("source_table", typeof(String));  //[26]
					csvDT.Columns.Add("comm_mode_prem", typeof(String));  //[27]
					#endregion

					#region 檢查資料內容
					List<AgentBonusDesc> descs = new List<AgentBonusDesc>();

					//List<AGWCSet> agwcsetB = _service.GetAGWCSetB(productionYM);  //取所有月備份實駐
					Dictionary<string, string> aginb = _service.GetAginbAgentCodeAndAgentNameByproductionYM(PayRollHelper.WYToRoc(productionYM));  //取所有月備份業務員ID與姓名
					Dictionary<string, string> orgin = _service.GetOrgin();
					List<GroupMapping> groups = _service.GetALLGroupMappingList();  //取所有原因碼與型態對應表
																					//List<company_set> companySets = _service.GetCompanySetList();  //取得所有保險公司代碼
					List<string> notinlist = _service.GetGroupMappingTOnotinList();  //取得R2、N4及NB原因碼不可人工調整
					List<string> minuslist = _service.GetGroupMappingTOminusList();  //取得金額應為負數的原因碼
					List<string> pluslist = _service.GetGroupMappingTOplusList();  //取得金額應為正數的原因碼
					List<string> agoccpind3List = _service.GetOrgibByagentcodeList(PayRollHelper.WYToRoc(productionYM), sequence);  //查詢任職型態3(內勤)人員
					List<string> FDstatus1List = _service.GetaginbByagentcodeList(PayRollHelper.WYToRoc(productionYM), sequence);  //查詢所有免評估人員
					List<string> modxseqList = _service.GetGroupMappingTOmodxseqList();  //取得繳别繳次必須要有內容的原因碼
					List<BlackList> blacklist = _service.GetBonusReserveSettingList(productionYM, sequence); //黑名單
					var PoagList = new List<Poag>();
					foreach (DataRow row in csvDT.Rows)
					{
						if (row.ItemArray[0].ToString().Equals("6"))
						{
							PoagList = _service.GetAgentNameByPoagList();  //poag保單
							break;
						}
					}
					List<string> sixList = _service.GetGroupMappingTOsixList();  //型態6原因馬
																				 //List<string> PoagList =  _service.GetAgentNameByPoagList();  //poag保單
					var firstyearList = _service.GetGroupMappingTOfirstyearList();  //初年度服務報酬
					var FYCList = _service.GetGroupMappingTOFYC();  //FYC
					var FYPList = _service.GetGroupMappingTOFYP();  //FYP
					var YearList = _service.GetGroupMappingTOYear();  //年期
					var PCList = _service.GetGroupMappingTOPlanCode();  //險種代碼
					var AFList = _service.GetGroupMappingTOAF();  //金額大於保費
					List<string> quits = _service.GetAginBQuitAG(PayRollHelper.WYToRoc(productionYM));

					int i = 2;//第1列為標題，資料列從第2列開始
					foreach (DataRow row in csvDT.Rows)
					{
						//型態[0] 保公代碼[1] 業務員代碼[2] 姓名[3] 調整原因碼[4] 保單號碼[5] 險種代碼[6] 年期[7] 繳別繳次[8] 金額[9] FYC[10] FYP[11] 說明[12] 來源資料表鍵值[13]

						//業務員代碼去頭尾空白
						row.SetField(2, row.ItemArray[2].ToString().TrimEnd().TrimStart());
						//業務員姓名去頭尾空白
						row.SetField(3, row.ItemArray[3].ToString().TrimEnd().TrimStart());
						//險種去頭尾空白
						row.SetField(6, row.ItemArray[6].ToString().TrimEnd().TrimStart());
						//繳別繳次去頭尾空白
						row.SetField(8, row.ItemArray[8].ToString().TrimEnd().TrimStart());

						string type1 = groups.Where(x => x.MappingSource.Equals(row.ItemArray[4].ToString().TrimEnd())).Select(x => x.MappingValue).FirstOrDefault();
						if (type1 == null || type1.TrimEnd().Equals(""))
						{
							errorList.Add(string.Format("檢測到資料列第{0}列，調整原因碼 無法對出正確的 型態，請檢查 調整原因碼 {1} 和 型態 {2}。", i, row.ItemArray[4].ToString(), row.ItemArray[0].ToString()));
						}
						else if(row.ItemArray[0].ToString() != type1.TrimEnd())
						{
							if (row.ItemArray[0].ToString() == "4" || row.ItemArray[0].ToString() == "8")
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，型態有誤，請檢查 調整原因碼 {1} 和 型態 {2}。", i, row.ItemArray[4].ToString(), row.ItemArray[0].ToString()));
							}
						}

						if ((row.ItemArray[4].ToString() == "01" || row.ItemArray[4].ToString() == "05") && (row.ItemArray[0].ToString() == "8"))
						{
							errorList.Add(string.Format("檢測到資料列第{0}列，型態有誤，請確認。", i));
						}

						if (!User.HasPermission("EB.SL.PayRoll.PRTX002.Admin"))
						{
							if (row.ItemArray[0].ToString() == "4" || row.ItemArray[0].ToString() == "6")
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，無權限執行型態4及型態6", i));
							}
						}

						//R2、N4及NB，此原因碼不可人工調整
						//check = _service.GetGroupMappingTOnotin(row.ItemArray[4].ToString());
						if (notinlist.Contains(row.ItemArray[4].ToString()))
						{
							errorList.Add(string.Format("檢測到資料列第{0}列，此原因碼不可人工調整，請確認 調整原因碼。錯誤的值為：{1}", i, row.ItemArray[4].ToString()));
						}

						//金額不可空
						if (row.ItemArray[9].ToString().TrimEnd().Equals(""))
						{
							//row.SetField(9, 0);
							errorList.Add(string.Format("檢測到資料列第{0}列，金額 不可為空，請重新檢查。", i));
						}

						if (row.ItemArray[0].ToString().Equals("8"))
						{
							//保公代碼需為空
							if (!row.ItemArray[1].ToString().TrimEnd().Equals(""))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，保險公司代碼 需為空值，請重新檢查。", i));
							}

							//繳別繳次需為空
							if (!row.ItemArray[8].ToString().TrimEnd().Equals(""))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，繳別繳次 需為空值，請重新檢查。", i));
							}

							//險種代碼需為空
							if (!row.ItemArray[6].ToString().TrimEnd().Equals(""))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，險種代碼 需為空值，請重新檢查。", i));
							}

							//保單號碼需為空
							if (!row.ItemArray[5].ToString().TrimEnd().Equals(""))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，保單號碼 需為空值，請重新檢查。", i));
							}

							//年期需為空
							if (!row.ItemArray[7].ToString().TrimEnd().Equals(""))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，年期 需為空值，請重新檢查。", i));
							}
						}

						//FYC 不可空
						if (FYCList.Contains(row.ItemArray[4].ToString()))
						{
							if (row.ItemArray[10].ToString().TrimEnd().Equals(""))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，FYC 不可空白，請確認。", i));
							}
						}
						else
						{
							if (row.ItemArray[10].ToString().TrimEnd().Equals(""))
							{
								row.SetField(10, 0);
							}
						}

						//FYP不可空
						if (FYPList.Contains(row.ItemArray[4].ToString()))
						{
							if (row.ItemArray[11].ToString().TrimEnd().Equals(""))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，FYP 不可空白，請確認。", i));
							}
						}
						else
						{
							if (row.ItemArray[11].ToString().TrimEnd().Equals(""))
							{
								row.SetField(11, 0);
							}
						}

						//年期不可空
						if (YearList.Contains(row.ItemArray[4].ToString()))
						{
							if (row.ItemArray[7].ToString().TrimEnd().Equals(""))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，年期 不可空白，請確認。", i));
							}
						}
						else
						{
							if (row.ItemArray[7].ToString().TrimEnd().Equals(""))
							{
								row.SetField(7, 0);
							}
						}

						//險種代碼不可空
						if (PCList.Contains(row.ItemArray[4].ToString()))
						{
							if (row.ItemArray[6].ToString().TrimEnd().Equals(""))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，險種代碼 不可為空值，請確認。", i));
							}
						}
						else
						{
							if (!row.ItemArray[6].ToString().TrimEnd().Equals(""))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，險種代碼 應為空值，請確認。", i));
							}
						}

						//金額不得大於保費
						if (AFList.Contains(row.ItemArray[4].ToString()))
						{
							if (PayRollHelper.ABS(Convert.ToInt32(row.ItemArray[9])) > PayRollHelper.ABS(Convert.ToInt32(row.ItemArray[11])))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，金額不得大於保費，請重新檢查。", i));
							}
						}

						//原因碼AD、AE、AU、AW、AY、T2、T4、T6限負數
						//check = _service.GetGroupMappingTOminus(row.ItemArray[4].ToString());
						if (minuslist.Contains(row.ItemArray[4].ToString()))
						{
							if (Convert.ToInt32(row.ItemArray[9]) > 0)
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，此為負項項目，金額應為負數。錯誤的值為：{1}", i, row.ItemArray[9].ToString()));
							}
						}

						//原因碼AF、AG、AT、AV、AX、T1、T3、T5限正數
						//check = _service.GetGroupMappingTOplus(row.ItemArray[4].ToString());
						if (pluslist.Contains(row.ItemArray[4].ToString()))
						{
							if (Convert.ToInt32(row.ItemArray[9]) < 0)
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，此為正項項目，金額應為正數。錯誤的值為：{1}", i, row.ItemArray[9].ToString()));
							}
						}

						//string Agoccpind = _service.GetOrgibByagentcode(row.ItemArray[2].ToString());
						//有內勤人員
						if (row.ItemArray[4].ToString() == "AT")
						{
							if (agoccpind3List.Contains(row.ItemArray[2].ToString()))
							{
								warnList.Add(string.Format("檢測到資料列第{0}列，有內勤人員，請確認。", i));
							}
						}

						//免評估人員
						if (row.ItemArray[4].ToString() == "0A")
						{
							//string FDStatus = _service.GetaginbByagentcode(row.ItemArray[2].ToString(), PayRollHelper.WYToRoc(productionYM), sequence);
							if (FDstatus1List.Contains(row.ItemArray[2].ToString()))
							{
								//errorList.Add(string.Format("檢測到資料列第{0}列，有內勤人員，請確認。錯誤的值為：{1}", i, row.ItemArray[2].ToString()));
								warnList.Add(string.Format("檢測到資料列第{0}列，發放人員為免評估，請確認。", i));
							}
						}

						//有黑名單人員
						string agent = blacklist.Where(x => x.AgentCode.Equals(row.ItemArray[2].ToString().TrimEnd())).Select(x => x.AgentCode).FirstOrDefault();
						if (!String.IsNullOrEmpty(agent))
						{
							warnList.Add(string.Format("檢測到資料列第{0}列，有黑名單人員，請確認。", i));
						}

						if (row.ItemArray[2].ToString().Equals(""))
						{
							errorList.Add(string.Format("檢測到資料列第{0}列，業務員代碼 不可空白，請確認。", i));
						}
						else
						{
							//業務員代碼不存在當月關檔資料
							if (!aginb.ContainsKey(row.ItemArray[2].ToString()))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，業務員代碼不存在當月關檔資料，請確認業務員代碼。錯誤的值為：{1}", i, row.ItemArray[2].ToString()));
							}
							//確認簽約日期是否大於業績年月
							orgin.TryGetValue(row.ItemArray[2].ToString(),out string RegisterDate);
							int result = DateTime.Compare(Convert.ToDateTime(RegisterDate), Convert.ToDateTime(productionYM));
							if (result == 1)
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，業務員簽約日期大於業績年月，請確認業務員代碼。錯誤的值為：{1}", i, row.ItemArray[2].ToString()));
							}
						}

						//業務員任職型態3(內勤)人員，FYC需為0
						if (row.ItemArray[2].ToString() != "")
						{
							//string Agoccpind = _service.GetOrgibByagentcode(row.ItemArray[2].ToString());
							if (agoccpind3List.Contains(row.ItemArray[2].ToString()))
							{
								if (row.ItemArray[10].ToString() != "0")
								{
									errorList.Add(string.Format("檢測到資料列第{0}列，業務員任職型態3(內勤)人員，FYC需為0", i));
								}
							}
						}

						//檢查保險公司代碼
						if (row.ItemArray[1].ToString() != "")
						{
							if (!PayRollHelper.CheckCompanyCodeIsExist(row.ItemArray[1].ToString()))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，保險公司代碼不存在，請確認。錯誤的值為：{1}", i, row.ItemArray[1].ToString()));
							}
						}

						//來源資料表鍵值:若非空白或數值，請顯示警示訊息『來源資料表鍵值，請確認』，但可入系統。						
						if (!row.ItemArray[13].ToString().Equals(""))
						{
							warnList.Add(string.Format("檢測到資料列第{0}列，來源資料表鍵值 非空白，請確認。", i));

							bool isNumeric = int.TryParse(row.ItemArray[13].ToString(), out int r);
							if (!isNumeric)
							{
								warnList.Add(string.Format("檢測到資料列第{0}列，來源資料表鍵值 非數值，請確認。", i));
							}
						}						

						//繳別繳次
						//check = _service.GetGroupMappingTOmodxseq(row.ItemArray[4].ToString());
						if (modxseqList.Contains(row.ItemArray[4].ToString()))
						{
							if (String.IsNullOrEmpty(row.ItemArray[8].ToString()))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，此原因碼繳别繳次必須要有內容，請重新檢查。", i));
							}
							else if (row.ItemArray[8].ToString().Length != 5)
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，繳別繳次格式錯誤，請重新檢查。錯誤的值為：{1}", i, row.ItemArray[8].ToString()));
							}
						}
						else
						{
							row.SetField(8, "");
						}
						Regex rgx = new Regex(@"^\d{4}[MSQDY]$");
						if (!string.IsNullOrEmpty(row.ItemArray[8].ToString()))
						{
							if (rgx.IsMatch(row.ItemArray[8].ToString()) == false)
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，繳別繳次格式錯誤，請重新檢查。錯誤的值為：{1}", i, row.ItemArray[8].ToString()));
							}
						}

						//型態4的相關資料檢核
						if (row.ItemArray[0].ToString().Equals("4"))
						{
							#region 初年度服務報酬(FYC)
							//check = _service.GetGroupMappingTOZeroFour(row.ItemArray[4].ToString());
							if (firstyearList.Contains(row.ItemArray[4].ToString()))
							{
								//金額不等於FYC 可入系統
								if (Convert.ToInt32(row.ItemArray[9]) != Convert.ToInt32(row.ItemArray[10]))
								{
									warnList.Add(string.Format("檢測到資料列第{0}列，金額不等於FYC。", i));
								}
							}
							#endregion

							//保險公司代碼不得『空白』
							if (String.IsNullOrEmpty(row.ItemArray[1].ToString()))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，此原因碼保險公司代碼不得『空白』，請確認 保險公司代碼。", i));
							}

							#region 舊檢核
							////FYC不可空
							//if (row.ItemArray[7].ToString().TrimEnd().Equals(""))
							//{
							//    errorList.Add(string.Format("檢測到資料列第{0}列，FYC 不可為空，請重新檢查。", i));
							//}

							////FYP不可空
							//if (row.ItemArray[8].ToString().TrimEnd().Equals(""))
							//{
							//    errorList.Add(string.Format("檢測到資料列第{0}列，FYP 不可為空，請重新檢查。", i));
							//}

							////年期不可空
							//if (row.ItemArray[11].ToString().TrimEnd().Equals(""))
							//{
							//    errorList.Add(string.Format("檢測到資料列第{0}列，年期 不可為空，請重新檢查。", i));
							//}
							//else
							//{
							//    row.SetField(11, Convert.ToInt16(row.ItemArray[11].ToString()));
							//}

							//繳別繳次
							//Regex rgx = new Regex(@"^\d{4}[MSQDY]$");
							//if (!string.IsNullOrEmpty(row.ItemArray[12].ToString()))
							//{
							//    if (rgx.IsMatch(row.ItemArray[12].ToString()) == false)
							//    {
							//        errorList.Add(string.Format("檢測到資料列第{0}列，繳別繳次格式錯誤，請重新檢查。錯誤的值為：{1}", i, row.ItemArray[12].ToString()));
							//    }
							//}

							//计佣保费不可空
							//if (row.ItemArray[13].ToString().TrimEnd().Equals(""))
							//{
							//    errorList.Add(string.Format("檢測到資料列第{0}列，計傭保費 不可為空，請重新檢查。", i));
							//}
							#endregion
						}

						//型態6的相關資料檢核
						if (row.ItemArray[0].ToString().Equals("6"))
						{
							if (!row.ItemArray[5].ToString().TrimEnd().Equals(""))
							{
								var PoagAgCodeList = PoagList.Where(x => x.PolicyNo2.Equals(row.ItemArray[5].ToString().TrimEnd())).Select(x => x.AgentCode).ToList();
								// 從poag確認有無該業務員
								//check = _service.CheckAgentNameByPoag(row.ItemArray[5].ToString());
								if(PoagAgCodeList.Count == 0)
								{
									errorList.Add(string.Format("檢測到資料列第{0}列，查無此保單號碼。錯誤的值為：{1}", i, row.ItemArray[5].ToString()));
								}
								else if (PoagAgCodeList.Count != 0 && !PoagAgCodeList.Contains(row.ItemArray[2].ToString()))
								{
									errorList.Add(string.Format("檢測到資料列第{0}列，經手人錯誤。錯誤的值為：{1}", i, row.ItemArray[2].ToString()));
								}
							}
							//確認型態為6的原因碼
							//check = _service.GetGroupMappingTOsix(row.ItemArray[4].ToString());
							if (!sixList.Contains(row.ItemArray[4].ToString()))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，型態有誤，請確認。", i));
							}

							#region 初年度服務報酬(FYC)
							//check = _service.GetGroupMappingTOZeroFour(row.ItemArray[4].ToString());
							if (firstyearList.Contains(row.ItemArray[4].ToString()))
							{
								//金額不等於FYC 可入系統
								if (Convert.ToInt32(row.ItemArray[9]) != Convert.ToInt32(row.ItemArray[10]))
								{
									warnList.Add(string.Format("檢測到資料列第{0}列，金額不等於FYC。", i));
								}
							}
							#endregion

							//保險公司代碼不得『空白』
							if (String.IsNullOrEmpty(row.ItemArray[1].ToString()))
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，此原因碼保險公司代碼不得『空白』，請確認 保險公司代碼。", i));
							}
						}

						//型態[0] 保公代碼[1] 業務員代碼[2] 姓名[3] 調整原因碼[4] 保單號碼[5] 險種代碼[6] 年期[7] 繳別繳次[8] 金額[9] FYC[10] FYP[11]  說明[12] 來源資料表鍵值[13]
						//型態8的相關資料設定
						if (row.ItemArray[0].ToString().Equals("8"))
						{
							row.SetField(10, 0);  //型態8不需要fyc，預設為0
							row.SetField(11, 0);  //型態8不需要fyp，預設為0
							//row.SetField(13, ""); //來源資料表Iden預設空值
							//row.SetField(26, "");  //來源資料表預設空值
						}

						//型態4的相關資料設定
						//if (row.ItemArray[0].ToString().Equals("4"))
						//{
						//	row.SetField(13, ""); //來源資料表Iden預設空值
						//	row.SetField(26, "");  //來源資料表預設空值
						//}

						//型態6的相關資料設定
						//if (row.ItemArray[0].ToString().Equals("6"))
						//{
						//if (!row.ItemArray[13].ToString().Equals(""))
						//{
						//	row.SetField(26, "MerSalCut"); //來源資料表
						//}
						//}

						//20241104 調整來源資料表不分型態，都要寫入
						if (!row.ItemArray[13].ToString().Equals(""))
						{
							row.SetField(26, "MerSalCut"); //來源資料表
						}

						//離職人員不調整
						if (termination)
						{
							//aginb -> ag_status_code in (2,3) 顯示給user
							if (quits.Where(x => x.Equals(row.ItemArray[2].ToString().TrimEnd())).Any())
							{
								errorList.Add(string.Format("檢測到資料列第{0}列，業務員 {1}，{2} 工作月已離職，不可調整，請修改後重新上傳",
									i, row.ItemArray[2].ToString(), productionYM));
							}
						}
						i++;
					}
					#endregion

					if (errorList.Count == 0)
					{
						#region 寫入系統需求欄位
						string processNo = PayRollHelper.GetNewProcessNo();

						foreach (DataRow row in csvDT.Rows)
						{
							//row.SetField(0, sequence);
							row.ItemArray[4] = row.ItemArray[4].ToString().ToUpper();
							row.SetField(1, row.ItemArray[1] == DBNull.Value ? "" : row.ItemArray[1]); //保公
							row.SetField(5, row.ItemArray[5] == DBNull.Value ? "" : row.ItemArray[5]); //保單
							row.SetField(6, row.ItemArray[6] == DBNull.Value ? "" : row.ItemArray[6]); //險種代碼
							row.SetField(7, row.ItemArray[7] == DBNull.Value ? "0" : row.ItemArray[7]); //年期
							row.SetField(8, row.ItemArray[8] == DBNull.Value ? "" : row.ItemArray[8]); //繳別繳次
							row.SetField(13, row.ItemArray[13] == DBNull.Value ? "" : row.ItemArray[13]); //來源資料表鍵值[13]
							row.SetField(14, Guid.NewGuid());
							row.SetField(15, productionYM);
							row.SetField(16, processNo);
							row.SetField(17, _programID);
							row.SetField(18, Guid.NewGuid());
							row.SetField(19, DateTime.Now);
							row.SetField(20, User.AccountInfo.ID);
							row.SetField(21, DateTime.Now);
							row.SetField(22, User.AccountInfo.ID);
							row.SetField(23, sequence);
							row.SetField(24, PayRollHelper.ChangeUnitID(User.MemberInfo.ID));
							row.SetField(25, PayRollHelper.ChangeUnitName(User.MemberInfo.ID));
							row.SetField(27, 0);  //计佣保费，預設為0

							//row.SetField(25, (companySets.Where(x => x.CompanyCode.Equals(row.ItemArray[2].ToString().TrimEnd())).Select(x => x.CombineCode)).FirstOrDefault() ?? "");
							#region 產生AgentBonusDesc資料
							AgentBonusDesc bd = new AgentBonusDesc();
							string DescContentMask = string.Empty;
							if (row.ItemArray[12].ToString() != "")
							{								
								for (int jj = 0; jj < row.ItemArray[12].ToString().Length; jj++)
								{
									DescContentMask += "*";
								}								
							}
							bd.DescContentMask = DescContentMask;
							bd.Guid = Guid.Parse(row.ItemArray[18].ToString());
							bd.ProductionYM = productionYM;
							bd.Sequence = Convert.ToInt16(sequence);
							bd.DescContent = !String.IsNullOrEmpty(row.ItemArray[12].ToString()) ? row.ItemArray[12].ToString() : string.Empty;
							bd.DescValue = "";
							bd.ProcessCode = _programID;
							bd.RelationTable = "AgentBonusAdjust";
							bd.CreateDatetime = DateTime.Now;
							bd.CreateUserCode = User.AccountInfo.ID;
							descs.Add(bd);
							#endregion
						}
						#endregion

						csvDT.Columns.RemoveAt(12); //說明

						//_service.DeleteAdjustByAccountID(productionYM, sequence, User.AccountInfo.ID);  //刪除大批資料，比對條件是 uploader ID & productionYM，不分單筆多筆調整
						_service.ag078BulkInsert(csvDT, descs);  //todo bulkcopy

						if (warnList.Count == 0)
						{
							AppendMessage("檔案上傳成功。", false);
							return Json(csvDT.Rows.Count);
						}
						else
						{
							Tuple<List<string>, int> tuple = new Tuple<List<string>, int>(warnList, csvDT.Rows.Count);
							AppendMessage("檔案上傳成功。", false);
							//Response.Write(("<script>confirm('是否查詢錯誤訊息?')</script>"));
							return Json(tuple);
						}
					}
					else
					{
						AppendMessage("檔案上傳失敗。", false);
						return Json(errorList);
					}
				}
				else
				{
					errorList.Add("檔案大小異常，請重新確認");

					AppendMessage("檔案上傳失敗。", false);
					return Json(errorList);
				}
			}
			catch (Exception ex)
			{
				if (ex.ToString().Contains("找不到資料行"))
				{
					AppendMessage("檔案上傳失敗，欄位格式有誤。", false);
				}
				else
				{
					AppendMessage("檔案上傳失敗，資料異常。", false);
				}
				return Json(errorList);
			}
		}

		/// <summary>
		/// 刪除
		/// </summary>
		[HttpPost]
		[HasPermission("EB.SL.PayRoll.PRTX003")]
		public void Delete(string productionYM, string sequence)
		{
			_service.DeleteAdjustByAccountID(productionYM, sequence, User.AccountInfo.ID, _programID);
		}

		/// <summary>
		///  上傳更新原因碼一覽表
		/// </summary>
		[HttpPost]
		[HasPermission("EB.SL.PayRoll.PRTX003")]
		public void UploadReasonPDF(HttpPostedFileBase reasonUpload)
		{
			try
			{
				//檢查檔案大小
				var fileSize = reasonUpload.ContentLength;
				if (fileSize > 0 &&
					(fileSize / 1048576.0) < Int32.Parse(EB.Web.PlatformHelper.GetVarConfig("UploadLimit").ToString()))
				{
					string dirPath = Server.MapPath("~/Areas/PayRoll/Template/");
					string fullPath = Path.Combine(dirPath, "適用原因碼.pdf");
					if (!Directory.Exists(dirPath))
					{
						Directory.CreateDirectory(dirPath);
					}
					reasonUpload.SaveAs(fullPath);

					//紀錄log
					SampleFileUploadLog uploadLog = new SampleFileUploadLog();
					uploadLog.ProgramName = _programID;
					uploadLog.FileName = reasonUpload.FileName;
					uploadLog.UploadUserCode = User.AccountInfo.ID;
					uploadLog.UploadDatetime = DateTime.Now;

					_service.CreateSampleFileUploadLog(uploadLog);

					AppendMessage("檔案上傳成功");
				}
				else
				{
					AppendMessage("檔案大小異常，請檢查檔案後重新上傳");
				}
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message);
			}
		}
	}
}