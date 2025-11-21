///========================================================================================== 
/// 程式名稱：人工調帳
/// 建立人員：Harrison
/// 建立日期：2022/07
/// 修改記錄：（[需求單號]、[修改內容]、日期、人員）
/// 需求單號:20240122004-因現有VLIFE系統(核心系統)使用已長達20多年，架構老舊，已不敷使用，且為提升資訊安全等級，故計劃執行VLIFE系統改版(新核心系統:eBroker系統)。; 修改內容:上線; 修改日期:20240613; 修改人員:Harrison;
/// 需求單號:20240807001-調整人工調帳系統產出之相關畫面及報表修改等功能。; 修改日期:20240807; 修改人員:Harrison;
///==========================================================================================
using EB.EBrokerModels;
using EB.SL.PayRoll.Models;
using EB.SL.PayRoll.Service;
using EB.SL.PayRoll.Service.Contracts;
using EB.SL.PayRoll.Web.Areas.PayRoll.Utilities;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EB.SL.PayRoll.Web.Areas.PayRoll.Controllers
{
	[Program("PRTX002")]
	public class PRTX002Controller : BaseController
	{
		private IPayRollService _service;
		private static string _programID = "PRTX002";
		public PRTX002Controller()
		{
			_service = ServiceHelper.Create<IPayRollService>();
		}
		// GET: PayRoll/PRTX002
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public ActionResult Index()
		{
			return View();
		}

		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public ActionResult Create()
		{
			return View();
		}


		/// <summary>
		/// update
		/// </summary>
		/// <param name="iden"></param>
		/// <returns></returns>
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public ActionResult Update(int tid)
		{
			Tuple<AgentBonusAdjust, AgentBonusDesc> data = _service.QueryAgentBonusByIden(tid);
			var result = new AgentBonusAdjustViewModel();
			result.AdjType = data.Item1.AdjType;
			result.AgentBonusDescGuid = data.Item1.AgentBonusDescGuid;
			result.AgentCode = data.Item1.AgentCode;
			result.AgentName = data.Item1.AgentName;
			result.Amount = data.Item1.Amount;
			result.CollectYear = data.Item1.CollectYear;
			//result.CommModePrem = data.Item1.CommModePrem;
			result.CompanyCode = data.Item1.CompanyCode ?? "";
			result.CreateDatetime = data.Item1.CreateDatetime;
			result.CreateUserCode = data.Item1.CreateUserCode;
			result.FYC = data.Item1.FYC;
			result.FYP = data.Item1.FYP;
			result.Guid = data.Item1.Guid;
			result.ModxSequence = data.Item1.ModxSequence ?? "";
			result.PlanCode = data.Item1.PlanCode ?? "";
			result.PolicyNo2 = data.Item1.PolicyNo2 ?? "";
			result.ProcessCode = data.Item1.ProcessCode;
			result.ProcessNo = data.Item1.ProcessNo;
			result.ProductionYM = data.Item1.ProductionYM;
			result.ReasonCode = data.Item1.ReasonCode;
			result.Sequence = data.Item1.Sequence;
			result.DescContent = data.Item2.DescContent ?? "";

			return View(result);
		}

		/// <summary>
		/// 寫入單筆調整
		/// </summary>
		/// <param name="model">單筆調整view model</param>
		/// <returns></returns>
		[HttpPost]
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public JsonResult Index(AgentBonusAdjustViewModel model)
		{
			//this.DefaultView.ViewName = "Create";
			this.TempData["model"] = model;
			string check = string.Empty;
			List<string> ErrorList = new List<string>();
			#region 檢核
			if (!User.HasPermission("EB.SL.PayRoll.PRTX002.Admin"))
			{
				if(model.AdjType == "4" || model.AdjType == "6")
				{
					Throw.BusinessError("錯誤訊息:無權限執行型態4及型態6");
				}
			}

			check = CheckOrgib(model.AgentCode, model.FYC.ToString(), model.ProductionYM, model.Sequence.ToString());
			if (check == "ERROR")
			{
				Throw.BusinessError("錯誤訊息:業務員任職型態3(內勤)人員，FYC需為0");
			}

			check = CheckRegisterDate(model.AgentCode, model.ProductionYM);
			if (check == "ERROR")
			{
				Throw.BusinessError("錯誤訊息:簽約日期不可大於業績年約");
			}

			check = _service.GetGroupMappingTOmodxseq(model.ReasonCode);
			if (!String.IsNullOrEmpty(check))
			{
				if (String.IsNullOrEmpty(model.ModxSequence))
				{
					Throw.BusinessError("錯誤訊息:此原因碼，繳别繳次必須要有內容");
				}
			}
			else
			{
				model.ModxSequence = string.Empty;
			}

			check = _service.GetGroupMappingTOnotin(model.ReasonCode);
			if (!String.IsNullOrEmpty(check))
			{
				Throw.BusinessError("錯誤訊息:此原因碼不可人工調整");
			}

			check = _service.GetGroupMappingTOplus(model.ReasonCode);
			if (!String.IsNullOrEmpty(check))
			{
				if (model.Amount < 0)
				{
					Throw.BusinessError("錯誤訊息:此為正項項目，金額應為正數");
				}
			}

			check = _service.GetGroupMappingTOminus(model.ReasonCode);
			if (!String.IsNullOrEmpty(check))
			{
				if (model.Amount > 0)
				{
					Throw.BusinessError("錯誤訊息:此為負項項目，金額應為負數");
				}
			}

			check = CheckBonusReserveSetting(model.ProductionYM, model.Sequence.ToString(), model.AgentCode, model.ReasonCode);
			if (check == "ERROR")
			{
				ErrorList.Add("有黑名單人員，請確認");
			}

			check = CheckOrgibTrue(model.AgentCode, model.ProductionYM, model.Sequence.ToString());
			if (check == "ERROR" && model.ReasonCode == "AT")
			{
				ErrorList.Add("有內勤人員，請確認");
			}

			check = CheckAginbFDStatus(model.ProductionYM, model.Sequence.ToString(), model.AgentCode);
			if (check == "ERROR" && model.ReasonCode == "0A")
			{
				ErrorList.Add("發放人員為免評估，請確認");
			}

			#region 初年度服務報酬
			var firstyearList = _service.GetGroupMappingTOfirstyearList();
			if (firstyearList.Contains(model.ReasonCode))
			{
				//金額不得大於保費
				if (PayRollHelper.ABS(model.Amount) > PayRollHelper.ABS(model.FYP))
				{
					Throw.BusinessError("錯誤訊息:金額不得大於保費，請重新檢查");
				}

				//金額不等於FYC 可入系統
				if (model.Amount != model.FYC)
				{
					ErrorList.Add("金額不等於FYC。");
				}

				//金額不得大於計佣保費 可入系統
				//if (PayRollHelper.ABS(model.Amount) > PayRollHelper.ABS(model.CommModePrem))
				//{
				//    ErrorList.Add("金額不得大於計佣保費。");
				//}

				////計佣保費不得大於保費 可入系統
				//if (PayRollHelper.ABS(model.CommModePrem) > PayRollHelper.ABS(model.FYP))
				//{
				//    ErrorList.Add("計佣保費不得大於保費。");
				//}
			}
			#endregion

			var AFList = _service.GetGroupMappingTOAF();
			if (AFList.Contains(model.ReasonCode))
			{
				//金額不得大於保費
				if (PayRollHelper.ABS(model.Amount) > PayRollHelper.ABS(model.FYP))
				{
					Throw.BusinessError("錯誤訊息:金額不得大於保費，請重新檢查");
				}
			}
			#endregion

			AgentBonusAdjust adjustModel = new AgentBonusAdjust();
			adjustModel.AdjType = model.AdjType;
			adjustModel.AgentBonusDescGuid = Guid.NewGuid();
			adjustModel.AgentCode = model.AgentCode;
			adjustModel.AgentName = model.AgentName;
			adjustModel.Amount = model.Amount;
			adjustModel.CollectYear = model.AdjType == "8" ? (short)0 : model.CollectYear;
			adjustModel.CommModePrem = 0;
			adjustModel.CompanyCode = model.AdjType == "8" ? "" : model.CompanyCode ?? "";
			//adjustModel.CompanyCodeSub = model.CompanyCodeSub;
			adjustModel.CreateDatetime = DateTime.Now;
			adjustModel.CreateUserCode = User.AccountInfo.ID;
			adjustModel.CreateUnit = PayRollHelper.ChangeUnitID(User.MemberInfo.ID);
			adjustModel.CreateUnitName = PayRollHelper.ChangeUnitName(User.MemberInfo.ID);
			adjustModel.UpdateDateTime = DateTime.Now;
			adjustModel.UpdateUserCode = User.AccountInfo.ID;
			adjustModel.FYC = model.AdjType == "8" ? 0 : model.FYC;
			adjustModel.FYP = model.AdjType == "8" ? 0 : model.FYP;
			adjustModel.Guid = Guid.NewGuid();
			adjustModel.WCCode = model.WCCode ?? "";
			adjustModel.ModxSequence = model.AdjType == "8" ? "" : model.ModxSequence ?? "";
			adjustModel.PlanCode = model.AdjType == "8" ? "" : model.PlanCode ?? "";
			adjustModel.PolicyNo2 = model.AdjType == "8" ? "" : model.PolicyNo2 ?? "";
			adjustModel.ProcessCode = _programID;
			adjustModel.ProcessNo = PayRollHelper.GetNewProcessNo();
			adjustModel.ProductionYM = model.ProductionYM;
			adjustModel.ReasonCode = model.ReasonCode.ToUpper();
			adjustModel.Sequence = model.Sequence;
			adjustModel.SourceTable = "";
			adjustModel.SourceTableKey = "";

			AgentBonusDesc descModel = new AgentBonusDesc();
			descModel.DescContent = model.DescContent ?? "";
			string DescContentMask = string.Empty;
			if (!String.IsNullOrEmpty(model.DescContent))
			{				
				for (int i = 0; i < model.DescContent.Length; i++)
				{
					DescContentMask += "*";
				}				
			}
			descModel.DescContentMask = DescContentMask;
			descModel.DescValue = "";
			descModel.Guid = adjustModel.AgentBonusDescGuid;
			descModel.ProductionYM = adjustModel.ProductionYM;
			descModel.Sequence = adjustModel.Sequence;
			descModel.ProcessCode = _programID;
			descModel.RelationTable = "AgentBonusAdjust";
			descModel.CreateDatetime = adjustModel.CreateDatetime;
			descModel.CreateUserCode = adjustModel.CreateUserCode;
			List<AgentBonusAdjust> agentBonus = new List<AgentBonusAdjust>();
			agentBonus = _service.InsertSingleAgentBonusAdjust(adjustModel, descModel);

			Tuple<List<AgentBonusAdjust>, List<string>> result = new Tuple<List<AgentBonusAdjust>, List<string>>(agentBonus, ErrorList);

			return Json(result, JsonRequestBehavior.AllowGet);
			//bool result =
			//if (result)
			//{
			//    AppendMessage(新增成功);
			//}
			//else
			//{
			//    AppendMessage(新增失敗);
			//}       
			//return View();
			//return RedirectToAction("Create", "PRTX003");
		}

		/// <summary>
		/// 查詢調整紀錄
		/// </summary>
		/// <param name="model">AgentBonusAdjustViewModel</param>
		[HttpPost]
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public void Query(AgentBonusAdjustViewModel model)
		{
			QueryAgentBonusCondition condition = new QueryAgentBonusCondition();
			condition.AdjType = model.AdjType;
			condition.AgentCode = model.AgentCode;
			condition.CompanyCode = model.CompanyCode;
			condition.ProductionYM = model.ProductionYM;
			condition.ReasonCode = model.ReasonCode;
			condition.Sequence = model.Sequence;
			condition.DescContent = model.DescContent;
			condition.AgentName = model.AgentName;
			condition.PolicyNo2 = model.PolicyNo2;
			condition.Amount = model.Amount;
			condition.FYC = model.FYC;
			condition.FYP = model.FYP;
			condition.PlanCode = model.PlanCode;
			condition.CollectYear = model.CollectYear;
			condition.ModxSequence = model.ModxSequence;
			//condition.CommModePrem = model.CommModePrem;
			condition.CreateUserCode = User.AccountInfo.ID;//只看的到自己所新增和大批匯入的

			var channel = new WebChannel<IPayRollService>();
			var gridList = Enumerable.Empty<AgentBonusAdjustViewModel>();
			channel.Use(service =>
			{
				var dataList = service.QueryAgentBonus(condition);
				gridList = dataList.Select(m =>
				{
					var result = new AgentBonusAdjustViewModel();
					result.Iden = m.Item1.iden;
					result.AdjType = m.Item1.AdjType;
					result.AgentBonusDescGuid = m.Item1.AgentBonusDescGuid;
					result.AgentCode = m.Item1.AgentCode;
					result.AgentName = m.Item1.AgentName;
					result.AmountTS = m.Item1.Amount != 0 ? m.Item1.Amount.ToString("###,###") : "0";
					result.CollectYear = m.Item1.CollectYear;
					//result.CommModePrem = m.Item1.CommModePrem;
					result.CompanyCode = m.Item1.CompanyCode;
					result.CreateDatetime = m.Item1.CreateDatetime;
					result.CreateUserCode = m.Item1.CreateUserCode;
					result.FYCTS = m.Item1.FYC != 0 ? m.Item1.FYC.ToString("###,###") : "0";
					result.FYPTS = m.Item1.FYP != 0 ? m.Item1.FYP.ToString("###,###") : "0";
					result.Guid = m.Item1.Guid;
					result.ModxSequence = m.Item1.ModxSequence;
					result.PlanCode = m.Item1.PlanCode;
					result.PolicyNo2 = m.Item1.PolicyNo2;
					result.ProcessCode = m.Item1.ProcessCode;
					result.ProcessNo = m.Item1.ProcessNo;
					result.ProductionYM = m.Item1.ProductionYM;
					result.ReasonCode = m.Item1.ReasonCode;
					result.Sequence = m.Item1.Sequence;
					result.DescContent = m.Item2.DescContent;
					return result;
				}).ToList();
			});

			var gridKey = channel.DataToCache(gridList);
			SetGridKey("AgentBonusAdjustQueryGrid", gridKey);
		}
		public JsonResult BindLogGrid(jqGridParam jqParams)
		{
			var cacheKey = GetGridKey("AgentBonusAdjustQueryGrid");
			return BaseGridBinding<AgentBonusAdjustViewModel>(jqParams,
				() => new WebChannel<IPayRollService, AgentBonusAdjustViewModel>().Get(cacheKey));
		}

		/// <summary>
		/// 刪除一筆調整資料
		/// </summary>
		/// <param name="agentBonusDescGuid">業務酬佣說明檔GUID</param>
		[HttpPost]
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public void Delete(string agentBonusDescGuid)
		{
			bool result = _service.DeleteAgentBonusAdjustByAgentBonusDescGuid(agentBonusDescGuid, User.AccountInfo.ID);

			if (result)
				AppendMessage("刪除成功", false);
			else
				AppendMessage("刪除失敗", false);
		}

		/// <summary>
		/// 更新紀錄
		/// </summary>
		/// <param name="model"></param>
		[HttpPost]
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public JsonResult Update(AgentBonusAdjustViewModel model)
		{
			this.TempData["model"] = model;
			string check = string.Empty;
			List<string> ErrorList = new List<string>();
			List<AgentBonusAdjust> agentBonus = new List<AgentBonusAdjust>();
			Tuple<List<AgentBonusAdjust>, List<string>> resultModel = new Tuple<List<AgentBonusAdjust>, List<string>>(agentBonus, ErrorList);
			try
			{
				AgentBonusAdjust adjustModel = new AgentBonusAdjust();
				adjustModel.AdjType = model.AdjType;
				adjustModel.AgentBonusDescGuid = model.AgentBonusDescGuid;
				adjustModel.AgentCode = model.AgentCode;
				adjustModel.AgentName = model.AgentName;
				adjustModel.Amount = model.Amount;
				adjustModel.CollectYear = model.AdjType == "8" ? (short)0 : model.CollectYear;
				adjustModel.CommModePrem = 0;
				adjustModel.CompanyCode = model.AdjType == "8" ? "" : model.CompanyCode ?? "";
				adjustModel.FYC = model.AdjType == "8" ? 0 : model.FYC;
				adjustModel.FYP = model.AdjType == "8" ? 0 : model.FYP;
				adjustModel.Guid = model.Guid;
				adjustModel.ModxSequence = model.AdjType == "8" ? "" : model.ModxSequence ?? "";
				adjustModel.PlanCode = model.AdjType == "8" ? "" : model.PlanCode ?? "";
				adjustModel.PolicyNo2 = model.AdjType == "8" ? "" : model.PolicyNo2 ?? "";
				adjustModel.WCCode = model.WCCode;
				adjustModel.ProcessCode = _programID;
				adjustModel.ProductionYM = model.ProductionYM;
				adjustModel.ReasonCode = model.ReasonCode.ToUpper();
				adjustModel.Sequence = model.Sequence;
				adjustModel.UpdateDateTime = DateTime.Now;
				adjustModel.UpdateUserCode = User.AccountInfo.ID;

				AgentBonusDesc descModel = new AgentBonusDesc();
				descModel.DescContent = model.DescContent ?? "";
				string DescContentMask = string.Empty;
				if (!String.IsNullOrEmpty(model.DescContent))
				{
					for (int i = 0; i < model.DescContent.Length; i++)
					{
						DescContentMask += "*";
					}
				}
				descModel.DescContentMask = DescContentMask;
				descModel.Guid = adjustModel.AgentBonusDescGuid;
				descModel.ProductionYM = adjustModel.ProductionYM;
				descModel.Sequence = adjustModel.Sequence;
				descModel.ProcessCode = _programID;
				descModel.CreateUserCode = User.AccountInfo.ID;
				agentBonus.Add(adjustModel);
				resultModel = new Tuple<List<AgentBonusAdjust>, List<string>>(agentBonus, ErrorList);
				
				#region 檢核
				if (!User.HasPermission("EB.SL.PayRoll.PRTX002.Admin"))
				{
					if (model.AdjType == "4" || model.AdjType == "6")
					{
						Throw.BusinessError("錯誤訊息:無權限執行型態4及型態6");
					}
				}
				check = CheckOrgib(model.AgentCode, model.FYC.ToString(), model.ProductionYM, model.Sequence.ToString());
				if (check == "ERROR")
				{
					Throw.BusinessError("錯誤訊息:業務員任職型態3(內勤)人員，FYC需為0");
				}

				check = CheckRegisterDate(model.AgentCode, model.ProductionYM);
				if (check == "ERROR")
				{
					Throw.BusinessError("錯誤訊息:簽約日期不可大於業績年約");
				}

				check = _service.GetGroupMappingTOmodxseq(model.ReasonCode);
				if (!String.IsNullOrEmpty(check))
				{
					if (String.IsNullOrEmpty(model.ModxSequence))
					{
						Throw.BusinessError("錯誤訊息:此原因碼，繳别繳次必須要有內容");
					}
				}
				else
				{
					model.ModxSequence = string.Empty;
				}

				check = _service.GetGroupMappingTOnotin(model.ReasonCode);
				if (!String.IsNullOrEmpty(check))
				{
					Throw.BusinessError("錯誤訊息:此原因碼不可人工調整");
				}

				check = _service.GetGroupMappingTOplus(model.ReasonCode);
				if (!String.IsNullOrEmpty(check))
				{
					if (model.Amount < 0)
					{
						Throw.BusinessError("錯誤訊息:此為正項項目，金額應為正數");
					}
				}

				check = _service.GetGroupMappingTOminus(model.ReasonCode);
				if (!String.IsNullOrEmpty(check))
				{
					if (model.Amount > 0)
					{
						Throw.BusinessError("錯誤訊息:此為負項項目，金額應為負數");
					}
				}

				check = CheckBonusReserveSetting(model.ProductionYM, model.Sequence.ToString(), model.AgentCode, model.ReasonCode);
				if (check == "ERROR")
				{
					ErrorList.Add("有黑名單人員，請確認");
				}

				check = CheckOrgibTrue(model.AgentCode, model.ProductionYM, model.Sequence.ToString());
				if (check == "ERROR" && model.ReasonCode == "AT")
				{
					ErrorList.Add("有內勤人員，請確認");
				}

				check = CheckAginbFDStatus(model.ProductionYM, model.Sequence.ToString(), model.AgentCode);
				if (check == "ERROR" && model.ReasonCode == "0A")
				{
					ErrorList.Add("發放人員為免評估，請確認");
				}

				#region 初年度服務報酬
				var firstyearList = _service.GetGroupMappingTOfirstyearList();
				if (firstyearList.Contains(model.ReasonCode))
				{
					//金額不得大於保費
					if (PayRollHelper.ABS(model.Amount) > PayRollHelper.ABS(model.FYP))
					{
						Throw.BusinessError("錯誤訊息:金額不得大於保費，請重新檢查");
					}

					//金額不等於FYC 可入系統
					if (model.Amount != model.FYC)
					{
						ErrorList.Add("金額不等於FYC。");
					}
				}
				#endregion

				var AFList = _service.GetGroupMappingTOAF();
				if (AFList.Contains(model.ReasonCode))
				{
					//金額不得大於保費
					if (PayRollHelper.ABS(model.Amount) > PayRollHelper.ABS(model.FYP))
					{
						Throw.BusinessError("錯誤訊息:金額不得大於保費，請重新檢查");
					}
				}
				#endregion

				bool result = _service.UpdateAgentBonusAdjust(adjustModel, descModel);
	
				if (result)
					AppendMessage("更新成功");
				else
					AppendMessage("更新失敗");

				return Json(resultModel, JsonRequestBehavior.AllowGet);
			}
			catch(Exception ex)
			{
				var r = new Tuple<List<AgentBonusAdjust>, string>(agentBonus, "ERROR");
				AppendMessage(ex.Message);
				return Json(r, JsonRequestBehavior.AllowGet);
			}
		}

		/// <summary>
		/// 該業務員代碼是否存在於該業績年月的Aginb
		/// </summary>
		/// <param name="productionYM">業績年月</param>
		/// <param name="agentCode">業務員代碼</param>
		/// <returns></returns>
		[HttpPost]
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public string CheckAginb(string productionYM, string agentCode)
		{
			productionYM = PayRollHelper.WYToRoc(productionYM);
			return PayRollHelper.CheckAginb(productionYM, agentCode);
		}

		/// <summary>
		/// 輸入保單號碼，取得業務員姓名
		/// </summary>
		/// <param name="PolicyNo2"></param>
		/// <returns></returns>
		[HttpPost]
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public ActionResult GetAgentNameByPoag(string PolicyNo2)
		{
			var result = _service.GetAgentNameByPoag(PolicyNo2);
			if (result.Count == 0)
			{
				return Json("Error");
			}
			else
			{
				return Json(result);
			}
		}

		/// <summary>
		/// 輸入原因碼取得對應的型態代碼
		/// </summary>
		/// <param name="reasonCode">原因碼</param>
		/// <returns></returns>
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public string GetTypeByReasonCode(string reasonCode)
		{
			return PayRollHelper.GetTypeByReasonCode(reasonCode.ToUpper());
		}

		/// <summary>
		/// 確認保險公司代碼
		/// </summary>
		/// <param name="CompanyCode">保險公司代碼</param>
		/// <returns></returns>
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public string CheckCompanyCode(string CompanyCode)
		{
			if (!PayRollHelper.CheckCompanyCodeIsExist(CompanyCode))
			{
				return null;
			}
			else
			{
				return CompanyCode;
			}
		}

		/// <summary>
		/// 確認任職型態3(內勤)人員，FYC需為0
		/// </summary>
		/// <param name="agentCode"></param>
		/// <param name="FYC"></param>
		/// <returns></returns>
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public string CheckOrgib(string agentCode, string FYC, string ProductionYM, string Sequence)
		{
			ProductionYM = PayRollHelper.WYToRoc(ProductionYM);
			string Agoccpind = _service.GetOrgibByagentcode(agentCode, ProductionYM, Sequence);

			if (Agoccpind == "3" && FYC != "0")
			{
				return "ERROR";
			}
			else
			{
				return "OK";
			}
		}

		/// <summary>
		/// 確認是否為(內勤)人員
		/// </summary>
		/// <param name="agentCode"></param>
		/// <param name="FYC"></param>
		/// <returns></returns>
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public string CheckOrgibTrue(string agentCode, string ProductionYM, string Sequence)
		{
			ProductionYM = PayRollHelper.WYToRoc(ProductionYM);
			string Agoccpind = _service.GetOrgibByagentcode(agentCode, ProductionYM, Sequence);

			if (Agoccpind == "3")
			{
				return "ERROR";
			}
			else
			{
				return "OK";
			}
		}

		/// <summary>
		/// 確認簽約日期是否大於業績年月
		/// </summary>
		/// <param name="agentCode"></param>
		/// <param name="ProductionYM"></param>
		/// <returns></returns>
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public string CheckRegisterDate(string agentCode, string ProductionYM)
		{
			string RegisterDate = _service.GetAginbbByAgentcode(agentCode);
			int result = DateTime.Compare(Convert.ToDateTime(RegisterDate), Convert.ToDateTime(ProductionYM));
			if (result == 1)
			{
				return "ERROR";
			}
			else
			{
				return "OK";
			}
		}

		/// <summary>
		/// 確認業務員是否為黑名單
		/// </summary>
		/// <param name="ProductionYM"></param>
		/// <param name="Sequence"></param>
		/// <param name="agentcode"></param>
		/// <param name="ReasonCode"></param>
		/// <returns></returns>
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public string CheckBonusReserveSetting(string ProductionYM, string Sequence, string agentcode, string ReasonCode)
		{
			string agent = _service.GetBonusReserveSetting(ProductionYM, Sequence, agentcode, ReasonCode);
			if (!String.IsNullOrEmpty(agent))
			{
				return "ERROR";
			}
			else
			{
				return "OK";
			}
		}

		/// <summary>
		/// 檢核對應的原因碼，此原因碼為初年度服務報酬
		/// </summary>
		/// <param name="ReasonCode"></param>
		/// <returns></returns>
		public string GetGroupMappingTOZeroFour(string ReasonCode)
		{
			string agent = _service.GetGroupMappingTOZeroFour(ReasonCode);
			if (String.IsNullOrEmpty(agent))
			{
				return "ERROR";
			}
			else
			{
				return "OK";
			}
		}

		/// <summary>
		/// 確認發放人員為免評估人員
		/// </summary>
		/// <param name="ProductionYM"></param>
		/// <param name="Sequence"></param>
		/// <param name="agentcode"></param>
		/// <param name="ReasonCode"></param>
		/// <returns></returns>
		[HasPermission("EB.SL.PayRoll.PRTX002")]
		public string CheckAginbFDStatus(string ProductionYM, string Sequence, string agentcode)
		{
			ProductionYM = PayRollHelper.WYToRoc(ProductionYM);
			string agent = _service.GetaginbByagentcode(agentcode, ProductionYM, Sequence);
			if (agent == "1")
			{
				return "ERROR";
			}
			else
			{
				return "OK";
			}
		}
	}
}