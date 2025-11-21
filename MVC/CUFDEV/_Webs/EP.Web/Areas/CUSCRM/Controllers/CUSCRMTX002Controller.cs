using EP.Common;
using EP.CUFModels;
using EP.H2OModels;
using EP.Platform.Service;
using EP.PSL.IB.Service;
using EP.SD.SalesSupport.CUSCRM.Service;
using EP.SD.SalesSupport.CUSCRM.Service.Contracts;
using EP.Web;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Service;
using Microsoft.CUF.Web;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net.Mail;

namespace EP.SD.SalesSupport.CUSCRM.Web.Areas.CUSCRM.Controllers
{
	public class CUSCRMTX002Controller : BaseController
	{
		// GET: CUSCRM/CUSCRMTX002
		private ICaseService caseService;
		private INotifyService _Service;
		private IMemberExtendService _MemberEService;
		private ICommonService _CommonService;
		public CUSCRMTX002Controller()
		{
			_Service = ServiceHelper.Create<INotifyService>();
			_MemberEService = ServiceHelper.Create<IMemberExtendService>();
			_CommonService = ServiceHelper.Create<ICommonService>();
			caseService = ServiceHelper.Create<ICaseService>();
		}
		#region 查詢
		[HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX002")]
		public ActionResult Index()
		{
			return View();
		}

		/// <summary>
		/// 通知作業-查詢
		/// </summary>
		/// <param name="LawNote"></param>
		[HttpPost]
		[HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX002")]
		[PdLogFilter("EIP客服-通知作業-查詢", PITraceType.Query)]
		public JsonResult Query(QueryNotifyCondition model)
		{
			//取資料
			List<NotifyViewModel> Viewmodel = new List<NotifyViewModel>();
			int j = 1;
			if (User.HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX002.Admin"))//客服
			{
				Viewmodel = _Service.QueryNotifyDataByCRM(model);
				for (int i = 0; i < Viewmodel.Count; i++)
				{
					string mStr = Viewmodel[i].AllToMemberID+"," + Viewmodel[i].AllCCMemberID;
					string mStr1 = "";
					if (!String.IsNullOrEmpty(mStr))
					{
						mStr1 = string.Join(",", mStr.Split(',').Distinct().ToArray());
					}					
					Viewmodel[i].AllToMemberID = mStr1;
					Viewmodel[i].ViewID = j;
					Viewmodel[i].CreateTime = Convert.ToDateTime(Viewmodel[i].CreateTime).ToString("yyyy/MM/dd");
					Viewmodel[i].AuditCreateTime = String.IsNullOrEmpty(Viewmodel[i].AuditCreateTime) ? "N" : Convert.ToDateTime(Viewmodel[i].AuditCreateTime).ToString("yyyy/MM/dd");
					Viewmodel[i].DoSCreateTime = String.IsNullOrEmpty(Viewmodel[i].DoSCreateTime) ? "N" : Convert.ToDateTime(Viewmodel[i].DoSCreateTime).ToString("yyyy/MM/dd");
					j++;
				}
			}
			else if (User.HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX002.Employee"))//行專
			{
				model.MemberID = User.MemberInfo.ID;
				Viewmodel = _Service.QueryNotifyEPDataByCRM(model);
				for (int i = 0; i < Viewmodel.Count; i++)
				{
					Viewmodel[i].ViewID = j;
					Viewmodel[i].CreateTime = Convert.ToDateTime(Viewmodel[i].CreateTime).ToString("yyyy/MM/dd");
					Viewmodel[i].AuditCreateTime = String.IsNullOrEmpty(Viewmodel[i].AuditCreateTime) ? "N" : Convert.ToDateTime(Viewmodel[i].AuditCreateTime).ToString("yyyy/MM/dd");
					Viewmodel[i].DoSCreateTime = String.IsNullOrEmpty(Viewmodel[i].DoSCreateTime) ? "N" : Convert.ToDateTime(Viewmodel[i].DoSCreateTime).ToString("yyyy/MM/dd");
					j++;
				}
			}
			else if (User.HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX002.User"))//業務員
			{
				model.AgentCode = User.MemberInfo.ID;
				Viewmodel = _Service.QueryNotifyDataByAgent(model);
				for (int i = 0; i < Viewmodel.Count; i++)
				{
					Viewmodel[i].ViewID = j;
					Viewmodel[i].CreateTime = Convert.ToDateTime(Viewmodel[i].CreateTime).ToString("yyyy/MM/dd");
					var CRMEDoS = _Service.GetCRMEDoS(Viewmodel[i].No);
					var CRMEAudit = _Service.GetCRMEAudit(Viewmodel[i].No);
					if(CRMEDoS != null && CRMEAudit != null)
					{
						var c = DateTime.Compare(CRMEDoS.CreateTime, CRMEAudit.CreateTime);
						if(c < 0)
						{
							int s = _Service.GetCRMEAuditSep(Viewmodel[i].No);
							Viewmodel[i].State = "稽催";
							Viewmodel[i].Seq = "第" + s + "次" + Viewmodel[i].State;
							
						}else
						{
							int s = _Service.GetCRMEDoSSep(Viewmodel[i].No);
							Viewmodel[i].State = "催辦";
							Viewmodel[i].Seq = "第" + s + "次" + Viewmodel[i].State;
						}
					}
					else if (CRMEDoS == null && CRMEAudit == null)
					{
						Viewmodel[i].Seq = "照會";
					}
					else
					{
						if (CRMEDoS != null && CRMEAudit == null)
						{
							int s = _Service.GetCRMEDoSSep(Viewmodel[i].No);
							Viewmodel[i].State = "催辦";
							Viewmodel[i].Seq = "第" + s + "次" + Viewmodel[i].State;
						}
						else
						{
							int s = _Service.GetCRMEAuditSep(Viewmodel[i].No);
							Viewmodel[i].State = "稽催";
							Viewmodel[i].Seq = "第" + s + "次" + Viewmodel[i].State;
						}
					}

					j++;
				}
			}

			//WebChannel<INotifyService> _channelService = new WebChannel<INotifyService>();
			//var gridKey = _channelService.DataToCache(Viewmodel.AsEnumerable());
			//SetGridKey("QueryGrid", gridKey);
			return Json(Viewmodel);
		}

		//public JsonResult BindGrid(jqGridParam jqParams)
		//{
		//	var cacheKey = GetGridKey("QueryGrid");
		//	return BaseGridBinding<NotifyViewModel>(jqParams,
		//		() => new WebChannel<INotifyService, NotifyViewModel>().Get(cacheKey));
		//}

		/// <summary>
		/// 依類別取得類型
		/// </summary>
		/// <param name="category"></param>
		/// <returns></returns>
		[HttpPost]
		public JsonResult GetType(Category? category)
		{
			var LS = CUSCRMHelper.GetCaseTypeList(category);
			return Json(LS, JsonRequestBehavior.AllowGet);
		}

		[HttpPost]
		public ActionResult CreateMessage(NotifyCaseViewModel viewModel)
		{
			//受理通知檔
			var CRMENotifyTo = new List<CRMENotifyTo>();
			//留言清單
			var idnameNewLst = new List<UserSimpleInfo>();
			try
			{
				//通知對象人員選擇
				List<AccountGroupJsonString> JsonToLst = new List<AccountGroupJsonString>();
				if (!String.IsNullOrEmpty(viewModel.RecipientToJson))
				{
					JsonToLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(viewModel.RecipientToJson);
				}
				List<UserSimpleInfo> To = PlatformHelper.GetRecipientUsersList(JsonToLst);			
				for (int i = 0; i < To.Count; i++)
				{
					CRMENotifyTo CMT = new CRMENotifyTo();
					CMT.MemberID = To[i].MemberId;
					CMT.Creator = User.MemberInfo.ID;
					CMT.NotifyType = NotifyType.Other;
					CMT.No = viewModel.No;
					CMT.CreateTime = DateTime.Now;
					CRMENotifyTo.Add(CMT);

					UserSimpleInfo item = new UserSimpleInfo();
					item.MemberId = To[i].MemberId;
					item.Name = To[i].Name;
					item.Unit = To[i].Unit;
					idnameNewLst.Add(item);
				}

				_Service.InsertCRMENotifyTo(CRMENotifyTo);
				var MsgService = ServiceHelper.Create<IMessageService>();
				var model = _Service.GetCRMECaseContent(viewModel.No);
				var Policy = _Service.GetCRMEInsurancePolicy(viewModel.No);
				var ToLst = _Service.GetToByNo(viewModel.No);
				string ToStr = string.Empty;
				for(int i = 0;i< ToLst.Count; i++)
				{
					ToStr = i == 0 ? Member.Get(ToLst[i].MemberID).Name  : ToStr + "、" + Member.Get(ToLst[i].MemberID).Name;
				}

				//保戶
				List<string> OwnerLst = new List<string>();
				string OwnerStr = string.Empty;
				for (int i = 0; i < Policy.Count; i++)
				{
					OwnerLst.Add(Policy[i].Owner);
				}
				OwnerLst = OwnerLst.Distinct().ToList();
				string encoded = string.Empty;
				for (int i = 0; i < OwnerLst.Count; i++)
				{
					for (int j = 0; j < OwnerLst[i].Length; j++)
					{
						if (OwnerLst[i].Length < 4)
						{
							if (j == 1)
							{
								encoded += "O";
							}
							else
							{
								encoded += OwnerLst[i][j];
							}
						}
						else if (OwnerLst[i].Length == 4)
						{
							if (j == 1 || j == 2)
							{
								encoded += "O";
							}
							else
							{
								encoded += OwnerLst[i][j];
							}
						}
						else
						{
							if (j == OwnerLst[i].Length - 2 || j == OwnerLst[i].Length - 1)
							{
								encoded += "O";
							}
							else
							{
								encoded += OwnerLst[i][j];
							}
						}
					}
					OwnerStr = i == 0 ? encoded : OwnerStr + "、" + encoded;
					encoded = string.Empty;
				}

				//連結
				UrlHelper u = new UrlHelper(this.ControllerContext.RequestContext);
				string Url = "業務發展 > 業務支援 > 客服業務 > 通知作業";
				//Url = HtmlHelper.GenerateLink(this.ControllerContext.RequestContext, System.Web.Routing.RouteTable.Routes, "照會單", "", "CSNote", "CUSCRMTX002", null, null);
				if (CRMENotifyTo.Count != 0)
				{
					//留言
					Message k = new Message();
					switch (model.Type)
					{
						case "CS":
							k.MSGSubject = "保戶(" + OwnerStr + "）服務照會單通知-" + viewModel.No + "(業務員：" + ToStr + ")";
							k.MSGDESC = "通知(服務人員：" + ToStr + ")保戶" + OwnerStr + "有服務需求，請於" + Url + "下載照會單，照會單號" + viewModel.No + "，並於「五個工作日內」書面回覆總公司客服部，謝謝!";
							break;
						case "CC":
							k.MSGSubject = "保戶(" + OwnerStr + "）申訴照會單通知-" + viewModel.No + "(業務員：" + ToStr + ")";
							k.MSGDESC = "通知保戶" + OwnerStr + "提出申訴(業務員：" + ToStr + ")，請於" + Url + "下載照會單（照會單號" + viewModel.No + "）及附件，並於「五個工作日內」，以書面回覆總公司客服部，以利申訴作業進行。謝謝!";
							break;
						case "SC":
							k.MSGSubject = "業務員申訴照會單通知-" + viewModel.No + "(申訴人：" + ToStr + ")";
							k.MSGDESC = "本部受理業務員（" + ToStr + "）申訴，請於" + Url + "下載照會單（照會單號" + viewModel.No + "）及附件，並於「五個工作日內」書面回覆總公司客服部，以利申訴作業進行。謝謝!";
							break;
						default:
							break;
					}
					k.MSGCreater = User.MemberInfo.ID;
					k.MSGCreateName = User.UnitInfo.Name + "-" + User.MemberInfo.Name;
					k.MSGCreateIP = PlatformHelper.GetClientIPv4();
					k.MSGTime = DateTime.Now;
					k.MSGClass = 5;
					//產生留言人員名單
					List<MessageTo> kt_list = null;
					// 產生留言附加檔案資料
					List<MessageFile> kf_list = new List<MessageFile>();

					MessageTo partdata = new MessageTo();
					partdata.MSGOBJDate = k.MSGTime;
					partdata.MSGOBJReaderIP = "";
					partdata.MSGOBJSendIP = k.MSGCreateIP;
					partdata.MSGOBJCreateTime = k.MSGTime;
					partdata.MSGOBJSendID = User.MemberInfo.ID;

					kt_list = GetMsgIdsToMsgTList(idnameNewLst, partdata);
					MsgService.CreateMessage(k, kf_list, kt_list, viewModel.TabUniqueId, out string retMsg);

					for (int i = 0; i < idnameNewLst.Count; i++)
					{
						if (idnameNewLst[i].MemberId.Length == 12)
						{
							idnameNewLst[i].AccountId = idnameNewLst[i].MemberId.Substring(0, 10);
						}
						else
						{
							idnameNewLst[i].AccountId = _Service.GetAccountId(idnameNewLst[i].MemberId);
						}
					}

					//推播
					_Service.SendLineNotify(k, idnameNewLst);
				}
				AppendMessage("新增成功", true);
				return RedirectToAction("index", "CUSCRMTX002");
			}
			catch (Exception ex)
			{
				AppendMessage("新增失敗");
				return View("Notify", viewModel);
			}
		}
		#endregion

		#region 立案通知業務員
		/// <summary>
		/// 立案通知業務員Detail
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public ActionResult Notify(string No)
		{
			NotifyCaseViewModel viewModel = new NotifyCaseViewModel();
			List<string> WcCenterList = new List<string>();
			List<Tuple<string, string>> WcCenterViewList = new List<Tuple<string, string>>();
			viewModel.No = No;
			viewModel.NotifyInsurance = _Service.GetCRMEInsurancePolicy(No);
			viewModel.ChinaOM = _Service.GetCRMEChinaOm(No);
			Dictionary<string, string> WcCenterName = _Service.GetWcCenterNameByAgentCode();
			Dictionary<string, string> LevelName = _Service.GetLevelNameChsByAgentCode();
			Dictionary<string, string> WcCenter = _Service.GetWcCenterByAgentCode();
			for(int i = 0;i < viewModel.ChinaOM.Count; i++)
			{
				WcCenterName.TryGetValue(viewModel.ChinaOM[i].CHAgentCode ?? "", out string CHAgentCodeWcCenterName);
				LevelName.TryGetValue(viewModel.ChinaOM[i].CHAgentCode ?? "", out string CHAgentCodeLevelName);
				WcCenter.TryGetValue(viewModel.ChinaOM[i].CHAgentCode ?? "", out string CHAgentCodeWcCenter);
				WcCenterList.Add(CHAgentCodeWcCenter);
				viewModel.ChinaOM[i].CHLvName = CHAgentCodeWcCenterName + " " + viewModel.ChinaOM[i].CHAgentName + " " + CHAgentCodeLevelName;
			}
			for (int i = 0; i < viewModel.NotifyInsurance.Count; i++)
			{
				WcCenterName.TryGetValue(viewModel.NotifyInsurance[i].AgentCode ?? "", out string AgentCodeWcCenterName);
				WcCenterName.TryGetValue(viewModel.NotifyInsurance[i].VMLeaderID ?? "", out string VMLeaderIDWcCenterName);
				WcCenterName.TryGetValue(viewModel.NotifyInsurance[i].SMLeaderID ?? "", out string SMLeaderIDWcCenterName);
				WcCenterName.TryGetValue(viewModel.NotifyInsurance[i].CenterLeaderID ?? "", out string CenterLeaderIDWcCenterName);
				WcCenterName.TryGetValue(viewModel.NotifyInsurance[i].SUAgentCode ?? "", out string SUAgentCodeWcCenterName);
				WcCenterName.TryGetValue(viewModel.NotifyInsurance[i].SUVMLeaderID ?? "", out string SUVMLeaderIDWcCenterName);
				WcCenterName.TryGetValue(viewModel.NotifyInsurance[i].SUSMLeaderID ?? "", out string SUSMLeaderIDWcCenterName);
				WcCenterName.TryGetValue(viewModel.NotifyInsurance[i].SUCenterLeaderID ?? "", out string SUCenterLeaderIDWcCenterName);

				LevelName.TryGetValue(viewModel.NotifyInsurance[i].AgentCode ?? "", out string AgentCodeLevelName);
				LevelName.TryGetValue(viewModel.NotifyInsurance[i].VMLeaderID ?? "", out string VMLeaderIDLevelName);
				LevelName.TryGetValue(viewModel.NotifyInsurance[i].SMLeaderID ?? "", out string SMLeaderIDLevelName);
				LevelName.TryGetValue(viewModel.NotifyInsurance[i].CenterLeaderID ?? "", out string CenterLeaderIDLevelName);
				LevelName.TryGetValue(viewModel.NotifyInsurance[i].SUAgentCode ?? "", out string SUAgentCodeLevelName);
				LevelName.TryGetValue(viewModel.NotifyInsurance[i].SUVMLeaderID ?? "", out string SUVMLeaderIDLevelName);
				LevelName.TryGetValue(viewModel.NotifyInsurance[i].SUSMLeaderID ?? "", out string SUSMLeaderIDLevelName);
				LevelName.TryGetValue(viewModel.NotifyInsurance[i].SUCenterLeaderID ?? "", out string SUCenterLeaderIDLevelName);

				WcCenter.TryGetValue(viewModel.NotifyInsurance[i].AgentCode ?? "", out string AgentCodeWcCenter);
				WcCenter.TryGetValue(viewModel.NotifyInsurance[i].VMLeaderID ?? "", out string VMLeaderIDWcCenter);
				WcCenter.TryGetValue(viewModel.NotifyInsurance[i].SMLeaderID ?? "", out string SMLeaderIDWcCenter);
				WcCenter.TryGetValue(viewModel.NotifyInsurance[i].CenterLeaderID ?? "", out string CenterLeaderIDWcCenter);
				WcCenter.TryGetValue(viewModel.NotifyInsurance[i].SUAgentCode ?? "", out string SUAgentCodeWcCenter);
				WcCenter.TryGetValue(viewModel.NotifyInsurance[i].SUVMLeaderID ?? "", out string SUVMLeaderIDWcCenter);
				WcCenter.TryGetValue(viewModel.NotifyInsurance[i].SUSMLeaderID ?? "", out string SUSMLeaderIDWcCenter);
				WcCenter.TryGetValue(viewModel.NotifyInsurance[i].SUCenterLeaderID ?? "", out string SUCenterLeaderIDWcCenter);

				if (!string.IsNullOrWhiteSpace(viewModel.NotifyInsurance[i].SUAgentCode))
				{
					viewModel.NotifyInsurance[i].SuLvName = SUAgentCodeWcCenterName + " " + viewModel.NotifyInsurance[i].SUAgentName + " " + SUAgentCodeLevelName;
					viewModel.NotifyInsurance[i].SuVmLvName = SUVMLeaderIDWcCenterName + " " + viewModel.NotifyInsurance[i].SUVMLeader + " " + SUVMLeaderIDLevelName;
					viewModel.NotifyInsurance[i].SuSmLvName = SUSMLeaderIDWcCenterName + " " + viewModel.NotifyInsurance[i].SUSMLeader + " " + SUSMLeaderIDLevelName;
					viewModel.NotifyInsurance[i].SuCLvName = SUCenterLeaderIDWcCenterName + " " + viewModel.NotifyInsurance[i].SUCenterLeader + " " + SUCenterLeaderIDLevelName;
					WcCenterList.Add(SUAgentCodeWcCenter);
					WcCenterList.Add(SUVMLeaderIDWcCenter);
					WcCenterList.Add(SUSMLeaderIDWcCenter);
					WcCenterList.Add(SUCenterLeaderIDWcCenter);
				}
				else
				{
					viewModel.NotifyInsurance[i].AgLvName = AgentCodeWcCenterName + " " + viewModel.NotifyInsurance[i].AgentName + " " + AgentCodeLevelName;
					viewModel.NotifyInsurance[i].AgVmLvName = VMLeaderIDWcCenterName + " " + viewModel.NotifyInsurance[i].VMLeader + " " + VMLeaderIDLevelName;
					viewModel.NotifyInsurance[i].AgSmLvName = SMLeaderIDWcCenterName + " " + viewModel.NotifyInsurance[i].SMLeader + " " + SMLeaderIDLevelName;
					viewModel.NotifyInsurance[i].AgCLvName = CenterLeaderIDWcCenterName + " " + viewModel.NotifyInsurance[i].CenterLeader + " " + CenterLeaderIDLevelName;
					WcCenterList.Add(AgentCodeWcCenter);
					WcCenterList.Add(VMLeaderIDWcCenter);
					WcCenterList.Add(SMLeaderIDWcCenter);
					WcCenterList.Add(CenterLeaderIDWcCenter);
				}
			}
			WcCenterList = WcCenterList.Distinct().ToList();
			for (int i = 0; i < WcCenterList.Count; i++)
			{
				var model = new List<NotifyCaseViewModel>();
				model = _Service.GetCRMENotifyEmployee(WcCenterList[i]);
				foreach (var item in model)
				{
					WcCenterViewList.Add(new Tuple<string, string>(item.PeopleOrgid, item.PeopleName));
				}
			}
			viewModel.NotifyEmployee = WcCenterViewList.Distinct().ToList();
			viewModel.DoUser = User.MemberInfo.Name;

			return View(viewModel);
		}

		/// <summary>
		/// 新增通知
		/// </summary>
		/// <param name="viewModel"></param>
		/// <returns></returns>
		[HttpPost]
		[PdLogFilter("EIP客服-新增通知", PITraceType.Insert)]
		public ActionResult Create(NotifyCaseViewModel viewModel)
		{
			//受理通知檔
			var CRMENotifyTo = new List<CRMENotifyTo>();
			//留言清單
			var idnameNewLst = new List<UserSimpleInfo>();
			try
			{
				//受文人員選擇
				List<AccountGroupJsonString> JsonToLst = new List<AccountGroupJsonString>();
				if (!String.IsNullOrEmpty(viewModel.RecipientToJson))
				{
					JsonToLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(viewModel.RecipientToJson);
				}
				List<UserSimpleInfo> To = PlatformHelper.GetRecipientUsersList(JsonToLst);
				//受文CheckBox
				if (viewModel.ToCheckbox != null)
				{
					for (int i = 0; i < viewModel.ToCheckbox.Length; i++)
					{
						UserSimpleInfo item = new UserSimpleInfo();
						item.MemberId = viewModel.ToCheckbox[i].ToString();
						var M = Member.Get(item.MemberId) ?? null;
						if (M != null)
						{
							item.Name = M.Name;
							var U = Unit.Get(M.UnitGUID);
							item.Unit = U.Name;
						}
						To.Add(item);
					}
				}
				string ToStr = string.Empty;
				for (int i = 0; i < To.Count; i++)
				{
					CRMENotifyTo CMT = new CRMENotifyTo();
					CMT.MemberID = To[i].MemberId;
					CMT.Creator = viewModel.DoUser;
					CMT.NotifyType = NotifyType.To;
					CMT.No = viewModel.No;
					CMT.CreateTime = DateTime.Now;
					CRMENotifyTo.Add(CMT);

					UserSimpleInfo item = new UserSimpleInfo();
					item.MemberId = To[i].MemberId;
					item.Name = To[i].Name;
					item.Unit = To[i].Unit;
					idnameNewLst.Add(item);

					ToStr = i == 0 ? To[i].Name : ToStr + "、" + To[i].Name;
				}

				//副本人員選擇
				List<AccountGroupJsonString> JsonCCLst = new List<AccountGroupJsonString>();
				if (!String.IsNullOrEmpty(viewModel.RecipientCCJson))
				{
					JsonCCLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(viewModel.RecipientCCJson);
				}
				List<UserSimpleInfo> CC = PlatformHelper.GetRecipientUsersList(JsonCCLst);
				//副本CheckBox
				if (viewModel.CCCheckbox != null)
				{
					for (int i = 0; i < viewModel.CCCheckbox.Length; i++)
					{
						UserSimpleInfo item = new UserSimpleInfo();
						item.MemberId = viewModel.CCCheckbox[i].ToString();
						var M = Member.Get(item.MemberId) ?? null;
						if (M != null)
						{
							item.Name = M.Name;
							var U = Unit.Get(M.UnitGUID);
							item.Unit = U.Name;
						}
						CC.Add(item);
					}
				}
				for (int i = 0; i < CC.Count; i++)
				{
					CRMENotifyTo CMT = new CRMENotifyTo();
					CMT.MemberID = CC[i].MemberId;
					CMT.Creator = viewModel.DoUser;
					CMT.NotifyType = NotifyType.CC;
					CMT.No = viewModel.No;
					CMT.CreateTime = DateTime.Now;
					CRMENotifyTo.Add(CMT);

					UserSimpleInfo item = new UserSimpleInfo();
					item.MemberId = CC[i].MemberId;
					item.Name = CC[i].Name;
					item.Unit = CC[i].Unit;
					idnameNewLst.Add(item);
				}

				//行專人員選擇
				List<AccountGroupJsonString> JsonEPLst = new List<AccountGroupJsonString>();
				if (!String.IsNullOrEmpty(viewModel.RecipientEPJson))
				{
					JsonEPLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(viewModel.RecipientEPJson);
				}
				List<UserSimpleInfo> EP = PlatformHelper.GetRecipientUsersList(JsonEPLst);
				//行專人員CheckBox
				if (viewModel.EPCheckbox != null)
				{
					for (int i = 0; i < viewModel.EPCheckbox.Length; i++)
					{
						UserSimpleInfo item = new UserSimpleInfo();
						var v = _MemberEService.GetMemberExtendIDbyExtendID(viewModel.EPCheckbox[i].ToString());
						item.MemberId = v.MemberID ?? null;
						var M = Member.Get(item.MemberId) ?? null;
						if (M != null)
						{
							item.Name = M.Name;
							var U = Unit.Get(M.UnitGUID);
							item.Unit = U.Name;
						}
						EP.Add(item);
					}
				}
				for (int i = 0; i < EP.Count; i++)
				{
					CRMENotifyTo CMT = new CRMENotifyTo();
					CMT.MemberID = EP[i].MemberId;
					CMT.Creator = viewModel.DoUser;
					CMT.NotifyType = NotifyType.Employee;
					CMT.No = viewModel.No;
					CMT.CreateTime = DateTime.Now;
					CRMENotifyTo.Add(CMT);

					UserSimpleInfo item = new UserSimpleInfo();
					item.MemberId = EP[i].MemberId;
					item.Name = EP[i].Name;
					item.Unit = EP[i].Unit;
					idnameNewLst.Add(item);
				}

				List<UserSimpleInfo> OMP = new List<UserSimpleInfo>();
				//處代理人CheckBox
				if (viewModel.OMProxyCheckbox != null)
				{
					for (int i = 0; i < viewModel.OMProxyCheckbox.Length; i++)
					{
						UserSimpleInfo item = new UserSimpleInfo();
						item.MemberId = viewModel.OMProxyCheckbox[i].ToString();
						var M = Member.Get(item.MemberId) ?? null;
						if (M != null)
						{
							item.Name = M.Name;
							var U = Unit.Get(M.UnitGUID);
							item.Unit = U.Name;
						}
						OMP.Add(item);
					}
				}
				for (int i = 0; i < OMP.Count; i++)
				{
					CRMENotifyTo CMT = new CRMENotifyTo();
					CMT.MemberID = OMP[i].MemberId;
					CMT.Creator = viewModel.DoUser;
					CMT.NotifyType = NotifyType.OMProxy;
					CMT.No = viewModel.No;
					CMT.CreateTime = DateTime.Now;
					CRMENotifyTo.Add(CMT);

					UserSimpleInfo item = new UserSimpleInfo();
					item.MemberId = OMP[i].MemberId;
					item.Name = OMP[i].Name;
					item.Unit = OMP[i].Unit;
					idnameNewLst.Add(item);
				}

				//去除重複人選   
				List<UserSimpleInfo> uniqueUsers = idnameNewLst
					.GroupBy(uu => uu.MemberId) // 按照MemberId进行分组
					.Select(g => g.First())   // 选择每个分组的第一个元素
					.ToList();

				CRMENotifyTo = CRMENotifyTo
					.GroupBy(uu => uu.MemberID) 
					.Select(g => g.First())   
					.ToList();

				//主檔
				CRMECaseContent content = new CRMECaseContent();
				content.Status = Convert.ToInt32(ContentStatus.Notified);
				content.No = viewModel.No;
				content.DoUser = viewModel.DoUser;
				content.DoUserTelExt = viewModel.DoUserTelExt ?? "";

				List<CRMEFile> F_list = new List<CRMEFile>();
				if (viewModel.UploadFilesName != null && viewModel.UploadFilesName.Length > 0)
				{
					string fnames = viewModel.UploadFilesName;

					F_list = (from item in fnames.Split('*')
							  let parts = item.Split('|')
							  select new CRMEFile { FileName = parts[0], FileMD5Name = parts[1] }).ToList();
					for (int i = 0; i < F_list.Count; i++)
					{
						F_list[i].FolderNo = viewModel.No;
						F_list[i].Creator = User.MemberInfo.ID;
						F_list[i].CreateTime = DateTime.Now;
						F_list[i].SourceType = 0;
						F_list[i].FileNo = "CUSCRMTX002";
					}
				}

				_Service.InsertCRMENotifyTo(CRMENotifyTo);
				_Service.UpdateCaseContentType(content);
				var Policy = _Service.GetCRMEInsurancePolicy(viewModel.No);
				var model = _Service.GetCRMECaseContent(viewModel.No);
				_Service.CreateCRMNotifyFile(model, F_list, viewModel.TabUniqueId, out string msg);
				var MsgService = ServiceHelper.Create<IMessageService>();

				//保戶
				List<string> OwnerLst = new List<string>();
				string OwnerStr = string.Empty;
				for (int i = 0; i < Policy.Count; i++)
				{
					OwnerLst.Add(Policy[i].Owner);
				}
				OwnerLst = OwnerLst.Distinct().ToList();
				string encoded = string.Empty;
				for (int i = 0; i < OwnerLst.Count; i++)
				{
					for (int j = 0; j < OwnerLst[i].Length; j++)
					{
						if (OwnerLst[i].Length < 4)
						{
							if (j == 1)
							{
								encoded += "O";
							}
							else
							{
								encoded += OwnerLst[i][j];
							}
						}
						else if (OwnerLst[i].Length == 4)
						{
							if (j == 1 || j == 2)
							{
								encoded += "O";
							}
							else
							{
								encoded += OwnerLst[i][j];
							}
						}
						else
						{
							if (j == OwnerLst[i].Length - 2 || j == OwnerLst[i].Length - 1)
							{
								encoded += "O";
							}
							else
							{
								encoded += OwnerLst[i][j];
							}
						}
					}
					OwnerStr = i == 0 ? encoded : OwnerStr + "、" + encoded;
					encoded = string.Empty;
				}

				//連結
				UrlHelper u = new UrlHelper(this.ControllerContext.RequestContext);
				string Url = "業務發展 > 業務支援 > 客服業務 > 通知作業";
				//Url = HtmlHelper.GenerateLink(this.ControllerContext.RequestContext, System.Web.Routing.RouteTable.Routes, "照會單", "", "CSNote", "CUSCRMTX002", null, null);
				if (CRMENotifyTo.Count != 0)
				{
					//留言
					Message k = new Message();
					switch (model.Type)
					{
						case "CS":
							k.MSGSubject = "保戶(" + OwnerStr + "）服務照會單通知-" + viewModel.No + "(業務員：" + ToStr + ")";
							k.MSGDESC = "通知(服務人員：" + ToStr + ")保戶" + OwnerStr + "有服務需求，請於" + Url + "下載照會單，照會單號" + viewModel.No + "，並於「五個工作日內」書面回覆總公司客服部，謝謝!";
							break;
						case "CC":
							k.MSGSubject = "保戶(" + OwnerStr + "）申訴照會單通知-" + viewModel.No + "(業務員：" + ToStr + ")";
							k.MSGDESC = "通知保戶" + OwnerStr + "提出申訴(業務員：" + ToStr + ")，請於" + Url + "下載照會單（照會單號" + viewModel.No + "）及附件，並於「五個工作日內」，以書面回覆總公司客服部，以利申訴作業進行。謝謝!";
							break;
						case "SC":
							k.MSGSubject = "業務員申訴照會單通知-" + viewModel.No + "(申訴人：" + ToStr + ")";
							k.MSGDESC = "本部受理業務員（" + ToStr + "）申訴，請於" + Url + "下載照會單（照會單號" + viewModel.No + "）及附件，並於「五個工作日內」書面回覆總公司客服部，以利申訴作業進行。謝謝!";
							break;
						default:
							break;
					}
					k.MSGCreater = viewModel.DoUser;
					k.MSGCreateName = Unit.Get(Member.Get(viewModel.DoUser).UnitGUID).Name + "-" + Member.Get(viewModel.DoUser).Name;
					k.MSGCreateIP = PlatformHelper.GetClientIPv4();
					k.MSGTime = DateTime.Now;
					k.MSGClass = 5;
					//產生留言人員名單
					List<MessageTo> kt_list = null;
					// 產生留言附加檔案資料
					List<MessageFile> kf_list = new List<MessageFile>();

					MessageTo partdata = new MessageTo();
					partdata.MSGOBJDate = k.MSGTime;
					partdata.MSGOBJReaderIP = "";
					partdata.MSGOBJSendIP = k.MSGCreateIP;
					partdata.MSGOBJCreateTime = k.MSGTime;
					partdata.MSGOBJSendID = viewModel.DoUser;

					kt_list = GetMsgIdsToMsgTList(uniqueUsers, partdata);
					MsgService.CreateMessage(k, kf_list, kt_list, viewModel.TabUniqueId, out string retMsg);

					for(int i = 0;i< uniqueUsers.Count; i++)
					{
						if(uniqueUsers[i].MemberId.Length == 12)
						{
							uniqueUsers[i].AccountId = uniqueUsers[i].MemberId.Substring(0, 10);
						}
						else
						{
							uniqueUsers[i].AccountId =_Service.GetAccountId(uniqueUsers[i].MemberId);
						}
					}

					//推播
					//需求單20240711003，改只推播給業務員，不給行專行政人員
					List<UserSimpleInfo> uniqueInfos = new List<UserSimpleInfo>();
					foreach (UserSimpleInfo info in uniqueUsers) 
					{
						if (info.Unit != "行政專員") { uniqueInfos.Add(info); }
					}
					_Service.SendLineNotify(k, uniqueInfos);
				}
				AppendMessage("新增成功", true);
				return RedirectToAction("index", "CUSCRMTX002");
			}
			catch (Exception ex)
			{
				AppendMessage("新增失敗");
				return View("Notify", viewModel);
			}			
		}
		/// <summary>
		/// 保單資料
		/// </summary>
		/// <param name="model">單筆調整view model</param>
		/// <returns></returns>
		//      [HttpPost]
		//public JsonResult Query()
		//{
		//	List<LawSys> QResultList = new List<LawSys>();
		//	var mService = new WebChannel<INotifyService>();

		//	mService.Use(service => service
		//	.GetLawSys()
		//	.ForEach(d =>
		//	{
		//		if (d != null)
		//		{
		//			var item = new LawSys();
		//			item.ID = d.ID;
		//			item.SysName = d.SysName;

		//			QResultList.Add(item);
		//		}
		//	}));

		//	return Json(QResultList, JsonRequestBehavior.AllowGet);
		//}
		#endregion

		#region 立案通知申訴人
		/// <summary>
		/// 取得已立案未通知的清單資料
		/// </summary>
		/// <returns>已立案未通知的清單資料</returns>
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult GetCCWaitNofity()
		{
			var result = caseService.GetCCWaitNofityDatas();
			return Json(result);
		}

		/// <summary>
		/// 立案通知申訴人
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public ActionResult CRMEAppeal(string No)
		{
			var model = _Service.GetCRMECaseContent(No);
			if (No == null || model == null)
			{
				AppendMessage("查無案件編號與相關資料，請重新確認！");
				return View("Index");
			}
			try
			{				
				ViewBag.No = No;
				ViewBag.Status = model.Status;
				ViewBag.AppealName = model.Complainant;

				return View();
			}
			catch (Exception ex)
			{
				AppendMessage("查無案件編號與相關資料，請重新確認！");
				return View("Index");
			}			
        }

        [HttpPost]
		[PdLogFilter("EIP客服-通知申訴人", PITraceType.Insert)]
		public ActionResult CRMEAppeal(CRMEAppealBy model)
        {
			try
			{
				var service = ServiceHelper.Create<IDataSettingService>();
				var Casemodel = _Service.GetCRMECaseContent(model.No);
				ViewBag.Status = Casemodel.Status;

				//申訴案件進度查詢
				string URL_Todo = "<a href='" + PlatformHelper.GetVarConfig("AppealBYUrl") + "'>" + service.GetVarConfigDetailModelByName("AppealBYUrl").VarDesc + "</a>";

				string DoUserFirstName;
				if (model.DoUserFirstName != null)
				{
					DoUserFirstName = model.DoUserFirstName;
				}
				else
				{
					DoUserFirstName = Member.Get(model.DoUser).Name.Substring(0, 1);
				}
				var appealTemplate = @"{0} 先生/小姐您好，本公司已受理台端申訴事項，案件編號：{1}，您可至{2}查詢案件進度，如有問題請隨時來電永達總公司專線 02-25212019 分機 {3}{4}{5}，將有專人為您服務。 ";
				var entrustdTemplate = @"{0} 先生/小姐您好，台端受{1}委託處理申訴事宜，本公司業已受理，案件編號：{2}，您可至{3}查詢案件進度，如有問題請隨時來電永達總公司專線 02-25212019 分機 {4}{5}{6}，將有專人為您服務。";

				//申訴人
				var AppealMobile_content = string.Format(appealTemplate, model.AppealName, model.No, URL_Todo, model.DoUserTelExt, DoUserFirstName, model.Title);

				var AppealEmail_content = AppealMobile_content;


				//受任人
				var EntrustdMobile_content = string.Format(entrustdTemplate, model.EntrustdName, model.AppealName, model.No, URL_Todo, model.DoUserTelExt, DoUserFirstName, model.Title);

				var EntrustdEmail_content = EntrustdMobile_content;


				CRMEAppealBy CRMEAppealBy = new CRMEAppealBy();
				CRMEAppealBy.No = model.No;
				CRMEAppealBy.AppealName = model.AppealName;
				CRMEAppealBy.AppealMobile = model.AppealMobile;
				CRMEAppealBy.AppealMobile_Content = AppealMobile_content;   //model.AppealMobile_Content;
				if (model.AppealEmail == null)
				{
					CRMEAppealBy.AppealEmail = "";
					CRMEAppealBy.AppealEmail_Content = "";     // model.AppealEmail_Content;
				}
				else 
				{ 
				CRMEAppealBy.AppealEmail = model.AppealEmail;
				CRMEAppealBy.AppealEmail_Content = AppealEmail_content;     // model.AppealEmail_Content;
				}
				if (model.EntrustdName == null)
				{
					CRMEAppealBy.EntrustdName = "";
					CRMEAppealBy.EntrustdMobile = "";
					CRMEAppealBy.EntrustdMobile_Content = "";   
					CRMEAppealBy.EntrustdEmail = "";
					CRMEAppealBy.EntrustdEmail_Content = "";   
                }
                else 
				{
					CRMEAppealBy.EntrustdName = model.EntrustdName;
					CRMEAppealBy.EntrustdMobile = model.EntrustdMobile;
					CRMEAppealBy.EntrustdMobile_Content = EntrustdMobile_content;	// model.EntrustdMobile_Content;
					CRMEAppealBy.EntrustdEmail = model.EntrustdEmail;
					CRMEAppealBy.EntrustdEmail_Content = EntrustdEmail_content;     // model.EntrustdEmail_Content;
				}
				CRMEAppealBy.Title = model.Title;
				CRMEAppealBy.DoUser = model.DoUser;
				CRMEAppealBy.DoUserFirstName = model.DoUserFirstName;
				CRMEAppealBy.DoUserTelExt = model.DoUserTelExt;
				CRMEAppealBy.Creator = User.MemberInfo.ID; 
				CRMEAppealBy.CreateTime = DateTime.Now;

				_Service.InsertCRMEAppealBy(CRMEAppealBy);


                if (model.AppealMobile != null)
                {
                    string SmsMsg = string.Empty;
                    SmsMsg = AppealMobile_content;
					string ErrMsg = string.Empty;

					SendSmsMessage(model.AppealMobile, SmsMsg, User.Property.LoginId, out ErrMsg);
					
					if (model.AppealEmail != null)
					{ 
						SendEmail(model.AppealEmail, AppealEmail_content); 
					}
						

				}

                if (model.EntrustdMobile != null)
                {
                    string SmsMsg = string.Empty;
                    SmsMsg = EntrustdMobile_content;
					string ErrMsg = string.Empty;

					SendSmsMessage(model.EntrustdMobile, SmsMsg, User.Property.LoginId, out ErrMsg);
					SendEmail(model.EntrustdEmail, EntrustdEmail_content);
				}

                AppendMessage("新增成功", true);

				if (ViewBag.Status == 0)
					return RedirectToAction("Notify", new { model.No });
				else
					return RedirectToAction("Index", "CUSCRMTX001");
			}
			catch (Exception ex)
			{
				AppendMessage("新增失敗");
				return RedirectToActionPermanent("CRMEAppeal", new { No = model.No });
			}


		}
		#endregion

		#region 發送簡訊
		//EIP客服客戶申訴通知
		[PdLogFilter("EIP客服-通知作業-申訴簡訊", PITraceType.Insert)]
		private bool SendSmsMessage(string postSms, string smsMessage, string accountId, out string retMsg)
		{
			retMsg = "";
			bool retBl = false;
			var service = ServiceHelper.Create<IAccountSecurityService>();

			try
			{
				retBl = service.SendTwSMS(postSms, smsMessage, accountId, "EIP客服客戶申訴通知");
			}
			catch (Exception ex)
			{
				retMsg = ex.Message;
			}

			return retBl;
		}
		#endregion

		#region 發送驗證電子郵件
		//發送電子郵件 
		[PdLogFilter("EIP客服-通知作業-申訴郵件", PITraceType.Insert)]
		private string SendEmail(string mailto, string mailBody)
		{
			string result = "";

			SmtpClient client = new SmtpClient("mail.everprobks.com.tw", 587);
			client.Credentials = new System.Net.NetworkCredential("service", "Yt55^Sfq9xS");
			client.DeliveryMethod = SmtpDeliveryMethod.Network;

			int InsertRecID = 0;

			System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
			msg.To.Add(mailto);
			msg.From = new MailAddress(@"service@mail.everprobks.com.tw", "永達保險經紀人");
			msg.Subject = "永達保經-申訴受理回覆";//主旨
			msg.IsBodyHtml = true;//是否是HTML郵件 
			msg.Body = @"<p style = 'color: black';font-size: 20;> " + mailBody + "</p>";

			try
			{
				client.Send(msg);

				if (InsertRecID > 0)
				{

				}
			}
			catch (Exception ex)
			{
				result = ex.ToString();
			}
			return result;
		}
		#endregion



		#region 照會單
		/// <summary>
		/// CC照會單
		/// </summary>
		/// <param name="CaseGuid"></param>
		/// <returns></returns>
		[PdLogFilter("EIP客服-通知作業-CC照會單", PITraceType.Download)]
		public ActionResult CCNote(string CaseGuid)
		{
			var content = _Service.QueryNotifyDataByGuid(CaseGuid);
			NotifyReportModel reportModel = new NotifyReportModel();
			CultureInfo taiwanCulture = new CultureInfo("zh-TW");

			reportModel.BizContract = content.BizContract ?? "";
			reportModel.SolicitingRpt = content.SolicitingRpt ?? "";
			reportModel.SourceDesc = content.SourceDesc ?? "";
			reportModel.Complainant = content.Complainant ?? "";
			reportModel.CreateTime = content.CreateTime.ToString("yyyy/MM/dd tt hh:mm:ss", taiwanCulture);
			reportModel.CreateFiveTime = content.CreateTime.AddDays(5).ToString("yyyy/MM/dd");
			reportModel.DoUserName = !String.IsNullOrEmpty(content.DoUser) ? Member.Get(content.DoUser).Name : "";
			reportModel.DoUserTelExt = content.DoUserTelExt;

			var To = _Service.QueryNotifyToByNo(content.No);
			Dictionary<string, string> LevelName = _Service.GetLevelNameChsByAgentCode();
			List<string> namelist = new List<string>();

			for (int i = 0; i < To.Count; i++)
			{
				LevelName.TryGetValue(To[i].MemberID ?? "", out string AgentCodeLevelName);
				AgentCodeLevelName = String.IsNullOrEmpty(AgentCodeLevelName) ? "" : AgentCodeLevelName.Trim();
				namelist.Add((Member.Get(To[i].MemberID) == null ? _Service.GetAgentName(To[i].MemberID) : Member.Get(To[i].MemberID).Name) + " " + AgentCodeLevelName);
			}
			reportModel.To = namelist;

			return View(reportModel);
		}

		/// <summary>
		/// CS照會單
		/// </summary>
		/// <param name="CaseGuid"></param>
		/// <returns></returns>
		[PdLogFilter("EIP客服-通知作業-CS照會單", PITraceType.Download)]
		public ActionResult CSNote(string CaseGuid)
		{
			var content = _Service.QueryNotifyDataByGuid(CaseGuid);
			NotifyReportModel reportModel = new NotifyReportModel();
			CultureInfo taiwanCulture = new CultureInfo("zh-TW");
			reportModel.CreateTime = content.CreateTime.ToString("yyyy/MM/dd tt hh:mm:ss", taiwanCulture);
			reportModel.CreateFiveTime = content.CreateTime.AddDays(5).ToString("yyyy/MM/dd");
			reportModel.DoUserName = !String.IsNullOrEmpty(content.DoUser) ? Member.Get(content.DoUser).Name : "";
			reportModel.Content = content.Content;
			reportModel.NotifyInsuranceViews = _Service.GetCRMEInsurancePolicy(content.No);
			reportModel.No = content.No;
			reportModel.DoUserTelExt = content.DoUserTelExt;

			var To = _Service.GetToByNo(content.No);
			var CC = _Service.GetCCByNo(content.No);
			var OM = _Service.GetOMByNo(content.No);
			Dictionary<string, string> LevelName = _Service.GetLevelNameChsByAgentCode();
			Dictionary<string, string> WcCenterName = _Service.GetWcCenterNameByAgentCode();
			List<string> namelist = new List<string>();
			List<string> CCnamelist = new List<string>();
			List<string> OMnamelist = new List<string>();

			for (int i = 0; i < To.Count; i++)
			{
				WcCenterName.TryGetValue(To[i].MemberID ?? "", out string AgentCodeWcCenterName);
				LevelName.TryGetValue(To[i].MemberID ?? "", out string AgentCodeLevelName);
				AgentCodeWcCenterName = String.IsNullOrEmpty(AgentCodeWcCenterName) ? "" : AgentCodeWcCenterName;
				AgentCodeLevelName = String.IsNullOrEmpty(AgentCodeLevelName) ? "" : AgentCodeLevelName.Trim();
				namelist.Add(AgentCodeWcCenterName + " " + (Member.Get(To[i].MemberID) == null ? _Service.GetAgentName(To[i].MemberID) : Member.Get(To[i].MemberID).Name) + " " + AgentCodeLevelName);
			}
			reportModel.To = namelist;

			for (int i = 0; i < CC.Count; i++)
			{
				WcCenterName.TryGetValue(CC[i].MemberID ?? "", out string CCAgentCodeWcCenterName);
				LevelName.TryGetValue(CC[i].MemberID ?? "", out string CCAgentCodeLevelName);
				CCAgentCodeWcCenterName = String.IsNullOrEmpty(CCAgentCodeWcCenterName) ? "" : CCAgentCodeWcCenterName;
				CCAgentCodeLevelName = String.IsNullOrEmpty(CCAgentCodeLevelName) ? "" : CCAgentCodeLevelName.Trim();
				CCnamelist.Add(CCAgentCodeWcCenterName + " " + (Member.Get(CC[i].MemberID) == null ? _Service.GetAgentName(CC[i].MemberID) : Member.Get(CC[i].MemberID).Name) + " " + CCAgentCodeLevelName);
			}
			reportModel.CC = CCnamelist;

			for (int i = 0; i < OM.Count; i++)
			{
				WcCenterName.TryGetValue(OM[i].MemberID ?? "", out string OMAgentCodeWcCenterName);
				LevelName.TryGetValue(OM[i].MemberID ?? "", out string OMAgentCodeLevelName);
				OMAgentCodeWcCenterName = String.IsNullOrEmpty(OMAgentCodeWcCenterName) ? "" : OMAgentCodeWcCenterName;
				OMAgentCodeLevelName = String.IsNullOrEmpty(OMAgentCodeLevelName) ? "" : OMAgentCodeLevelName.Trim();
				OMnamelist.Add(OMAgentCodeWcCenterName + " " + (Member.Get(OM[i].MemberID) == null ? _Service.GetAgentName(OM[i].MemberID) : Member.Get(OM[i].MemberID).Name) + " " + "處代理人");
			}
			reportModel.OM = OMnamelist;
			
			return View(reportModel);
		}

		/// <summary>
		/// SC照會單
		/// </summary>
		/// <param name="CaseGuid"></param>
		/// <returns></returns>
		[PdLogFilter("EIP客服-通知作業-SC照會單", PITraceType.Download)]
		public ActionResult SCNote(string CaseGuid)
		{
			var content = _Service.QueryNotifyDataByGuid(CaseGuid);
			NotifyReportModel reportModel = new NotifyReportModel();
			CultureInfo taiwanCulture = new CultureInfo("zh-TW");

			reportModel.BizContract = content.BizContract ?? "";
			reportModel.SolicitingRpt = content.SolicitingRpt ?? "";
			reportModel.SourceDesc = content.SourceDesc ?? "";
			reportModel.Complainant = content.Complainant ?? "";
			reportModel.CreateTime = content.CreateTime.ToString("yyyy/MM/dd tt hh:mm:ss", taiwanCulture);
			reportModel.CreateFiveTime = content.CreateTime.AddDays(5).ToString("yyyy/MM/dd");
			reportModel.DoUserName = !String.IsNullOrEmpty(content.DoUser) ? Member.Get(content.DoUser).Name : "";
			reportModel.DoUserTelExt = content.DoUserTelExt;
			//reportModel.SourceName = _Service.GetDiscipTypeByID(Convert.ToInt32(content.SourceID));

			var To = _Service.QueryNotifyToByNo(content.No);
			Dictionary<string, string> LevelName = _Service.GetLevelNameChsByAgentCode();
			List<string> namelist = new List<string>();
			for (int i = 0; i < To.Count; i++)
			{
				LevelName.TryGetValue(To[i].MemberID ?? "", out string AgentCodeLevelName);
				AgentCodeLevelName = String.IsNullOrEmpty(AgentCodeLevelName) ? "" : AgentCodeLevelName.Trim();
				namelist.Add((Member.Get(To[i].MemberID) == null ? _Service.GetAgentName(To[i].MemberID) : Member.Get(To[i].MemberID).Name) + " " + AgentCodeLevelName);
			}
			reportModel.To = namelist;

			return View(reportModel);
		}

		/// <summary>
		/// 催辦
		/// </summary>
		/// <param name="CaseGuid"></param>
		/// <returns></returns>
		[PdLogFilter("EIP客服-通知作業-催辦", PITraceType.Download)]
		public ActionResult DoSNote(string CaseGuid)
		{
			var content = _Service.QueryNotifyDataByGuid(CaseGuid);
			NotifyReportModel reportModel = new NotifyReportModel();
			CultureInfo taiwanCulture = new CultureInfo("zh-TW");
			reportModel.CreateTime = content.CreateTime.ToString("yyyy年MM月dd日");
			reportModel.DoUserName = !String.IsNullOrEmpty(content.DoUser) ? Member.Get(content.DoUser).Name : "";
			reportModel.Content = content.Content;
			reportModel.NotifyInsuranceViews = _Service.GetCRMEInsurancePolicy(content.No);
			reportModel.No = content.No;
			reportModel.DoUserTelExt = content.DoUserTelExt;			

			var To = _Service.GetToByNo(content.No);
			var CC = _Service.GetCCByNo(content.No);
			var OM = _Service.GetOMByNo(content.No);
			Dictionary<string, string> LevelName = _Service.GetLevelNameChsByAgentCode();
			Dictionary<string, string> WcCenterName = _Service.GetWcCenterNameByAgentCode();
			List<string> namelist = new List<string>();
			List<string> CCnamelist = new List<string>();
			List<string> OMnamelist = new List<string>();

			for (int i = 0; i < To.Count; i++)
			{
				WcCenterName.TryGetValue(To[i].MemberID ?? "", out string AgentCodeWcCenterName);
				LevelName.TryGetValue(To[i].MemberID ?? "", out string AgentCodeLevelName);
				AgentCodeWcCenterName = String.IsNullOrEmpty(AgentCodeWcCenterName) ? "" : AgentCodeWcCenterName;
				AgentCodeLevelName = String.IsNullOrEmpty(AgentCodeLevelName) ? "" : AgentCodeLevelName.Trim();
				namelist.Add(AgentCodeWcCenterName + " " + (Member.Get(To[i].MemberID) == null ? _Service.GetAgentName(To[i].MemberID) : Member.Get(To[i].MemberID).Name) + " " + AgentCodeLevelName);
			}
			reportModel.To = namelist;

			for (int i = 0; i < CC.Count; i++)
			{
				WcCenterName.TryGetValue(CC[i].MemberID ?? "", out string CCAgentCodeWcCenterName);
				LevelName.TryGetValue(CC[i].MemberID ?? "", out string CCAgentCodeLevelName);
				CCAgentCodeWcCenterName = String.IsNullOrEmpty(CCAgentCodeWcCenterName) ? "" : CCAgentCodeWcCenterName;
				CCAgentCodeLevelName = String.IsNullOrEmpty(CCAgentCodeLevelName) ? "" : CCAgentCodeLevelName.Trim();
				CCnamelist.Add(CCAgentCodeWcCenterName + " " + (Member.Get(CC[i].MemberID) == null ? _Service.GetAgentName(CC[i].MemberID) : Member.Get(CC[i].MemberID).Name) + " " + CCAgentCodeLevelName);
			}
			reportModel.CC = CCnamelist;

			for (int i = 0; i < OM.Count; i++)
			{
				WcCenterName.TryGetValue(OM[i].MemberID ?? "", out string OMAgentCodeWcCenterName);
				LevelName.TryGetValue(OM[i].MemberID ?? "", out string OMAgentCodeLevelName);
				OMAgentCodeWcCenterName = String.IsNullOrEmpty(OMAgentCodeWcCenterName) ? "" : OMAgentCodeWcCenterName;
				OMAgentCodeLevelName = String.IsNullOrEmpty(OMAgentCodeLevelName) ? "" : OMAgentCodeLevelName.Trim();
				OMnamelist.Add(OMAgentCodeWcCenterName + " " + (Member.Get(OM[i].MemberID) == null ? _Service.GetAgentName(OM[i].MemberID) : Member.Get(OM[i].MemberID).Name) + " " + "處代理人");
			}
			reportModel.OM = OMnamelist;

			var Dos = _Service.QueryCRMEDoSByNo(content.No);
			var listDos = _Service.QueryListCRMEDoSByNo(content.No);
			var DoSCount = _Service.QueryCRMEDoSCountByNo(content.No);
			if(Dos != null)
			{
				reportModel.DoSContent = Dos.Content;
				reportModel.DoSAuditCount = DoSCount;
				//reportModel.DoSAuditTime = Dos.CreateTime.ToString("yyyy年MM月dd日");
				reportModel.DoSAuditAddTime = Dos.CreateTime.AddDays(3).ToString("yyyy年MM月dd日");
				List<string> listContents = new List<string>();
				foreach (var crmeDoS in listDos)
				{
					string Content = crmeDoS.CreateTime.ToString("yyyy年MM月dd日") + ":" + crmeDoS.Content;
					listContents.Add(Content);
				}
				reportModel.listContent = listContents;
			}

			return View(reportModel);
		}

		/// <summary>
		/// 稽催
		/// </summary>
		/// <param name="CaseGuid"></param>
		/// <returns></returns>
		[PdLogFilter("EIP客服-通知作業-稽催", PITraceType.Download)]
		public ActionResult AuditNote(string CaseGuid)
		{
			var content = _Service.QueryNotifyDataByGuid(CaseGuid);
			NotifyReportModel reportModel = new NotifyReportModel();
			CultureInfo taiwanCulture = new CultureInfo("zh-TW");
			reportModel.CreateTime = content.CreateTime.ToString("yyyy年MM月dd日");
			reportModel.DoUserName = !String.IsNullOrEmpty(content.DoUser) ? Member.Get(content.DoUser).Name : "";
			reportModel.Content = content.Content;
			reportModel.NotifyInsuranceViews = _Service.GetCRMEInsurancePolicy(content.No);
			reportModel.No = content.No;
			reportModel.DoUserTelExt = content.DoUserTelExt;

			var To = _Service.GetToByNo(content.No);
			var CC = _Service.GetCCByNo(content.No);
			var OM = _Service.GetOMByNo(content.No);
			Dictionary<string, string> LevelName = _Service.GetLevelNameChsByAgentCode();
			Dictionary<string, string> WcCenterName = _Service.GetWcCenterNameByAgentCode();
			List<string> namelist = new List<string>();
			List<string> CCnamelist = new List<string>();
			List<string> OMnamelist = new List<string>();

			for (int i = 0; i < To.Count; i++)
			{
				WcCenterName.TryGetValue(To[i].MemberID ?? "", out string AgentCodeWcCenterName);
				LevelName.TryGetValue(To[i].MemberID ?? "", out string AgentCodeLevelName);
				AgentCodeWcCenterName = String.IsNullOrEmpty(AgentCodeWcCenterName) ? "" : AgentCodeWcCenterName;
				AgentCodeLevelName = String.IsNullOrEmpty(AgentCodeLevelName) ? "" : AgentCodeLevelName.Trim();
				namelist.Add(AgentCodeWcCenterName + " " + (Member.Get(To[i].MemberID) == null ? _Service.GetAgentName(To[i].MemberID) : Member.Get(To[i].MemberID).Name) + " " + AgentCodeLevelName);
			}
			reportModel.To = namelist;

			for (int i = 0; i < CC.Count; i++)
			{
				WcCenterName.TryGetValue(CC[i].MemberID ?? "", out string CCAgentCodeWcCenterName);
				LevelName.TryGetValue(CC[i].MemberID ?? "", out string CCAgentCodeLevelName);
				CCAgentCodeWcCenterName = String.IsNullOrEmpty(CCAgentCodeWcCenterName) ? "" : CCAgentCodeWcCenterName;
				CCAgentCodeLevelName = String.IsNullOrEmpty(CCAgentCodeLevelName) ? "" : CCAgentCodeLevelName.Trim();
				CCnamelist.Add(CCAgentCodeWcCenterName + " " + (Member.Get(CC[i].MemberID) == null ? _Service.GetAgentName(CC[i].MemberID) : Member.Get(CC[i].MemberID).Name) + " " + CCAgentCodeLevelName);
			}
			reportModel.CC = CCnamelist;

			for (int i = 0; i < OM.Count; i++)
			{
				WcCenterName.TryGetValue(OM[i].MemberID ?? "", out string OMAgentCodeWcCenterName);
				LevelName.TryGetValue(OM[i].MemberID ?? "", out string OMAgentCodeLevelName);
				OMAgentCodeWcCenterName = String.IsNullOrEmpty(OMAgentCodeWcCenterName) ? "" : OMAgentCodeWcCenterName;
				OMAgentCodeLevelName = String.IsNullOrEmpty(OMAgentCodeLevelName) ? "" : OMAgentCodeLevelName.Trim();
				OMnamelist.Add(OMAgentCodeWcCenterName + " " + (Member.Get(OM[i].MemberID) == null ? _Service.GetAgentName(OM[i].MemberID) : Member.Get(OM[i].MemberID).Name) + " " + "處代理人");
			}
			reportModel.OM = OMnamelist;

			var Dos = _Service.QueryCRMEDoSByNo(content.No);
			var Audit = _Service.QueryCRMEAuditByNo(content.No);
			var AuditCount = _Service.QueryCRMEAuditCountByNo(content.No);

			reportModel.DoSContent = Dos != null ? Dos.Content : "";
			reportModel.DoSAuditCount = AuditCount;
			if(Audit != null)
			{
				reportModel.DoSAuditTime = Audit.CreateTime.ToString("yyyy年MM月dd日");
				reportModel.DoSAuditAddTime = Audit.CreateTime.AddDays(2).ToString("yyyy年MM月dd日");
			}

			return View(reportModel);
		}
		#endregion

		#region 人員處理
		/// <summary>
		/// 組成留言接收者清單 
		/// </summary>
		/// <param name="idNamelist">接收者資料</param>
		/// <param name="partdata">共用參數</param>
		/// <returns></returns>
		private List<MessageTo> GetMsgIdsToMsgTList(List<UserSimpleInfo> idNamelist, MessageTo partdata)
		{
			List<MessageTo> list = new List<MessageTo>();

			for (int i = 0; i < idNamelist.Count; i++)
			{
				MessageTo kt = new MessageTo();
				//kt.mo_mgid =
				//收件Id
				kt.MSGOBJID = idNamelist[i].MemberId;// Convert.ToInt32(idAry[i]);
				kt.MSGOBJDate = partdata.MSGOBJDate;
				kt.MSGOBJReaderIP = partdata.MSGOBJReaderIP;
				kt.MSGOBJSendIP = partdata.MSGOBJSendIP;
				kt.MSGOBJCreateTime = partdata.MSGOBJCreateTime;
				kt.MSGOBJSendID = partdata.MSGOBJSendID;
				//收件名稱
				//kt.MSGOBJName = idNamelist[i].Name  ;
				kt.MSGOBJName = idNamelist[i].Unit + "-" + idNamelist[i].Name;
				list.Add(kt);

			}
			return list;
		}
		#endregion

		#region 下載檔案
		/// <summary>
		/// 取得檔案名稱
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		public JsonResult GetFileName(string ID)
		{
			var LS = _Service.GetCRMEFileLstByMfid(ID);
			return Json(LS, JsonRequestBehavior.AllowGet);
		}

		/// <summary>
		/// 下載檔案
		/// </summary>
		/// <param name="fid"></param>
		/// <returns></returns>
		[PdLogFilter("EIP客服-通知作業-下載檔案", PITraceType.Download)]
		public ActionResult GetCRMEFile(int fid)
		{
			RemoteFileInfo rf = null;
			var mSevice = ServiceHelper.Create<INotifyService>();
			CRMEFile cf = mSevice.GetCRMEFileByMfid(fid);
			var streamservice = ServiceHelper.Create<IStreamMediaService>();
			DownloadRequest dr = new DownloadRequest();
			dr.FileName = cf.FileMD5Name;
			//var mimeType = "application/octet-stream";
			try
			{
				rf = streamservice.DownloadCUSCRMFile(dr);
			}
			catch
			{
				//return Content(ee.Message );
				return Content("Get File fail");
				//return Content("<script language='javascript' type='text/javascript'>alert('download file fail!');</script>"); 
			}
			return File(rf.FileByteStream, rf.MimeType, cf.FileName);
		}
		#endregion
	}
}