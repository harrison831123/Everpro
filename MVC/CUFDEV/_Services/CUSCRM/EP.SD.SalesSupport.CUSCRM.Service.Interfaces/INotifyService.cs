using EP.CUFModels;
using EP.H2OModels;
using EP.Platform.Service;
using EP.SD.SalesSupport.CUSCRM.Service.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
	[ServiceContract]
	public interface INotifyService 
	{
		/// <summary>
		/// 取得立案通知資料的受文者、副本受文者、受理保單資料
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		List<NotifyInsuranceViewModel> GetCRMEInsurancePolicy(string No);

		/// <summary>
		/// 處代理人
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		List<NotifyInsuranceViewModel> GetCRMEChinaOm(string No);

		/// <summary>
		/// 取得行專資料
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		List<NotifyCaseViewModel> GetCRMENotifyEmployee(string code);


		/// <summary>
		/// 取得主表資料
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		CRMECaseContent GetCRMECaseContent(string No);

		/// <summary>
		/// 取得客服部人員
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		List<sc_member> GetCustomerUnit();

		/// <summary>
		/// 取得實駐
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		Dictionary<string, string> GetWcCenterByAgentCode();

		/// <summary>
		/// 取得實駐名稱
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		Dictionary<string, string> GetWcCenterNameByAgentCode();

		/// <summary>
		/// 取得職稱
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		Dictionary<string, string> GetLevelNameChsByAgentCode();

		/// <summary>
		/// 新增立案通知對象
		/// </summary>
		/// <param name="modeltoList"></param>
		[OperationContract]
		void InsertCRMENotifyTo(List<CRMENotifyTo> modeltoList);

		/// <summary>
		/// 更新主表
		/// </summary>
		/// <param name="model"></param>
		/// <returns></returns>
		[OperationContract]
		bool UpdateCaseContentType(CRMECaseContent model);

		/// <summary>
		/// 新增照會附件
		/// </summary>
		/// <param name="model">主檔</param>
		/// <param name="modelfileList">檔案附件物件</param>
		/// <param name="tabUniqueId">前端傳來唯一碼</param>
		/// <param name="RetMsg">回傳訊息</param>
		[OperationContract]
		void CreateCRMNotifyFile(CRMECaseContent model, List<CRMEFile> modelfileList, string tabUniqueId, out string RetMsg);

		/// <summary>
		/// 發送推播
		/// </summary>
		/// <param name="model"></param>
		/// <param name="Account"></param>
		/// <returns></returns>
		[OperationContract]
		void SendLineNotify(Message model, List<UserSimpleInfo> Account);

		/// <summary>
		/// 取得進度
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		[OperationContract]
		NotifyViewModel GetCRMESeq(string No);

		/// <summary>
		/// 取得VLife人名
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		string GetAgentName(string agentcode);

		/// <summary>
		/// 取得人員ID
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		string GetAccountId(string MemberId);

		/// <summary>
		/// 查詢通知作業-客服
		/// </summary>
		/// <param name="condition"></param>
		/// <returns></returns>
		[OperationContract]
		List<NotifyViewModel> QueryNotifyDataByCRM(QueryNotifyCondition condition);

		/// <summary>
		/// 查詢通知作業-行專
		/// </summary>
		/// <param name="condition"></param>
		/// <returns></returns>
		[OperationContract]
		List<NotifyViewModel> QueryNotifyEPDataByCRM(QueryNotifyCondition condition);

		/// <summary>
		/// 查詢通知作業-業務員
		/// </summary>
		/// <param name="condition"></param>
		/// <returns></returns>
		[OperationContract]
		List<NotifyViewModel> QueryNotifyDataByAgent(QueryNotifyCondition condition);

		/// <summary>
		/// 取得客服業務檔案資料
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		[OperationContract]
		CRMEFile GetCRMEFileByMfid(int ID);

		/// <summary>
		/// 取得客服業務檔案資料
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		[OperationContract]
		List<CRMEFile> GetCRMEFileLstByMfid(string ID);

		/// <summary>
		/// 依Guid取得主表資料
		/// </summary>
		/// <param name="condition"></param>
		/// <returns></returns>
		[OperationContract]
		CRMECaseContent QueryNotifyDataByGuid(string CaseGuid);

		/// <summary>
		/// 依No取得受文者
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		List<CRMENotifyTo> QueryNotifyToByNo(string No);

		/// <summary>
		/// 依No取得受文者
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		List<CRMENotifyTo> GetToByNo(string No);

		/// <summary>
		/// 依No取得副本受文者
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		List<CRMENotifyTo> GetCCByNo(string No);

		/// <summary>
		/// 依No取得處代理人
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		List<CRMENotifyTo> GetOMByNo(string No);

		/// <summary>
		/// 依TypeID取得對應名稱
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		string GetDiscipTypeByID(int ID);

		/// <summary>
		/// 取得催辦內容
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		CRMEDoS QueryCRMEDoSByNo(string No);

		/// <summary>
		/// 取得催辦內容次數
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		int QueryCRMEDoSCountByNo(string No);

		/// <summary>
		/// 取得所有催辦內容
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		List<CRMEDoS> QueryListCRMEDoSByNo(string No);

		/// <summary>
		/// 取得稽催內容
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		CRMEAudit QueryCRMEAuditByNo(string No);

		/// <summary>
		/// 取得稽催內容次數
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		int QueryCRMEAuditCountByNo(string No);

		/// <summary>
		/// 取得催辦進度
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		[OperationContract]
		CRMEDoS GetCRMEDoS(string No);

		/// <summary>
		/// 取得稽催進度
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		[OperationContract]
		CRMEAudit GetCRMEAudit(string No);

		/// <summary>
		/// 取得催辦進度次數
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		[OperationContract]
		int GetCRMEDoSSep(string No);

		/// <summary>
		/// 取得稽催進度次數
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		[OperationContract]
		int GetCRMEAuditSep(string No);

		/// <summary>
		/// 新增申訴通知
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		[OperationContract]
		void InsertCRMEAppealBy(CRMEAppealBy model);




	}
}
