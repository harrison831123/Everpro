///========================================================================================== 
/// 程式名稱：人工調帳
/// 建立人員：Harrison
/// 建立日期：2022/07
/// 修改記錄：（[需求單號]、[修改內容]、日期、人員）
/// 需求單號:20240122004-因現有VLIFE系統(核心系統)使用已長達20多年，架構老舊，已不敷使用，且為提升資訊安全等級，故計劃執行VLIFE系統改版(新核心系統:eBroker系統)。; 修改內容:上線; 修改日期:20240613; 修改人員:Harrison;
/// 需求單號:20240807001-調整人工調帳系統產出之相關畫面及報表修改等功能。; 修改日期:20240807; 修改人員:Harrison;
///==========================================================================================
using EB.CUFModels;
using EB.EBrokerModels;
using EB.SL.PayRoll.Models;
using EB.SL.PayRoll.Service.Contracts;
using EB.VLifeModels;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.PayRoll.Service
{
	[ServiceContract]
	public interface IPayRollService
	{
		#region 單筆調整
		/// <summary>
		/// 查詢一筆OpCalendar時間
		/// </summary>
		[OperationContract]
		OpCalendar QueryOpCalendar(OpCalendar model);

		/// <summary>
		/// 輸入原因碼取得對應的型態代碼
		/// </summary>
		/// <param name="reasonCode">原因碼</param>
		/// <returns></returns>
		[OperationContract]
		string GetReasonCodeToAdjTypeMappingValue(string reasonCode);

		/// <summary>
		/// 查詢調整紀錄
		/// </summary>
		/// <param name="QueryAgentBonusCondition">查詢調整條件</param>
		[OperationContract]
		IEnumerable<Tuple<AgentBonusAdjust, AgentBonusDesc>> QueryAgentBonus(QueryAgentBonusCondition condition);

		/// <summary>
		/// 新增單筆調整
		/// </summary>
		/// <param name="adjustModel">業務佣酬調整資料檔</param>
		/// <param name="descModel">業務佣酬說明檔</param>
		[OperationContract]
		List<AgentBonusAdjust> InsertSingleAgentBonusAdjust(AgentBonusAdjust adjustModel, AgentBonusDesc descModel);

		/// <summary>
		/// 用業務酬佣說明檔GUID刪除單筆調整某一筆資料
		/// </summary>
		/// <param name="agentBonusDescGuid">業務酬佣說明檔GUID</param>
		/// <param name="UserID">UserID</param>
		/// <returns></returns>
		[OperationContract]
		bool DeleteAgentBonusAdjustByAgentBonusDescGuid(string agentBonusDescGuid, string UserID);

		/// <summary>
		/// 更新紀錄
		/// </summary>
		/// <param name="adjustModel">業務佣酬調整資料檔</param>
		/// <param name="descModel">業務佣酬說明檔</param>
		[OperationContract]
		bool UpdateAgentBonusAdjust(AgentBonusAdjust adjustModel, AgentBonusDesc descModel);

		/// <summary>
		/// 取單筆資料by iden
		/// </summary>
		/// <param name="iden"></param>
		/// <returns></returns>
		[OperationContract]
		Tuple<AgentBonusAdjust, AgentBonusDesc> QueryAgentBonusByIden(int iden);

		/// <summary>
		/// 取得所有實駐
		/// </summary>
		/// <returns></returns>
		//[OperationContract]
		//List<AGWCSet> GetAGWCSet();

		/// <summary>
		/// 用保險公司分公司代碼取得保險公司代碼
		/// </summary>
		/// <param name="subCode">保險公司分公司代碼</param>
		/// <returns></returns>
		[OperationContract]
		string GetCompanyCodeBySubCode(string subCode);

		/// <summary>
		/// 輸入業務員代碼，從aginb串nain取得業務員姓名
		/// </summary>
		/// <param name="productionYM">業績年月</param>
		/// <param name="agentCode">業務員代碼</param>
		/// <returns></returns>
		[OperationContract]
		string GetAgentNameByAginb(string productionYM, string agentCode);

		/// <summary>
		/// 查詢任職型態3(內勤)人員，FYC需為0
		/// </summary>
		/// <param name="subCode">保險公司分公司代碼</param>
		/// <returns></returns>
		[OperationContract]
		string GetOrgibByagentcode(string agentcode, string ProductionYM, string Sequence);

		/// <summary>
		/// 查詢任職型態3(內勤)人員
		/// </summary>
		/// <param name="subCode">保險公司分公司代碼</param>
		/// <returns></returns>
		[OperationContract]
		List<string> GetOrgibByagentcodeList(string ProductionYM, string Sequence);

		/// <summary>
		/// 查詢簽約日期(報聘日)
		/// </summary>
		/// <param name="subCode">保險公司分公司代碼</param>
		/// <returns></returns>
		[OperationContract]
		string GetAginbbByAgentcode(string agentcode);

		/// <summary>
		/// 確認業務員是否在黑名單
		/// </summary>
		/// <param name="subCode">保險公司分公司代碼</param>
		/// <returns></returns>
		[OperationContract]
		string GetBonusReserveSetting(string ProductionYM, string Sequence, string agentcode, string ReasonCode);

		/// <summary>
		/// 取得黑名單
		/// </summary>
		/// <param name="ProductionYM"></param>
		/// <param name="Sequence"></param>
		/// <param name="agentcode"></param>
		/// <param name="ReasonCode"></param>
		/// <returns></returns>
		[OperationContract]
		List<BlackList> GetBonusReserveSettingList(string ProductionYM, string Sequence);

		/// <summary>
		/// 查詢是否為免評估人員
		/// </summary>
		/// <param name="agentcode"></param>
		/// <returns></returns>
		[OperationContract]
		string GetaginbByagentcode(string agentcode, string ProductionYM, string Sequence);

		/// <summary>
		/// 查詢所有免評估人員
		/// </summary>
		/// <param name="agentcode"></param>
		/// <returns></returns>
		[OperationContract]
		List<string> GetaginbByagentcodeList(string ProductionYM, string Sequence);

		/// <summary>
		/// 檢核對應的原因碼，繳别繳次必須要有內容
		/// </summary>
		/// <param name="agentcode"></param>
		/// <returns></returns>
		[OperationContract]
		string GetGroupMappingTOmodxseq(string reasoncode);

		/// <summary>
		/// 取得繳别繳次必須要有內容的原因碼
		/// </summary>
		/// <param name="reasoncode"></param>
		/// <returns></returns>
		[OperationContract]
		List<string> GetGroupMappingTOmodxseqList();

		/// <summary>
		/// 檢核對應的原因碼，此原因碼不可人工調整
		/// </summary>
		/// <param name="agentcode"></param>
		/// <returns></returns>
		[OperationContract]
		string GetGroupMappingTOnotin(string reasoncode);

		/// <summary>
		/// 取得R2、N4及NB原因碼不可人工調整
		/// </summary>
		/// <param name="reasoncode"></param>
		/// <returns></returns>
		[OperationContract]
		List<string> GetGroupMappingTOnotinList();

		/// <summary>
		/// 檢核對應的原因碼，此原因碼金額應為正數
		/// </summary>
		/// <param name="reasoncode"></param>
		/// <returns></returns>
		[OperationContract]
		string GetGroupMappingTOplus(string reasoncode);

		/// <summary>
		/// 取得金額應為正數的原因碼
		/// </summary>
		/// <param name="reasoncode"></param>
		/// <returns></returns>
		[OperationContract]
		List<string> GetGroupMappingTOplusList();

		/// <summary>
		/// 檢核對應的原因碼，此原因碼金額應為負數
		/// </summary>
		/// <param name="agentcode"></param>
		/// <returns></returns>
		[OperationContract]
		string GetGroupMappingTOminus(string reasoncode);

		/// <summary>
		/// 取得金額應為負數的原因碼
		/// </summary>
		/// <param name="reasoncode"></param>
		/// <returns></returns>
		[OperationContract]
		List<string> GetGroupMappingTOminusList();


		/// <summary>
		/// 取得初年度服務報酬檢核的原因碼
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<string> GetGroupMappingTOfirstyearList();

		/// <summary>
		/// 檢核對應的原因碼，此原因碼為初年度服務報酬
		/// </summary>
		/// <param name="reasoncode"></param>
		/// <returns></returns>
		[OperationContract]
		string GetGroupMappingTOZeroFour(string reasoncode);

		/// <summary>
		/// 檢核對應的原因碼，查詢型態為6的原因碼
		/// </summary>
		/// <param name="reasoncode"></param>
		/// <returns></returns>
		[OperationContract]
		string GetGroupMappingTOsix(string reasoncode);

		/// <summary>
		/// 查詢型態為6的原因碼的原因碼
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<string> GetGroupMappingTOsixList();

		/// <summary>
		/// 從poag取得所有保單
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<Poag> GetAgentNameByPoagList();

		/// <summary>
		/// 查詢所有業務員、簽約日期(報聘日)
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		Dictionary<string, string> GetOrgin();

		/// <summary>
		/// 檢核FYC僅原因碼01~04不可空白，其餘原因碼系統預設為0
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<string> GetGroupMappingTOFYC();

		/// <summary>
		/// 檢核FYP除原因碼01~07、0A~0G不得空白，其餘原因碼系動自動帶0
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<string> GetGroupMappingTOFYP();

		/// <summary>
		/// 檢核年期原因碼01~07、0A~0B+0G，其餘型態預設為0，但查詢畫面及報表產出為空白
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<string> GetGroupMappingTOYear();

		/// <summary>
		/// 檢核險種代碼僅原因碼01~03、05~06及0A~0B須有險種代號
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<string> GetGroupMappingTOPlanCode();

		/// <summary>
		/// 檢核金額>保費，原因碼01~07、0A~0G
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<string> GetGroupMappingTOAF();
		#endregion

		#region 大批調整
		/// <summary>
		/// 大批調整刪除該user該工作月資料
		/// </summary>
		/// <param name="productionYM"></param>
		/// <param name="sequence"></param>
		/// <param name="accountID"></param>
		/// <returns></returns>
		[OperationContract]
		void DeleteAdjustByAccountID(string productionYM, string sequence, string accountID, string ProcessCode);

		/// <summary>
		/// 大批調整bulkinsert
		/// </summary>
		/// <param name="dt">data</param>
		[OperationContract]
		void ag078BulkInsert(DataTable dt, List<AgentBonusDesc> descs);

		/// <summary>
		/// 取得所有原因碼與型態對應表
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<GroupMapping> GetALLGroupMappingList();

		/// <summary>
		/// 新增上傳原因碼log
		/// </summary>
		/// <param name="log">SampleFileUploadLog model</param>
		[OperationContract]
		void CreateSampleFileUploadLog(SampleFileUploadLog log);

		/// <summary>
		/// 從AginB取得指定工作月的所有業務員ID與姓名
		/// </summary>
		/// <param name="productionYM">工作月</param>
		/// <returns></returns>
		[OperationContract]
		Dictionary<string, string> GetAginbAgentCodeAndAgentNameByproductionYM(string productionYM);

		/// <summary>
		/// 取得保險公司代碼清單
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<TermVal> GetCompanySetList();

		/// <summary>
		/// 取得該工作月離職人員ID清單
		/// </summary>
		/// <param name="productionYM">業績年月</param>
		/// <returns></returns>
		[OperationContract]
		List<string> GetAginBQuitAG(string productionYM);

		/// <summary>
		/// 從poag確認有無該業務員
		/// </summary>
		/// <param name="PolicyNo2">保單號碼</param>
		/// <returns></returns>
		[OperationContract]
		string CheckAgentNameByPoag(string PolicyNo2);
		#endregion

		#region AG078報表
		/// <summary>
		/// 取得所有總公司所有部門
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<sc_unit> Getnunit();

		/// <summary>
		/// 依照年分撈取業績年月
		/// </summary>
		/// <param name="sql"></param>
		[OperationContract]
		List<string> GetAllYMData(string ym);

		/// <summary>
		/// 撈取 業務佣酬調整資料檔
		/// </summary>
		/// <param name="sql"></param>
		[OperationContract]
		List<AgentBonusAdjustViewModel> GetAgentBonusAdjust(AgentBonusAdjustViewModel condition);

		/// <summary>
		/// 產出報表
		/// </summary>
		[OperationContract]
		Stream GetAG078ReportList(AgentBonusAdjustViewModel condition);
		#endregion

		#region GetProcessNo(取得處理號碼)

		/// <summary>
		/// 取得處理號碼
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		string GetProcessNo();

		/// <summary>
		/// 取得工作月
		/// </summary>
		/// <param name="strSel">query type</param>
		/// <param name="exclude88">是否排掉 sequence=88</param>
		/// <returns></returns>
		[OperationContract]
		List<agym> GetYMData(string strSel, bool exclude88, int selectTop = 1);

		/// <summary>
		/// 輸入保單號碼，取得業務員姓名
		/// </summary>
		/// <param name="PolicyNo2">保單號碼</param>
		/// <returns></returns>
		[OperationContract]
		List<string> GetAgentNameByPoag(string PolicyNo2);

		#endregion

		#region 原因碼總表
		/// <summary>
		/// 取得AgentBonusAdjust工作月
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<string> GetAgentBonusAdjustProductionYm(string ProductionY, string CreateUnit);

		/// <summary>
		/// 取得AgentBonusAdjust工作月(業行、財務)
		/// </summary>
		/// <returns></returns>
		[OperationContract]
		List<string> GetAgentBonusAdjustMainProductionYm(string ProductionY, string CreateUnit);

		/// <summary>
		/// 確認原因碼總表是否為空
		/// </summary>
		/// <param name="condition"></param>
		/// <returns></returns>
		[OperationContract]
		List<TotalReasonCodeReportModel> CheckTotalReasonCodeReport(AgentBonusAdjustCondition condition);

		/// <summary>
		/// 原因碼總表
		/// </summary>
		/// <param name="year">年度</param>
		/// <returns>Stream</returns>
		[OperationContract]
		Stream GetTotalReasonCodeReport(AgentBonusAdjustCondition condition);

		/// <summary>
		/// 取得部門level
		/// </summary>
		/// <param name="unitid">部門編號</param>
		[OperationContract]
		int GetUnitlevel(string unitid);
		#endregion
	}
}
