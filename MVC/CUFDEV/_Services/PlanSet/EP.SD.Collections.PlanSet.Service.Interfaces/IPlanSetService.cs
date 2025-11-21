using System;
using EP.Platform.Service;
using EP.SD.Collections.PlanSet.Models;
using System.Collections.Generic;
using System.ServiceModel;
using Microsoft.CUF;
using EP.VBEPModels;
using System.IO;

namespace EP.SD.Collections.PlanSet.Service
{
    /// <summary>
    /// 需求單號：20241210003 現售商品佣獎查詢 2024.12 BY VITA
    /// 需求單號：20250707001 競賽計C商品清單。 2025/07/03 BY Harrison 
    /// 202505 by Fion 20250527001_佣酬預估試算
    /// </summary>
    [ServiceContract]
	public interface IPlanSetService
	{
        /// <summary>
        /// 取得保公
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        List<ValueText> GetCompanyCode();

        /// <summary>
        /// 查詢初年度業績換算率(業行部)、保公獎勵內容(業行部)、永達競賽獎勵(業支部)
        /// </summary>
        /// <param name = "model" ></ param >
        /// < returns ></ returns >
        [OperationContract]
        PlanSetAll GetQueryArea(PlanSetMainInput model);

        /// <summary>
        /// 需求單號：202503xx00X 險種中文名稱查詢 by vita 2025.03
        /// </summary>
        /// <param name="company_code"></param>
        /// <param name="plan_title"></param>
        /// <param name="chktype"></param>
        /// <returns></returns>
        [OperationContract]
        List<PlanTitle> GetPlanTitle(string company_code, string plan_title, string chktype);

        #region 佣酬預估試算 202505 by Fion 20250527001_佣酬預估試算
        /// <summary>
        /// 取得業務員基本資料
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        [OperationContract]
        ExtraData GetAgentData(AgentRewardPolicyCondition condition);

        /// <summary>
        /// 業務員最新MDRT屆數
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        [OperationContract]
        string GetAgentRewardMDRT(AgentRewardPolicyCondition condition);

        /// <summary>
        /// 個人達成報酬率
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        [OperationContract]
        IEnumerable<AgentRewardRange> GetAgentRewardRange(AgentRewardPolicyCondition condition);

        /// <summary>
        /// 未核保保單資訊
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        [OperationContract]
        IEnumerable<AgentRewardPolicyInfo> QueryUnPaidRewardPolicyData(AgentRewardPolicyCondition condition);
        
        /// <summary>
        /// 202510 by Fion 20250901003_佣酬預估試算優化
        /// </summary>
        /// <param name="agentCode"></param>
        /// <returns></returns>
        [OperationContract]
        ExtraData QueryAgentRewardTotIncome(string agentCode);
        #endregion 佣酬預估試算 


        #region 競賽計C商品清單
        /// <summary>
        /// 取得報表工作日
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        string GetWorkDate();

        /// <summary>
        /// 取得商品清單資料
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>

        [OperationContract]
        List<PlanSetWarptSet> GetPlanSetWarptSet(PlanSetWarptSetCondition condition);

        /// <summary>
        /// 取得報表
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        [OperationContract]
        Stream GetPlanSetWarptSetReport(PlanSetWarptSetCondition condition);
        #endregion
        
    }
}