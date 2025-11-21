using EP.Platform.Service;
using Microsoft.CUF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    /// <summary>
    /// 立案相關服務
    /// </summary>
    [ServiceContract]
    public interface ICaseService
    {
        /// <summary>
        /// 依要保人ID取得保單資料
        /// </summary>
        /// <param name="ownerID">要保人ID</param>
        /// <returns>要保人底下的所有產、壽保單資料清單</returns>
        [OperationContract]
        IEnumerable<CRMEInsurancePolicy> GetInsPolicyByOwnerID(string ownerID);

        /// <summary>
        /// 檢核保單號碼的狀態
        /// </summary>
        /// <param name="policyNo">保單號碼</param>
        /// <returns>
        /// 1. false:接續人為公司 true:有相同的保單號碼或是查無保單
        /// 2. 保單號碼的資料，如果1為true時，且沒有資料，代表查無保單
        /// </returns>
        [OperationContract]
        Tuple<bool, IEnumerable<CRMEInsurancePolicy>> CheckPolicyNo(string policyNo);

        /// <summary>
        /// 依服務申訴類型取得資料來源
        /// </summary>
        /// <param name="type">服務申訴類型</param>
        /// <returns>資料來源清單</returns>
        [OperationContract]
        IEnumerable<CRMEDiscipType> GetSourceListByType(DiscipTypeCode type);

        /// <summary>
        /// 依服務申訴類型取得來電者
        /// </summary>
        /// <param name="type">服務申訴類型</param>
        /// <returns>來電者清單</returns>
        [OperationContract]
        IEnumerable<CRMEDiscipType> GetCallerListByType(DiscipTypeCode type);

        /// <summary>
        /// 依服務申訴類型取得案件類別
        /// </summary>
        /// <param name="type">服務申訴類型</param>
        /// <returns>案件類別清單</returns>
        [OperationContract]
        IEnumerable<CRMEDiscipType> GetCaseTypeListByType(DiscipTypeCode type);

        /// <summary>
        /// 依服務申訴類型取得案件類型
        /// </summary>
        /// <param name="type">服務申訴類型</param>
        /// <returns>案件類型清單</returns>
        [OperationContract]
        IEnumerable<CRMEDiscipType> GetCaseCategoryListByType(DiscipTypeCode type);

        /// <summary>
        /// 取得保險司清單(先壽後產)
        /// </summary>
        /// <returns>保險公司清單</returns>
        [OperationContract]
        IEnumerable<ValueText> GetInsCompanyList();

        /// <summary>
        /// 取得要保人手號碼
        /// </summary>
        /// <param name="ownerID">要保人ID</param>
        /// <returns>要保人手機號碼</returns>
        [OperationContract]
        string GetOwnerMobile(string ownerID);

        /// <summary>
        /// 新增立案
        /// </summary>
        /// <param name="content">立案的資料</param>
        /// <param name="policys">保單資料清單</param>
        /// <returns>立案後的受理編號</returns>
        [OperationContract]
        string CreateCaseContent(CRMECaseContent content, List<CRMEInsurancePolicy> policys);

        /// <summary>
        /// 取得等待通知的案件資料清單
        /// </summary>
        /// <returns>等待通知的案件資料清單</returns>
        [OperationContract]
        IEnumerable<CRMECaseContent> GetWaitNofityDatas();

        /// <summary>
        /// 取得等待通知申訴人的案件資料清單
        /// </summary>
        /// <returns>等待通知申訴人的案件資料清單</returns>
        [OperationContract]
        IEnumerable<CRMECaseContent> GetCCWaitNofityDatas();

        /// <summary>
        /// 檢核是否有未結案的保單號碼
        /// </summary>
        /// <param name="policyNo">保單號碼</param>
        /// <returns>true:有未結案 false:都結案</returns>
        [OperationContract]
        bool CheckNotClosedPolicyNo(string policyNo);

        /// <summary>
        /// 取得業務員資訊
        /// </summary>
        /// <param name="agentCode">業務員代碼</param>
        /// <returns>業務員資訊</returns>
        [OperationContract]
        ExtraData GetAgentInfo(string agentCode);
    }
}
