using System;
using System.Collections.Generic;
using System.ServiceModel;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    /// <summary>
    /// 客服業務系統的維護服務
    /// </summary>
    [ServiceContract]
    public interface ICUSCRMTX003Service
    {
        /// <summary>
        /// 查詢待維護清單
        /// </summary>
        /// <param name="condition">查詢條件</param>
        /// <returns></returns>
        [OperationContract]
        List<MaintainInfo> GetMtnData(QueryMaintainCondition condition);

        /// <summary>
        /// 執行稽催
        /// </summary>
        /// <param name="cRMEAudit">稽催model</param>
        /// <param name="ip">user id</param>
        /// <param name="userMemberName">user name</param>
        /// <param name="userMemberId">user memberId</param>
        /// <returns></returns>
        [OperationContract]
        bool DoAudit(CRMEAudit cRMEAudit, string ip, string userMemberName, string userMemberId);

        /// <summary>
        /// 執行催辦
        /// </summary>
        /// <param name="CRMEDoS">催辦model</param>
        /// <param name="ip">user id</param>
        /// <param name="userMemberName">user name</param>
        /// <param name="userMemberId">user memberId</param>
        /// <returns></returns>
        [OperationContract]
        bool DoS(CRMEDoS cRMEDoS, string ip, string userMemberName, string userMemberId);

        /// <summary>
        /// 產生處理單資料
        /// </summary>
        /// <param name="no">受理編號</param>
        /// <returns></returns>
        [OperationContract]
        ProcessFormData ProduceProcessFormData(string no);

        /// <summary>
        /// 根據類型取得歷史維護清單
        /// </summary>
        /// <param name="no">受理編號</param>
        /// <returns></returns>
        [OperationContract]
        MtnHistoryInfo GetDoHistory(string no);

        /// <summary>
        /// 將上傳的檔案存入DB
        /// </summary>
        /// <param name="cRMEFile">檔案model</param>
        [OperationContract]
        bool AddUploadFile(CRMEFile cRMEFile);

        /// <summary>
        /// 存入一筆維護紀錄
        /// </summary>
        /// <param name="cRMEDo">維護紀錄model</param>
        [OperationContract]
        bool InserCRMEDo(CRMEDo cRMEDo);

        /// <summary>
        /// 依檔案ID取檔案資訊
        /// </summary>
        /// <param name="id">檔案ID</param>
        /// <returns></returns>
        [OperationContract]
        CRMEFile GetCRMEFileById(int id);

        /// <summary>
        /// 依照檔案ID刪除實體檔案
        /// </summary>
        /// <param name="id">檔案ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteFileByCRMEFIleId(CRMEFile cRMEFile);

        /// <summary>
        /// 依檔案ID刪除實體檔案與檔案資訊
        /// </summary>
        /// <param name="list">檔案ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteCRMEFileDbAndFileById(List<int> list);

        /// <summary>
        /// 結案
        /// </summary>
        /// <param name="cRMECloseLog">結案model</param>
        /// <returns></returns>
        [OperationContract]
        bool InsertCloseLog(CRMECloseLog cRMECloseLog);
    }
}
