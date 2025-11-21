using System.Collections.Generic;
using System.ServiceModel;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    /// <summary>
    /// 客服業務系統的共用服務
    /// </summary>
    [ServiceContract]
    public interface ICommonService
    {
        /// <summary>
        /// 新增資料設定
        /// </summary>
        /// <param name="model">要新增的資料</param>
        /// <returns>新增後的資料</returns>
        [OperationContract]
        CRMEDiscipType CreateCRMEDiscipType(CRMEDiscipType model);

        /// <summary>
        /// 修改資料設定
        /// </summary>
        /// <param name="model">要修改的資料</param>
        /// <returns>修改後的資料</returns>
        [OperationContract]
        CRMEDiscipType UpdateCRMEDiscipType(CRMEDiscipType model);

        /// <summary>
        /// 刪除資料設定
        /// </summary>
        /// <param name="id">要刪除的資料的自動編號</param>
        [OperationContract]
        void DeleteCRMEDiscipType(int id);

        /// <summary>
        /// 查詢資料設定的資料清單
        /// </summary>
        /// <param name="condition">查詢條件</param>
        /// <returns>資料設定的資料清單</returns>
        [OperationContract]
        IEnumerable<CRMEDiscipType> QueryCRMEDiscipTypeDatas(QueryDiscipTypeCondition condition);

        /// <summary>
        /// 取得類型對應的案件的類別
        /// </summary>
        /// <param name="category">類型</param>
        /// <returns>類別清單</returns>
        [OperationContract]
        IEnumerable<CRMEDiscipType> GetCaseType(Category? category);

        /// <summary>
        /// 依自動編號取得對應的資料設定
        /// </summary>
        /// <param name="id">自動編號</param>
        /// <returns>資料設定</returns>
        [OperationContract]
        CRMEDiscipType GetCRMEDiscipTypeByID(int id);

        /// <summary>
        /// 依受理編號取得保單對應的通知對像id
        /// </summary>
        /// <param name="no">受理編號</param>
        /// <returns>保單對應的通知對像id清單</returns>
        [OperationContract]
        IEnumerable<string> GetDefaultNotifyMemberIdByCRMENo(string no);

    }
}
