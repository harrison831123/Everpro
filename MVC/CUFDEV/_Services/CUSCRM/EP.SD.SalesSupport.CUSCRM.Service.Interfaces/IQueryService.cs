using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    /// <summary>
    /// 客服業務系統的查詢服務
    /// </summary>
    [ServiceContract]
    public interface IQueryService 
    {

        #region h2o歷史Service
        /// <summary>
        /// 查歷史客服業務列表
        /// </summary>
        /// <returns></returns>
        [OperationContract]

        IEnumerable<HistoryCSViewModel> QueryHistoryCSGridDatas(HistoryQueryCondition condition);
        /// <summary>
        /// 查歷史聯繫紀錄列表
        /// </summary>
        /// <param name="crm_no">受理編號</param>
        /// <returns>歷史聯繫紀錄列表</returns>
        [OperationContract]

        IEnumerable<tcrm_do> QueryHistoryMaintainRecordList(string crm_no);


        /// <summary>
        /// 查歷史聯繫紀錄檔案列表
        /// </summary>
        /// <param name="crm_no">受理編號</param>
        /// <returns>歷史聯繫紀錄檔案列表</returns>
        [OperationContract]
        IEnumerable<crm_do_file> QueryHistoryMaintainFileList(string crm_no);

        /// <summary>
        /// 新增聯繫紀錄
        /// </summary>
        /// <param name="model">聯繫紀錄</param>
        [OperationContract]

        void CreateCrm_do(tcrm_do model);
        /// <summary>
        /// 新增上傳檔案
        /// </summary>
        /// <param name="model"></param>

        [OperationContract]
        void CreateCrm_do_file(crm_do_file model);
        /// <summary>
        /// 新增結案紀錄
        /// </summary>
        /// <param name="model"></param>

        [OperationContract]
        void CreateCloseRecord(crm_close_log model);

        #endregion

        #region 報表查詢Service
        [OperationContract]
        Stream GetReport(QueryReportCondition condition);


        #endregion
    }
}
