using EP.SD.SalesZone.AGUPG.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesZone.AGUPG.Service
{
    [ServiceContract]
    public interface IAGUPGService
    {
        /// <summary>
        /// 取得查詢資料
        /// </summary>
        /// <param name="AgentCode"></param>
        /// <returns></returns>
        [OperationContract]
        HrUpg25Dto GetQueryHrUpg25(HrUpg25QueryCondition model);

        /// <summary>
        /// 取得下拉選單
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        List<HrUpg25RstViewModel.HrUpg25Item> GetHrUpg25Rst();

        /// <summary>
        /// 取得明細資料
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [OperationContract]
        HrUpgGet25Dto GetQueryHrUpgGet25WebShowDetail(HrUpg25QueryCondition model);

        /// <summary>
        /// 取得職等
        /// </summary>
        /// <param name="AgentCode"></param>
        /// <returns></returns>
        [OperationContract]
        string GetAgLevel(string AgentCode);

        /// <summary>
        /// 依照身分別，取得畫面明細資料
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [OperationContract]
        FamilyDto GetHrUpg25RstFamilyTree(HrUpg25QueryCondition model);
    }
}
