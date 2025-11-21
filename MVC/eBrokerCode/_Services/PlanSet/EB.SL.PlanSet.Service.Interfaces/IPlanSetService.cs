using EB.EBrokerModels;
using EB.SL.PlanSet.Models;
using EB.VLifeModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.PlanSet.Service
{
	[ServiceContract]
	public interface IPlanSetService
	{
        #region 調整起迄時間設定
        /// <summary>
        /// 取得工作月
        /// </summary>
        /// <param name="strSel">query type</param>
        /// <param name="exclude88">是否排掉 sequence=88</param>
        /// <returns></returns>
        [OperationContract]
        List<agym> GetYMData(string strSel, bool exclude88, int selectTop = 1);

        /// <summary>
        /// 查詢一筆OpCalendar時間
        /// </summary>
        [OperationContract]
        OpCalendarViewModel QueryOpCalendar(OpCalendar model);

        /// <summary>
        /// 依照ID查詢OpCalendar
        /// </summary>
        [OperationContract]
        OpCalendar GetOpCalendarByID(string iden);

        /// <summary>
        /// 取得當次RUN佣日期
        /// </summary>
        [OperationContract]
        OpCalendar GetsalrundateNow();

        /// <summary>
        /// 更新一筆調整起迄時間
        /// </summary>
        [OperationContract]
        bool UpdateOpCalendar(OpCalendar model);

        /// <summary>
        /// 查詢調整起迄時間LOG
        /// </summary>
        [OperationContract]
        List<OpCalendarLog> QueryAdjDateTimeUpateLog(OpCalendar model);

        /// <summary>
        /// 新增一筆OpCalendar紀錄
        /// </summary>
        [OperationContract]
        string InsertOpCalendar(OpCalendar model);

        /// <summary>
        /// 刪除OpCalendar紀錄
        /// </summary>
        /// <param name="iden">自動識別碼</param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteOpCalendarByIden(string iden, string UserID);

        /// <summary>
        /// 取得OpCalendar
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        List<OpCalendar> GetOpCalendar();

        /// <summary>
        /// 報表
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        Stream GetOpCalendarReportList(string productionYM);
        #endregion
    }
}
