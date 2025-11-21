using EP.Platform.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Data;
using System.Collections.Generic;
using System.Linq;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    /// <summary>
    /// 客服業務系統的共用服務
    /// </summary>>
    public class CommonService : ICommonService
    {

        /// <summary>
        /// 新增資料設定
        /// </summary>
        /// <param name="model">要新增的資料</param>
        /// <returns>新增後的資料</returns>
        public CRMEDiscipType CreateCRMEDiscipType(CRMEDiscipType model)
        {
            model.Insert(new[] { H2ORepository.ConnectionStringName });
            return model;
        }

        /// <summary>
        /// 修改資料設定
        /// </summary>
        /// <param name="model">要修改的資料</param>
        /// <returns>修改後的資料</returns>
        public CRMEDiscipType UpdateCRMEDiscipType(CRMEDiscipType model)
        {
            model.Update(new { model.ID},  new[] { H2ORepository.ConnectionStringName });
            return model;
        }

        /// <summary>
        /// 刪除資料設定
        /// </summary>
        /// <param name="id">要刪除的資料的自動編號</param>
        public void DeleteCRMEDiscipType(int id)
        {
            H2ORepository.Delete<CRMEDiscipType>(new { ID = id });
        }

        /// <summary>
        /// 查詢資料設定的資料清單
        /// </summary>
        /// <param name="condition">查詢條件</param>
        /// <returns>資料設定的資料清單</returns>
        public IEnumerable<CRMEDiscipType> QueryCRMEDiscipTypeDatas(QueryDiscipTypeCondition condition)
        {
            var sql = @"
                Select *
                    From CRMEDiscipType
                Where 1 = 1
            ";

            #region sql條件

            if (condition != null)
            {
                if (condition.Code.HasValue)
                {
                    sql += " And Code = @Code";
                }

                // 代碼名稱模糊比對
                if (!string.IsNullOrEmpty(condition.Name))
                {
                    sql += " And Name like '%' + @Name + '%'";
                }

                if (condition.Kind.HasValue)
                {
                    sql += " And Kind = @Kind";
                }

                if (condition.Status.HasValue)
                {
                    sql += " And Status = @Status";
                }
            }
            #endregion

            sql += " order by Sort";

            var result = H2ORepository.Select<CRMEDiscipType>(sql, condition);

            return result;
        }

        /// <summary>
        /// 取得類型對應的案件的類別
        /// </summary>
        /// <param name="category">類型</param>
        /// <returns>類別清單</returns>
        public IEnumerable<CRMEDiscipType> GetCaseType(Category? category)
        {
            var condition = new QueryDiscipTypeCondition() { Kind = DiscipTypeKind.Type, Status = EnableStatus.Enabled };
            var datas = QueryCRMEDiscipTypeDatas(condition);
            if (category.HasValue)
            {
                var filter = new DiscipTypeCode[] { };
                if (category.Value == Category.Service)
                {
                    filter = new DiscipTypeCode[] { DiscipTypeCode.CS, DiscipTypeCode.SS };
                }
                else
                {
                    filter = new DiscipTypeCode[] { DiscipTypeCode.CC, DiscipTypeCode.SC };
                }
                return datas.Where(m => filter.Contains(m.Code));
            }
            return datas;
        }

        /// <summary>
        /// 依自動編號取得對應的資料設定
        /// </summary>
        /// <param name="id">自動編號</param>
        /// <returns>資料設定</returns>
        public CRMEDiscipType GetCRMEDiscipTypeByID(int id)
        {
            return H2ORepository.Select<CRMEDiscipType>(new { ID = id }).FirstOrDefault();
        }

        /// <summary>
        /// 依受理編號取得保單對應的通知對像id
        /// </summary>
        /// <param name="no">受理編號</param>
        /// <returns>保單對應的通知對像id清單</returns>
        public IEnumerable<string> GetDefaultNotifyMemberIdByCRMENo(string no)
        {
            var models = H2ORepository.Select<CRMEInsurancePolicy>(new { No = no });

            var memberIDs = new HashSet<string>();

            models.ForEach(m =>
            {
                if (!string.IsNullOrWhiteSpace(m.SUAgentCode))
                {
                    memberIDs.Add(m.SUAgentCode);
                    memberIDs.Add(m.SUVMLeaderID);
                    memberIDs.Add(m.SUSMLeaderID);
                    memberIDs.Add(m.SUCenterLeaderID);
                }
                else
                {
                    memberIDs.Add(m.AgentCode);
                    memberIDs.Add(m.VMLeaderID);
                    memberIDs.Add(m.SMLeaderID);
                    memberIDs.Add(m.CenterLeaderID);
                }
            });

            if (memberIDs.Contains("Z99999999901"))
            {
                memberIDs.Remove("Z99999999901");
            }

            return memberIDs.AsEnumerable();

        }
    }
}
