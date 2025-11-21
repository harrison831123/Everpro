using EP.Platform.Service;
using EP.SD.SalesZone.AGUPG.Models;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EP.SD.SalesZone.AGUPG.Models.Enumerations;

namespace EP.SD.SalesZone.AGUPG.Service
{
    public class AGUPGService : IAGUPGService
    {
        /// <summary>
        /// 取得查詢資料
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public HrUpg25Dto GetQueryHrUpg25(HrUpg25QueryCondition model)
        {
            string[] v = model.YYYYSeason.Split('|');

            //EXEC usp_HrUpgGet25WebShow '2025','3','B22239015001'
            string sql = @"EXEC usp_HrUpgGet25WebShow @yyyy,@season,@agent_code";
            var result = DbHelper.QueryMultiple(VLifeRepository.ConnectionStringName, sql, new
            {
                yyyy = v[0],
                season = v[1],
                agent_code = model.AgentCode,
            },
            resultTypes: new Type[] { typeof(HrUpg25RstTitle), typeof(HrUpg25RstGrid1), typeof(HrUpg25RstGrid2), typeof(HrUpg25RstGrid3) });

            HrUpg25Dto dtResult = new HrUpg25Dto()
            {
                HrUpg25RstTitle = ((List<HrUpg25RstTitle>)result[0]).FirstOrDefault(),
                HrUpg25RstGrid1 = (List<HrUpg25RstGrid1>)(IEnumerable<HrUpg25RstGrid1>)result[1],
                HrUpg25RstGrid2 = (List<HrUpg25RstGrid2>)(IEnumerable<HrUpg25RstGrid2>)result[2],
                HrUpg25RstGrid3 = (List<HrUpg25RstGrid3>)(IEnumerable<HrUpg25RstGrid3>)result[3],
            };

            return dtResult;
        }

        /// <summary>
        /// 取得下拉選單
        /// </summary>
        /// <returns></returns>
        public List<HrUpg25RstViewModel.HrUpg25Item> GetHrUpg25Rst()
        {
            string sql = @"select top 3 '已核實FYC結算區間：'+FYCrange+' (目前計算截至'+FycYmSeq+')' as NowData,YYYY + '|' + Season as YYYYSeason
                            from HrUpg25Rst
                            group by FYCrange,FycYmSeq,YYYY,Season
                            order by FYCrange desc,FycYmSeq";
            var result = DbHelper.Query<HrUpg25RstViewModel.HrUpg25Item>(VLifeRepository.ConnectionStringName, sql).ToList();

            return result;
        }

        /// <summary>
        /// 取得明細資料
        /// </summary>
        /// EXEC usp_HrUpgGet25WebShowDetail '2025' ,'3','C12087994201' ,'4Season' --agentcode是業務員本人 
        /// EXEC usp_HrUpgGet25WebShowDetail '2025' ,''  ,'A22201016502' ,'IntroduceReturn' --agentcode是業務員本人
        /// EXEC usp_HrUpgGet25WebShowDetail '2025' ,'3','A22012735401' ,'RightEmpOM01' --agentcode是業務員本人 
        /// EXEC usp_HrUpgGet25WebShowDetail ''     ,''  ,'S22268780301' ,'VBPolicy'--agentcode是業務員本人
        /// <param name="model"></param>
        /// <returns></returns>
        public HrUpgGet25Dto GetQueryHrUpgGet25WebShowDetail(HrUpg25QueryCondition model)
        {
            string[] v = model.YYYYSeason.Split('|');
            string sql = @"EXEC usp_HrUpgGet25WebShowDetail @yyyy,@season,@agent_code,@DetailType";

            var parameters = new
            {
                yyyy = v.ElementAtOrDefault(0),
                season = v.ElementAtOrDefault(1),
                agent_code = model.AgentCode,
                DetailType = model.DetailType
            };

            var result = new HrUpgGet25Dto();

            // 建立查詢型別對應表
            var queryMap = new Dictionary<string, Action>()
            {
                ["RightEmpOM01"] = () => result.HrUpgGet25Detail1 =
                    DbHelper.Query<HrUpgGet25Detail1>(VLifeRepository.ConnectionStringName, sql, parameters).ToList(),

                ["4Season"] = () => result.HrUpgGet25Detail2 =
                    DbHelper.Query<HrUpgGet25Detail2>(VLifeRepository.ConnectionStringName, sql, parameters).ToList(),

                ["IntroduceReturn"] = () => result.HrUpgGet25Detail3 =
                    DbHelper.Query<HrUpgGet25Detail3>(VLifeRepository.ConnectionStringName, sql, parameters).ToList(),

                ["VBPolicy"] = () => result.HrUpgGet25Detail4 =
                    DbHelper.Query<HrUpgGet25Detail4>(VLifeRepository.ConnectionStringName, sql, parameters).ToList()
            };

            // 根據 DetailType 執行對應查詢(action)
            if (queryMap.TryGetValue(model.DetailType, out var action))
                action();

            return result;
        }

        /// <summary>
        /// 取得職等
        /// </summary>
        /// <param name="AgentCode"></param>
        /// <returns></returns>
        public string GetAgLevel(string AgentCode)
        {
            return DbHelper.Query<string>(VLifeRepository.ConnectionStringName, @"select ag_level from agid where agent_code=@AgentCode", new { AgentCode = AgentCode, }).FirstOrDefault();
        }

        /// <summary>
        /// 65以上主管畫面
        /// </summary>
        /// <param name="agentCode"></param>
        /// <returns></returns>
        public FamilyDto GetAdminFamilyTree(string agentCode)
        {
            #region sql
            string sql = @"SELECT a.*,acc.center_name,ums.um_code,um_name,an.ag_status_date,an.register_date
                    INTO #FamilyTree
                    FROM dbo.[family_tree](@AgentCode,'','1') a
                    LEFT JOIN accc acc on  a.center_code = acc.center_code
                    LEFT JOIN agum_set ums on a.agent_code=ums.um_leader_id
                    LEFT JOIN agin an on a.agent_code=an.agent_code
                    where an.ag_status_code in ('0','1')

                    select FYCrange,agent_code,case when Rst_Upgrade='-' then '' else'【已達】' end Rst_Upgrade
                    INTO #HrUpg25Rst
                    from HrUpg25Rst
                    where YYYY+Season =(select max(YYYY+Season)  from HrUpg25Rst)

                    SELECT '【'+a.center_name+'】' LEVEL_1,a.agent_code,a.agent_name,a.ag_level+'-'+v1.level_name_chs AgLevelName 
                    FROM #FamilyTree a
                    LEFT JOIN v_aglevel_occpind v1 on a.ag_level=v1.AG_LEVEL and a.ag_occp_ind=v1.ag_occp_ind
                    WHERE agent_code=@AgentCode

                    SELECT '第'+CONVERT(VARCHAR,b.mg_no)+'代 ' MgNo,a.agent_name,a.ag_level+'-'+v1.level_name_chs AgLevelName,rst.Rst_Upgrade,a.agent_code,ag_status_date,ag_status_code,a.register_date
                    FROM #FamilyTree a
                    INNER JOIN (select * from smrpt_m_cen_dir where proc_ym in (select max(proc_ym) proc_ym from smrpt_m_cen_dir where len(proc_ym)=7)) b on a.agent_code=b.agent_code
                    LEFT JOIN v_aglevel_occpind v1 on a.ag_level=v1.AG_LEVEL and a.ag_occp_ind=v1.ag_occp_ind
                    LEFT JOIN #HrUpg25Rst rst ON a.agent_code=rst.agent_code
                    WHERE a.agent_code<>@AgentCode
                    AND a.ag_level<'55'
                    ORDER BY b.mg_no

	                select '人力資料為「'+right(convert(char(10),Proc_Date,111),5)+'現實人力」且已排除「免評估ＴＳ～ＵＭ、身故人員、終止人員」' AgData
	                from (
	                Select Rpt_Date,max(Proc_Date) Proc_Date
	                From SMRPT_Log   
	                where Rpt_Name ='smrpt_m_cen_dir' 
	                and Rpt_Date in (select max(proc_ym) proc_ym from smrpt_m_cen_dir where len(proc_ym)=7) 
	                group by Rpt_Date
	                )a";
            #endregion

            var family = DbHelper.QueryMultiple(VLifeRepository.ConnectionStringName, sql
                 , new { AgentCode = agentCode }
                 , resultTypes: new Type[] { typeof(FamilyBoss), typeof(FamilyTree), typeof(string) });

            var result = new FamilyDto
            {
                FamilyBoss = ((IEnumerable<FamilyBoss>)family[0]).ToList(),
                FamilyTree = ((IEnumerable<FamilyTree>)family[1]).ToList(),
                AgData = ((IEnumerable<string>)family[2]).FirstOrDefault()
            };
            return result;

        }

        /// <summary>
        /// 依照身分別，取得畫面明細資料
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public FamilyDto GetHrUpg25RstFamilyTree(HrUpg25QueryCondition model)
        {
            //string agLevel = GetAgLevel(model.AgentCode);
            string umCode = string.Empty;
            string sql = string.Empty;
            var result = new FamilyDto();

            switch (model.UserType)
            {
                case AGUPGUserType.Admin:
                    return GetAdminFamilyTree(model.AgentCode);

                case AGUPGUserType.PreAdmin:
                    #region sql
                    sql = @"SELECT a.*,acc.center_name,ums.um_code,um_name,an.ag_status_date,an.register_date
                    INTO #FamilyTree
                    FROM dbo.[family_tree](@AgentCode,'','1') a
                    LEFT JOIN accc acc on  a.center_code = acc.center_code
                    LEFT JOIN agum_set ums on a.agent_code=ums.um_leader_id
                    LEFT JOIN agin an on a.agent_code=an.agent_code
                    where an.ag_status_code in ('0','1') 

                    SELECT um_code FROM #FamilyTree WHERE agent_code=@AgentCode AND ISNULL(um_code,'')<>''";
                    #endregion

                    umCode = DbHelper.Query<string>(VLifeRepository.ConnectionStringName, sql, new { agent_code = model.AgentCode, }).FirstOrDefault();
                    break;

                default:
                    return null;
            }

            return result;
        }
    }
}
