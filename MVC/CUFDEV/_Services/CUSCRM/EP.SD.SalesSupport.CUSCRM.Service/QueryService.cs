using EP.Platform.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Data;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    public class QueryService : IQueryService
    {
        #region h2o歷史Service


        /// <summary>
        /// 查歷史客服業務列表
        /// </summary>
        /// <returns></returns>

        public IEnumerable<HistoryCSViewModel> QueryHistoryCSGridDatas(HistoryQueryCondition condition)
        {
            if (condition.crmtype == 6)
                return QueryHistoryCCGridDatas(condition);
            string sql = @"select distinct c.company_code,c.company_name,a.crm_no,a.policy_no2,a.crm_douser,a.crm_content,a.crm_status,a.owner,a.insured,a.crm_closeday,a.crm_close,a.cov_no,convert(char(10),a.crm_closedate,111) as crm_closedate,a.crm_closedesc,a.crm_source,a.issdate,a.crm_dotype
         , b.crm_no,b.crm_type,b.crm_no_createname,convert(char(10), b.crm_no_createdate, 111) as crm_no_createdate
         ,e.vm_group_name,e.sm_name,e.wc_center_name,d.name as orgid_agname,f.name as now_agname,e.name as sub_agname

         ,e.vm_group_name as vmname,e.sm_name as smname,e.center_name as wc_centername
         ,f.vm_group_name as now_vmname,f.sm_name as now_smname,f.wc_center_name as now_wc_centername
          from(select distinct crm_content.crm_no, crm_douser, crm_status, crm_content, crm_insu.owner, crm_insu.insured, crm_closedate, crm_content.crm_source, crm_content.crm_dotype, crm_insu.policy_no2, crm_insu.issdate, crm_content.crm_closeday, crm_content.crm_close, crm_content.crm_closedesc, crm_insu.cov_no from crm_content, crm_insu where 2 > 1 and crm_content.crm_no = crm_insu.crm_no ";
            //申訴類型 必填
            sql += " and crm_type=@crmtype and  crm_insu.cov_no=1 and crm_insu.ismark=1";
            //結案狀態 必填
            if (condition.close_status == 1)
                sql += " and crm_status = 3 ";
            else
                sql += " and crm_status not in (3) ";

            if (!String.IsNullOrWhiteSpace(condition.crm_no))
                sql += "  and crm_insu.crm_no like '%'+@crm_no+'%' ";
            if (!String.IsNullOrWhiteSpace(condition.source))
                sql += " and crm_source =@source ";
            if (!String.IsNullOrWhiteSpace(condition.do_type))
                sql += " and crm_dotype =@do_type ";
            #region 若為結案狀態 另外union
            if (condition.close_status == 1)
            {
                sql += @"union
			 select distinct crm_content.crm_no,crm_douser,crm_status,crm_content,crm_insu.owner,crm_insu.insured,crm_closedate,crm_content.crm_source,crm_content.crm_dotype,crm_insu.policy_no2,crm_insu.issdate,crm_content.crm_closeday,crm_content.crm_close,crm_content.crm_closedesc,crm_insu.cov_no from crm_content,crm_insu,crm_close_log
            where 2>1 and crm_content.crm_no=crm_insu.crm_no  and crm_type=@crmtype and crm_insu.cov_no=1 and crm_status not in(3)  and crm_content.crm_no=crm_close_log.crm_no ";
                if (!String.IsNullOrWhiteSpace(condition.crm_no))
                    sql += "  and crm_insu.crm_no like '%'+@crm_no+'%' ";
                if (!String.IsNullOrWhiteSpace(condition.source))
                    sql += " and crm_source =@source ";
                if (!String.IsNullOrWhiteSpace(condition.do_type))
                    sql += " and crm_dotype =@do_type ";
            }
            #endregion



            sql += @"   ) as a 
         left join
        (select crm_no, crm_type, crm_no_createname, crm_no_createdate from crm_no group by crm_no, crm_type, crm_no_createname, crm_no_createdate) as b
         on a.crm_no = b.crm_no
         left join
        (select* from crm_insu where cov_no= 1) as c on a.policy_no2 = c.policy_no2 and c.crm_no = a.crm_no
         left join
        (select s.*from crm_orgin_agent s
        left join crm_now_agent u on u.crm_no = s.crm_no and u.policy_no2 = s.policy_no2
         left join crm_insu t on s.crm_no = t.crm_no and s.policy_no2 = t.policy_no2 and t.cov_no = 1
         where((s.ag_status_code = u.ag_status_code and s.name = u.name) or(s.ag_status_code <> u.ag_status_code and u.insu_agent_code = s.agent_code)) 
        ) as d on a.policy_no2 = d.policy_no2 and d.crm_no = a.crm_no
         left join
        (select distinct x.crm_no, x.policy_no2, x.vm_group_name, x.sm_name, x.wc_center_name, x.agent_code, x.name, x.ag_status_code, x.insu_agent_code, y.now_agent_code,x.admin_name  from crm_now_agent x
         left join crm_sub_agent y on x.crm_no= y.crm_no and x.policy_no2= y.policy_no2 and x.ag_status_code in (2,3) and x.agent_code = y.now_agent_code
         left join crm_orgin_agent z on x.crm_no = z.crm_no and x.policy_no2 = z.policy_no2
         and z.ag_status_code in (2, 3) and z.agent_code = x.insu_agent_code and(x.ag_status_code in (0, 1) and y.name is null) and(x.ag_status_code in (2, 3) and x.insu_agent_code = z.agent_code and y.now_agent_code = x.agent_code)
        ) as f on a.policy_no2 = f.policy_no2 and f.crm_no = a.crm_no
         and((f.insu_agent_code = d.agent_code and d.ag_status_code = f.ag_status_code) or(f.ag_status_code in (0, 1) and d.ag_status_code in (2, 3) and f.insu_agent_code = d.agent_code)) or(f.ag_status_code in (0, 1) and a.policy_no2 = f.policy_no2 and f.crm_no = a.crm_no and d.name is null)
         or(d.name is null and f.ag_status_code in (2, 3) and f.agent_code = f.insu_agent_code and f.now_agent_code = f.agent_code and a.policy_no2 = f.policy_no2 and f.crm_no = a.crm_no)
         left join
         crm_sub_agent e on a.policy_no2 = e.policy_no2 and e.crm_no = a.crm_no and e.now_agent_code = f.agent_code
         and(d.ag_status_code in (2, 3) and f.ag_status_code in (2, 3)) and((e.name is not null or e.name <> '') and(f.name is not null or f.name <> ''))
         or(f.ag_status_code in (2, 3) and f.agent_code = e.now_agent_code and a.policy_no2 = e.policy_no2 and e.crm_no = a.crm_no)

         where 2 > 1 ";
            sql += " and convert(char(10),b.crm_no_createdate,111) between @startdate and @closedate ";
            if (!String.IsNullOrWhiteSpace(condition.issdate))
                sql += " and left(c.issdate,4)=@issdate ";
            if (!String.IsNullOrWhiteSpace(condition.company_code))
                sql += " and c.company_code=@company_code ";
            if (!String.IsNullOrWhiteSpace(condition.policy_no2))
                sql += " and c.policy_no2=@policy_no2 ";
            if (!String.IsNullOrWhiteSpace(condition.owner))
                sql += " and c.owner like '%'+@owner+'%' ";

            if (!String.IsNullOrWhiteSpace(condition.agent_name))
                sql += " and (c.agent_name like '%'+@agent_name+'%' or d.name like '%'+@agent_name+'%' or e.name like '%'+@agent_name+'%' or f.name like '%'+@agent_name+'%') ";
            if (!String.IsNullOrWhiteSpace(condition.center_manager))
                sql += " and ( d.admin_name like '%'+@center_manager+'%' or e.admin_name like '%'+@center_manager+'%' or f.admin_name like '%'+@center_manager+'%' ) ";
            if (!String.IsNullOrWhiteSpace(condition.worker))
                sql += " and a.crm_douser like '%'+@worker+'%' ";
            sql += "order by a.crm_no";
            var result = DbHelper.Query<HistoryCSViewModel>(OldH2ORepository.ConnectionStringName, sql, condition);
            return result;
        }

        /// <summary>
        /// 查歷史客服業務列表(CC)
        /// </summary>
        /// <returns></returns>

        public IEnumerable<HistoryCSViewModel> QueryHistoryCCGridDatas(HistoryQueryCondition condition)
        {

            string sql = @"select distinct c.company_code,c.company_name,a.crm_no,a.policy_no2,a.crm_douser,a.crm_content,a.crm_status,a.owner,a.insured,a.crm_closeday,a.crm_close,a.cov_no,convert(char(10),a.crm_closedate,111) as crm_closedate,a.crm_closedesc,a.crm_source,a.issdate,a.crm_dotype
         ,b.crm_no,b.crm_type,b.crm_no_createname,convert(char(10),b.crm_no_createdate,111) as crm_no_createdate
         ,e.vm_group_name,e.sm_name,e.wc_center_name,d.name as orgid_agname,f.name as now_agname,e.name as sub_agname
         ,d.vm_group_name as vmname,d.sm_name as smname,d.wc_center_name as wc_centername
         ,e.vm_group_name as vmname,e.sm_name as smname,e.center_name as wc_centername
         ,f.vm_group_name as now_vmname,f.sm_name as now_smname,f.wc_center_name as now_wc_centername 
          from  (
		  select distinct crm_content.crm_no,crm_douser,crm_status,crm_content,crm_insu.owner,crm_insu.insured,crm_closedate,crm_content.crm_source,crm_content.crm_dotype,crm_insu.policy_no2,crm_insu.issdate,crm_content.crm_closeday,crm_content.crm_close,crm_content.crm_closedesc,crm_insu.cov_no from crm_content,crm_insu
		  where 2>1
		  and crm_type=6 and crm_insu.cov_no=1 and crm_content.crm_no=crm_insu.crm_no and crm_insu.ismark=1";
            //結案狀態 必填
            if (condition.close_status == 1)
                sql += " and crm_status = 3 ";
            else
                sql += " and crm_status not in (3) ";

            if (!String.IsNullOrWhiteSpace(condition.crm_no))
                sql += "  and crm_insu.crm_no like '%'+@crm_no+'%' ";
            if (!String.IsNullOrWhiteSpace(condition.source))
                sql += " and crm_source =@source ";
            if (!String.IsNullOrWhiteSpace(condition.do_type))
                sql += " and crm_dotype =@do_type ";
            #region 若為結案狀態 另外union
            if (condition.close_status == 1)
            {
                sql += @"union
			 select distinct crm_content.crm_no,crm_douser,crm_status,crm_content,crm_insu.owner,crm_insu.insured,crm_closedate,crm_content.crm_source,crm_content.crm_dotype,crm_insu.policy_no2,crm_insu.issdate,crm_content.crm_closeday,crm_content.crm_close,crm_content.crm_closedesc,crm_insu.cov_no from crm_content,crm_insu,crm_close_log
            where 2>1 and crm_content.crm_no=crm_insu.crm_no  and crm_type=@crmtype and crm_insu.cov_no=1 and crm_status not in(3)  and crm_content.crm_no=crm_close_log.crm_no ";
                if (!String.IsNullOrWhiteSpace(condition.crm_no))
                    sql += "  and crm_insu.crm_no like '%'+@crm_no+'%' ";
                if (!String.IsNullOrWhiteSpace(condition.source))
                    sql += " and crm_source =@source ";
                if (!String.IsNullOrWhiteSpace(condition.do_type))
                    sql += " and crm_dotype =@do_type ";
            }
            #endregion
            sql += @"   ) as a 
             left join
            (select crm_no, crm_type, crm_no_createname, crm_no_createdate from crm_no group by crm_no, crm_type, crm_no_createname, crm_no_createdate) as b
             on a.crm_no = b.crm_no
             left join
            (select* from crm_insu where cov_no= 1) as c on a.policy_no2 = c.policy_no2 and c.crm_no = a.crm_no
             left join
            (select s.*from crm_orgin_agent s
            left join crm_now_agent u on u.crm_no = s.crm_no and u.policy_no2 = s.policy_no2
             left join crm_insu t on s.crm_no = t.crm_no and s.policy_no2 = t.policy_no2 and t.cov_no = 1
             where((s.ag_status_code = u.ag_status_code and s.name = u.name) or(s.ag_status_code <> u.ag_status_code and u.insu_agent_code = s.agent_code)) 
            ) as d on a.policy_no2 = d.policy_no2 and d.crm_no = a.crm_no
             left join
            (select distinct x.crm_no, x.policy_no2, x.vm_group_name, x.sm_name, x.wc_center_name, x.agent_code, x.name, x.ag_status_code, x.insu_agent_code, y.now_agent_code,x.admin_name  from crm_now_agent x
             left join crm_sub_agent y on x.crm_no= y.crm_no and x.policy_no2= y.policy_no2 and x.ag_status_code in (2,3) and x.agent_code = y.now_agent_code
             left join crm_orgin_agent z on x.crm_no = z.crm_no and x.policy_no2 = z.policy_no2
             and z.ag_status_code in (2, 3) and z.agent_code = x.insu_agent_code and(x.ag_status_code in (0, 1) and y.name is null) and(x.ag_status_code in (2, 3) and x.insu_agent_code = z.agent_code and y.now_agent_code = x.agent_code)
            ) as f on a.policy_no2 = f.policy_no2 and f.crm_no = a.crm_no
             and((f.insu_agent_code = d.agent_code and d.ag_status_code = f.ag_status_code) or(f.ag_status_code in (0, 1) and d.ag_status_code in (2, 3) and f.insu_agent_code = d.agent_code)) or(f.ag_status_code in (0, 1) and a.policy_no2 = f.policy_no2 and f.crm_no = a.crm_no and d.name is null)
             or(d.name is null and f.ag_status_code in (2, 3) and f.agent_code = f.insu_agent_code and f.now_agent_code = f.agent_code and a.policy_no2 = f.policy_no2 and f.crm_no = a.crm_no)
             left join
             crm_sub_agent e on a.policy_no2 = e.policy_no2 and e.crm_no = a.crm_no and e.now_agent_code = f.agent_code
             and(d.ag_status_code in (2, 3) and f.ag_status_code in (2, 3)) and((e.name is not null or e.name <> '') and(f.name is not null or f.name <> ''))
             or(f.ag_status_code in (2, 3) and f.agent_code = e.now_agent_code and a.policy_no2 = e.policy_no2 and e.crm_no = a.crm_no)

             where 2 > 1 ";
            sql += " and convert(char(10),b.crm_no_createdate,111) between @startdate and @closedate ";
            if (!String.IsNullOrWhiteSpace(condition.issdate))
                sql += " and left(c.issdate,4)=@issdate ";
            if (!String.IsNullOrWhiteSpace(condition.company_code))
                sql += " and c.company_code=@company_code ";
            if (!String.IsNullOrWhiteSpace(condition.policy_no2))
                sql += " and c.policy_no2=@policy_no2 ";
            if (!String.IsNullOrWhiteSpace(condition.owner))
                sql += " and c.owner like '%'+@owner+'%' ";

            if (!String.IsNullOrWhiteSpace(condition.agent_name))
                sql += " and (c.agent_name like '%'+@agent_name+'%' or d.name like '%'+@agent_name+'%' or e.name like '%'+@agent_name+'%' or f.name like '%'+@agent_name+'%') ";
            if (!String.IsNullOrWhiteSpace(condition.center_manager))
                sql += " and ( d.admin_name like '%'+@center_manager+'%' or e.admin_name like '%'+@center_manager+'%' or f.admin_name like '%'+@center_manager+'%' ) ";
            if (!String.IsNullOrWhiteSpace(condition.worker))
                sql += " and a.crm_douser like '%'+@worker+'%' ";
            sql += "order by a.crm_no";
            var result = DbHelper.Query<HistoryCSViewModel>(OldH2ORepository.ConnectionStringName, sql, condition);
            return result;
        }



        /// <summary>
        /// 查歷史聯繫紀錄列表
        /// </summary>
        /// <param name="crm_no">受理編號</param>
        /// <returns>歷史聯繫紀錄列表</returns>
        public IEnumerable<tcrm_do> QueryHistoryMaintainRecordList(string crm_no)
        {
            string sql = "select* from crm_do where crm_no =@crm_no";
            var result = DbHelper.Query<tcrm_do>(OldH2ORepository.ConnectionStringName, sql, new { crm_no }).ToList();
            return result;
        }
        /// <summary>
        /// 查歷史聯繫紀錄檔案列表
        /// </summary>
        /// <param name="crm_no">受理編號</param>
        /// <returns>歷史聯繫紀錄檔案列表</returns>
        public IEnumerable<crm_do_file> QueryHistoryMaintainFileList(string crm_no)
        {
            string sql = "select* from crm_do_file where crm_no =@crm_no";
            var result = DbHelper.Query<crm_do_file>(OldH2ORepository.ConnectionStringName, sql, new { crm_no }).ToList();
            return result;
        }
        #region 維護&結案
        /// <summary>
        /// 新增聯繫紀錄
        /// </summary>
        /// <param name="model">聯繫紀錄</param>

        public void CreateCrm_do(tcrm_do model)
        {
            var result = model.Insert(OldH2ORepository.ConnectionStringName);
        }
        /// <summary>
        /// 新增上傳檔案
        /// </summary>
        /// <param name="model"></param>
        public void CreateCrm_do_file(crm_do_file model)
        {
            var result = model.Insert(OldH2ORepository.ConnectionStringName);

        }
        /// <summary>
        /// 更新crm_content
        /// </summary>
        /// <param name="model"></param>
        public void UpdateCrm_ContentForClose(crm_close_log model)
        {

            string sql = @"Update crm_content set crm_status= 3,
            crm_close=@crm_close,
            crm_closeday=@crm_closeday ,
            crm_closerid=@crm_closerid,
            crm_closer=@crm_closer,
            crm_closedate=@crm_closedate,
            crm_closedesc=@crm_closedesc,
            crm_closecreatedate=@crm_closecreatedate,
            crm_close_n=@crm_close
            where crm_no=@crm_no";
            var condition = new
            {
                crm_close = model.new_crm_close ,
                crm_closerid = model.creater_orgid,
                crm_closer = GetMemberFormat(model.creater_orgid),
                crm_closedate = model.new_closedate,
                crm_closeday = model.crm_closeday,
                crm_closedesc = model.desc_log,
                crm_closecreatedate = DateTime.Now,
                crm_no = model.crm_no,
            };
            var result = DbHelper.Execute(OldH2ORepository.ConnectionStringName, sql, condition);
        }
        /// <summary>
        /// 依orgid 獲得部門+性名 
        /// </summary>
        /// <param name="orgid"></param>
        /// <returns></returns>
        public string GetMemberFormat(int orgid) {
            string sql = "select orgname from org where orgid =@orgid";
            return DbHelper.Query<string>(OldH2ORepository.ConnectionStringName, sql, new { orgid }).FirstOrDefault();

        }
        /// <summary>
        /// 結案
        /// </summary>
        /// <param name="model"></param>
        public void CreateCloseRecord(crm_close_log model)
        {
            //寫入crm_close_log
            model.crm_closeday = GetHistoryCloseDay(model.new_closedate, model.crm_no);
            model.Insert(OldH2ORepository.ConnectionStringName);
            //寫入crm_content
            UpdateCrm_ContentForClose(model);
        }
        /// <summary>
        /// 獲得結案天數(歷史資料用)
        /// </summary>
        /// <param name="closetime"></param>
        /// <param name="crm_no"></param>
        /// <returns></returns>
        public int GetHistoryCloseDay(DateTime closetime, string crm_no) {
            string sql = "select crm_createdate from crm_content where crm_no=@crm_no";
            var createtime=DbHelper.Query<DateTime>(OldH2ORepository.ConnectionStringName, sql, new { crm_no }).FirstOrDefault();
            TimeSpan span =closetime.Subtract(createtime);
            return ((int)span.TotalDays + 1);
        }
        #endregion


        #endregion

        #region 查詢報表

        public Stream GetReport(QueryReportCondition condition)
        {
            switch (condition.code)
            {
                case "CS":
                    return GetCSReport(condition);
                case "SS":
                    return GetSSReport(condition);
                case "CC":
                    return GetCCReport(condition);
                case "SC":
                    return GetSCReport(condition);
                default:
                    break;
            }
            Throw.BusinessError("報表類型錯誤");
            return null;
        }

        /// <summary>
        /// CS報表
        /// </summary>
        public Stream GetCSReport(QueryReportCondition condition)
        {
            try
            {
                #region SQL 語法
                string sql = @" SELECT distinct
              CASE  WHEN (C.CODE='CS' or C.CODE='SS') THEN '客服'  ELSE '申訴' END AS Category, 
              C.CODE,C.No,CDT.Name as SourceType,
              CASE WHEN ISNULL(CCC.DoUser,'')<>'' THEN V1.nmember ELSE V.nmember END AS Creator,
              Month(C.CreateTime)as m,DaY(C.CreateTime) as d, 
              ISNULL(STUFF(CompanyNameList.CompanyNames, 1, 1, ''), '') AS CompanyNames, 
              ISNULL(STUFF(OwnerList.Owners, 1, 1, ''), '') AS Owners, 
              ISNULL(STUFF(InsuredList.Insureds, 1, 1, ''), '') AS Insureds, 
              ISNULL(STUFF(IssDateList.IssDates, 1, 1, ''), '') AS IssDates, 
              ISNULL(STUFF(PolicyList.PolicyValue, 1, 1, ''), '') AS PolicyNos, 
               CDT1.Name as CaseType,Content, 
             ISNULL(STUFF(VMNameList.VMNames, 1, 1, ''), '') AS VMNames, 
              ISNULL(STUFF(SMNameList.SMNames, 1, 1, ''), '') AS SMNames, 
               ISNULL(STUFF(WCNameList.WCNames, 1, 1, ''), '') AS WCNames, 
               ISNULL(STUFF(CenterNameList.CenterNames, 1, 1, ''), '') AS CenterNames, 
               ISNULL(STUFF(AgentNameList.AgentNames, 1, 1, ''), '') AS AgentNames, 
               ISNULL(STUFF(SUWCNameList.SUWCNames, 1, 1, ''), '') AS SUWCNames, 
			   ISNULL(STUFF(SUCenterNameList.SUCenterNames, 1, 1, ''), '') AS SUCenterNames, 
               ISNULL(STUFF(SUAgentNameList.SUAgentNames, 1, 1, ''), '') AS SUAgentNames, 
             CONVERT(VARCHAR,CCL.CreateTime,111) as CloseDay, 
              CASE WHEN (ISNULL(CCL.CreateTime,'')<>'') THEN DATEDIFF(d,CCC.CreateTime,CCL.CreateTime) ELSE DATEDIFF(d,C.CreateTime,getdate()) END AS DoDay,
              ISNULL(STUFF(CRMEDoList.ContentValue, 1, 1, ''), '') AS Contents 
              FROM( 
              select * from [H2O].[dbo].[CRMENo] where isnull(No,'')<>'' 
              and CODE=@code ";
                if (!string.IsNullOrWhiteSpace(condition.crm_no))
                    sql += @" and No=@crm_no ";
                if (condition.startdate.HasValue)
                    sql += @" and CreateTime >= @startdate ";
                if (condition.closedate.HasValue)
                    sql += @" and CreateTime <= @closedate +' 23:59:59' ";
                sql += @" ) C 
              LEFT JOIN  [H2O].[dbo].[CRMECaseContent] CCC ON C.No=CCC.No 
              LEFT JOIN  [H2O].[dbo].[CRMEInsurancePolicy] CIP ON C.No=CIP.No 
              LEFT JOIN [H2O].[dbo].[CRMEDiscipType] CDT ON CCC.SourceID=CDT.ID 
	          LEFT JOIN [H2O].[dbo].[CRMEDiscipType] CDT1 ON CCC.CaseTypeID=CDT1.ID 
              LEFT JOIN [CUF].[dbo].[v_member_account] V ON V.imember=C.Creator
	          LEFT JOIN (SELECT No,ResultCode,Creator,MaxCCL.createtime  FROM CRMECloseLog,(SELECT MAX(createtime) AS createtime FROM CRMECloseLog GROUP BY No) AS MaxCCL WHERE CRMECloseLog.createtime = MaxCCL.createtime)  CCL ON C.No=CCL.No  
			  LEFT JOIN [CUF].[dbo].[v_member_account] V1 ON V1.imember=CCC.DoUser
              OUTER APPLY (
                SELECT '|' + CD.Content 
                FROM dbo.CRMEDo CD
                WHERE C.No = CD.No
                FOR XML PATH ('')
            ) CRMEDoList (ContentValue)
            OUTER APPLY (
                SELECT distinct ',' + CIP.CompanyName 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) CompanyNameList (CompanyNames)
            OUTER APPLY (
                SELECT distinct ',' + CIP.Owner 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) OwnerList (Owners)
            OUTER APPLY (
                SELECT distinct ',' + CIP.Insured
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) InsuredList (Insureds)
            OUTER APPLY (
                SELECT distinct ',' + CIP.IssDate
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) IssDateList (IssDates)
             OUTER APPLY (
                SELECT distinct ',' + CIP.PolicyNo 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) PolicyList (PolicyValue)
             OUTER APPLY (
                SELECT	distinct ',' + CIP.VMName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) VMNameList (VMNames)
             OUTER APPLY (
                SELECT distinct ',' + CIP.SMName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SMNameList (SMNames) 
             OUTER APPLY (
                SELECT	distinct ',' + CIP.WCName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) WCNameList (WCNames)
             OUTER APPLY (
                SELECT	distinct ',' + CIP.centerName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) CenterNameList (CenterNames)
            OUTER APPLY (
                SELECT 	distinct ',' + CIP.AgentName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) AgentNameList (AgentNames)
            OUTER APPLY (
                SELECT	distinct ',' + CIP.SUWCName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUWCNameList (SUWCNames)
			 OUTER APPLY (
                SELECT	distinct ',' + CIP.SUCenterName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUCenterNameList (SUCenterNames)
            OUTER APPLY (
                SELECT	distinct ',' + CIP.SUAgentName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUAgentNameList (SUAgentNames)
            where 1=1 ";
                if (!string.IsNullOrWhiteSpace(condition.PolicyNo))
                    sql += @" and PolicyNo like '%'+@PolicyNo+'%' ";
                if (!string.IsNullOrWhiteSpace(condition.owner))
                    sql += @" and owner like '%'+@owner+'%' ";
                switch (condition.close_status)
                {
                    case 0:
                        sql += @" and isnull( CCL.CreateTime,'')='' ";
                        break;
                    case 1:
                        sql += @" and isnull( CCL.CreateTime,'')<>'' ";
                        break;
                }
                if (!string.IsNullOrWhiteSpace(condition.worker))
                    sql += @" and V1.nmember like '%'+@worker+'%' OR (ISNULL(CCC.DoUser,'')='' AND V.nmember like '%'+@worker+'%') ";

                #endregion

                var models = DbHelper.Query<dynamic>(OldH2ORepository.ConnectionStringName, sql, condition).Select(m => (IDictionary<string, object>)m);
                MemoryStream ms = new MemoryStream();
                ExcelPackage excel = new ExcelPackage();

                // 頁籤名稱設定
                ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("客服業務報表");
                var rowPos = 0;
                var colPos = 0;
                rowPos++;
                string[] csArr = new string[] { "序號", "類別", "類型", "照會單號", "來源","承辦人", "受理月", "受理日", "保險公司", "要保人",
                "被保險人", "生效日", "保單號碼", "案件類別", "照會內容", "副總團隊", "協理體系", "經手人實駐", "經手人處別", "原經手人","服務人員實駐","服務人員處別", "服務人員",
                "結案日", "處理天數", "處理過程" };

                foreach (string title in csArr)
                {
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = title;
                }
                models.ForEach(m =>
                {
                    rowPos++;
                    colPos = 1;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = rowPos - 1;//序號     
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Category"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CODE"));
                    colPos++; ;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("No"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SourceType"));
                    colPos++;
                    //新增承辦人(立案人)
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Creator"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("m"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("d"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CompanyNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Owners")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Insureds")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("IssDates")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("PolicyNos")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CaseType"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Content"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("VMNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SMNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("WCNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CenterNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("AgentNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUWCNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUCenterNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUAgentNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CloseDay"));//結案日
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("DoDay"));//處理天數
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Contents")).Replace("|", "\n");//處理流程
                });

                // 自動調整欄位大小
                for (int i = 1; i <= csArr.Length; i++)
                {
                    sheet.Column(i).AutoFit();
                    if (sheet.Column(i).Width > 75)
                        sheet.Column(i).Width = 75;
                }

                var cells = sheet.Cells[1, 1, rowPos, csArr.Length];
                cells.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cells.Style.WrapText = true;
                excel.SaveAs(ms);
                excel.Dispose();
                ms.Position = 0;
                return ms;
            }
            catch (Exception ex)
            {

            }
            return null;

        }
        /// <summary>
        /// SS報表
        /// </summary>
        public Stream GetSSReport(QueryReportCondition condition)
        {
            try
            {
                #region SQL 語法
                string sql = @" SELECT distinct
              CASE  WHEN (C.CODE='CS' or C.CODE='SS') THEN '客服'  ELSE '申訴' END AS Category, 
              C.CODE,C.No,CDT.Name as SourceType,
              V.nmember AS Creator,
              Month(C.CreateTime)as m,DaY(C.CreateTime) as d, 
              ISNULL(STUFF(CompanyNameList.CompanyNames, 1, 1, ''), '') AS CompanyNames, 
              ISNULL(STUFF(OwnerList.Owners, 1, 1, ''), '') AS Owners, 
              ISNULL(STUFF(InsuredList.Insureds, 1, 1, ''), '') AS Insureds, 
              ISNULL(STUFF(IssDateList.IssDates, 1, 1, ''), '') AS IssDates, 
              ISNULL(STUFF(PolicyList.PolicyValue, 1, 1, ''), '') AS PolicyNos, 
               CDT1.Name as CaseType,Content, 
             ISNULL(STUFF(VMNameList.VMNames, 1, 1, ''), '') AS VMNames, 
              ISNULL(STUFF(SMNameList.SMNames, 1, 1, ''), '') AS SMNames, 
               ISNULL(STUFF(WCNameList.WCNames, 1, 1, ''), '') AS WCNames, 
               ISNULL(STUFF(CenterNameList.CenterNames, 1, 1, ''), '') AS CenterNames, 
               ISNULL(STUFF(AgentNameList.AgentNames, 1, 1, ''), '') AS AgentNames, 
               ISNULL(STUFF(SUWCNameList.SUWCNames, 1, 1, ''), '') AS SUWCNames, 
			   ISNULL(STUFF(SUCenterNameList.SUCenterNames, 1, 1, ''), '') AS SUCenterNames, 
               ISNULL(STUFF(SUAgentNameList.SUAgentNames, 1, 1, ''), '') AS SUAgentNames, 
             ISNULL(STUFF(BUSContactCompanyDateList.BUSContactCompanyDates, 1, 1, ''), '') AS BUSContactCompanyDates, 
			   ISNULL(STUFF(ReplyCompanyDateList.ReplyCompanyDates, 1, 1, ''), '') AS ReplyCompanyDates, 
			   ISNULL(STUFF(ReplyYallowBillDateList.ReplyYallowBillDates, 1, 1, ''), '') AS ReplyYallowBillDates, 
              CASE WHEN (ISNULL(CD.BUSContactCompanyDate,'')<>'' and ISNULL(CD.ReplyYallowBillDate,'')<>'' ) THEN DATEDIFF(d,CD.BUSContactCompanyDate,CD.ReplyYallowBillDate) ELSE DATEDIFF(d,C.CreateTime,getdate()) END AS DoDay,
              ISNULL(CDT2.Name,'') AS CompanyResult ,
              ISNULL(STUFF(CRMEDoList.ContentValue, 1, 1, ''), '') AS Contents ,
              ISNULL(CDT3.Name,'') AS Result
              FROM( 
              select * from [H2O].[dbo].[CRMENo] where isnull(No,'')<>'' 
              and CODE=@code ";
                if (!string.IsNullOrWhiteSpace(condition.crm_no))
                    sql += @" and No=@crm_no ";
                if (condition.startdate.HasValue)
                    sql += @" and CreateTime >= @startdate ";
                if (condition.closedate.HasValue)
                    sql += @" and CreateTime <= @closedate +' 23:59:59' ";
                sql += @" ) C 
              LEFT JOIN [H2O].[dbo].[CRMECaseContent] CCC ON C.No=CCC.No 
              LEFT JOIN [H2O].[dbo].[CRMEInsurancePolicy] CIP ON C.No=CIP.No 
              LEFT JOIN [H2O].[dbo].[CRMEDiscipType] CDT ON CCC.SourceID=CDT.ID 
	          LEFT JOIN [H2O].[dbo].[CRMEDiscipType] CDT1 ON CCC.CaseTypeID=CDT1.ID 
              LEFT JOIN [CUF].[dbo].[v_member_account] V ON V.imember=C.Creator
              LEFT JOIN (SELECT No,ResultCode,Creator,MaxCCL.createtime  FROM CRMECloseLog,(SELECT MAX(createtime) AS createtime FROM CRMECloseLog GROUP BY No) AS MaxCCL WHERE CRMECloseLog.createtime = MaxCCL.createtime)  CCL ON C.No=CCL.No  
			   LEFT JOIN (SELECT No,CRMEDo.CompanyResult,MaxCD.createtime  FROM CRMEDo,(SELECT MAX(createtime) AS createtime FROM CRMEDo GROUP BY No) AS MaxCD WHERE CRMEDo.createtime = MaxCD.createtime)  CD2 ON C.No=CD2.No  
              LEFT JOIN [H2O].[dbo].[CRMEDiscipType] CDT2 ON CD2.CompanyResult=CDT2.ID 
              LEFT JOIN [H2O].[dbo].[CRMEDiscipType] CDT3 ON CCL.ResultCode=CDT3.ID 
	          LEFT JOIN (SELECT NO,MIN(BUSContactCompanyDate) AS BUSContactCompanyDate ,MAX(ReplyYallowBillDate)  AS ReplyYallowBillDate FROM [H2O].[dbo].CRMEDo GROUP BY No)AS CD ON C.No=CD.No 
              OUTER APPLY (
                SELECT '|' + CD.Content 
                FROM dbo.CRMEDo CD
                WHERE C.No = CD.No
                FOR XML PATH ('')
            ) CRMEDoList (ContentValue)
            OUTER APPLY (
                SELECT distinct ',' + CIP.CompanyName 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) CompanyNameList (CompanyNames)
            OUTER APPLY (
                SELECT distinct ',' + CIP.Owner 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) OwnerList (Owners)
            OUTER APPLY (
                SELECT distinct ',' + CIP.Insured
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) InsuredList (Insureds)
            OUTER APPLY (
                SELECT distinct ',' + CIP.IssDate
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) IssDateList (IssDates)
             OUTER APPLY (
                SELECT distinct ',' + CIP.PolicyNo 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) PolicyList (PolicyValue)
             OUTER APPLY (
                SELECT	distinct ',' + CIP.VMName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) VMNameList (VMNames)
             OUTER APPLY (
                SELECT distinct ',' + CIP.SMName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SMNameList (SMNames)
             OUTER APPLY (
                SELECT	distinct ',' + CIP.WCName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) WCNameList (WCNames)
             OUTER APPLY (
                SELECT	distinct ',' + CIP.centerName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) CenterNameList (CenterNames)
            OUTER APPLY (
                SELECT 	distinct ',' + CIP.AgentName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) AgentNameList (AgentNames)
            OUTER APPLY (
                SELECT	distinct ',' + CIP.SUWCName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUWCNameList (SUWCNames)
			  OUTER APPLY (
                SELECT	distinct ',' + CIP.SUCenterName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUCenterNameList (SUCenterNames)
            OUTER APPLY (
                SELECT	distinct ',' + CIP.SUAgentName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUAgentNameList (SUAgentNames)
			OUTER APPLY (
                SELECT	distinct ',' + CONVERT(VARCHAR,CD.BUSContactCompanyDate,111)
                FROM dbo.CRMEDo CD
                WHERE C.No = CD.No
                FOR XML PATH ('')
            ) BUSContactCompanyDateList (BUSContactCompanyDates)
			OUTER APPLY (
                SELECT	distinct ',' + CONVERT(VARCHAR,CD.ReplyCompanyDate,111)
                FROM dbo.CRMEDo CD
                WHERE C.No = CD.No
                FOR XML PATH ('')
            ) ReplyCompanyDateList (ReplyCompanyDates)
			OUTER APPLY (
                SELECT	distinct ',' + CONVERT(VARCHAR,CD.ReplyYallowBillDate,111)
                FROM dbo.CRMEDo CD
                WHERE C.No = CD.No
                FOR XML PATH ('')
            ) ReplyYallowBillDateList (ReplyYallowBillDates)
            where 1=1 ";
                if (!string.IsNullOrWhiteSpace(condition.PolicyNo))
                    sql += @" and PolicyNo like '%'+@PolicyNo+'%' ";
                if (!string.IsNullOrWhiteSpace(condition.owner))
                    sql += @" and owner like '%'+@owner+'%' ";
                switch (condition.close_status)
                {
                    case 0:
                        sql += @" and isnull( CCL.CreateTime,'')='' ";
                        break;
                    case 1:
                        sql += @" and isnull( CCL.CreateTime,'')<>'' ";
                        break;
                }
                if (!string.IsNullOrWhiteSpace(condition.worker))
                    sql += @" and V.nmember like '%'+@worker+'%' ";
                #endregion

                var models = DbHelper.Query<dynamic>(OldH2ORepository.ConnectionStringName, sql, condition).Select(m => (IDictionary<string, object>)m);
                MemoryStream ms = new MemoryStream();
                ExcelPackage excel = new ExcelPackage();

                // 頁籤名稱設定
                ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("客服業務報表");
                var rowPos = 0;
                var colPos = 0;
                rowPos++;
                string[] ssArr = new string[] { "序號", "類別", "類型", "照會單號", "來源","承辦人", "受理月", "受理日", "保險公司", "要保人",
                    "被保險人", "生效日", "保單號碼", "案件類別", "照會內容", "副總團隊", "協理體系", "經手人實駐", "經手人處別", "原經手人","服務人員實駐","服務人員處別", "服務人員",
                    "發文保公","保公回覆","回覆單位", "處理天數","保公決議", "處理過程","處理結果" };
     


                foreach (string title in ssArr)
                {
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = title;
                }
                models.ForEach(m =>
                {
                    rowPos++;
                    colPos = 1;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = rowPos - 1;//序號     
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Category"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CODE"));
                    colPos++; ;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("No"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SourceType"));
                    colPos++;
                    //新增承辦人(立案人)
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Creator"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("m"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("d"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CompanyNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Owners")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Insureds")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("IssDates")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("PolicyNos")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CaseType"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Content"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("VMNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SMNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("WCNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CenterNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("AgentNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUWCNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUCenterNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUAgentNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("BUSContactCompanyDates")).Replace(",", "\n");//發文保公
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("ReplyCompanyDates")).Replace(",", "\n");//保公回覆
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("ReplyYallowBillDates")).Replace(",", "\n");//回覆單位
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("DoDay"));//處理天數
                    colPos++;
                    //新增保公決議
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CompanyResult"));//保公決議
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Contents")).Replace("|", "\n");//處理流程
                    //新增處理結果
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Result"));//處理結果

                });

                // 自動調整欄位大小
                for (int i = 1; i <= ssArr.Length; i++)
                {
                    sheet.Column(i).AutoFit();
                    if (sheet.Column(i).Width > 75)
                        sheet.Column(i).Width = 75;
                }

                var cells = sheet.Cells[1, 1, rowPos, ssArr.Length];
                cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cells.Style.WrapText = true;
                excel.SaveAs(ms);
                excel.Dispose();
                ms.Position = 0;
                return ms;
            }
            catch (Exception ex)
            {

            }
            return null;

        }

        /// <summary>
        /// CC報表
        /// </summary>
        public Stream GetCCReport(QueryReportCondition condition)
        {
            try
            {
                #region SQL 語法
                string sql = @" SELECT distinct
              CASE  WHEN (C.CODE='CS' or C.CODE='SS') THEN '客服'  ELSE '申訴' END AS Category, 
              C.CODE,C.No,CDT.Name as SourceType,
             CASE WHEN ISNULL(CCC.DoUser,'')<>'' THEN V1.nmember ELSE V.nmember END AS Creator,
              Month(C.CreateTime)as m,DaY(C.CreateTime) as d, 
			  CONVERT(VARCHAR,CCC.ReceiveDateTime,111) AS ReceiveDateTime,
			  CONVERT(VARCHAR,CCL.CreateTime,111) AS ReplayDDLDateTime,
               ISNULL(STUFF(CompanyNameList.CompanyNames, 1, 1, ''), '') AS CompanyNames, 
              ISNULL(STUFF(OwnerList.Owners, 1, 1, ''), '') AS Owners, 
              ISNULL(STUFF(InsuredList.Insureds, 1, 1, ''), '') AS Insureds, 
              ISNULL(STUFF(IssDateList.IssDates, 1, 1, ''), '') AS IssDates, 
              ISNULL(STUFF(PolicyList.PolicyValue, 1, 1, ''), '') AS PolicyNos, 
              ISNULL(STUFF(ModPremList.ModPremValue, 1, 1, ''), '') AS ModPrems,
               CDT1.Name as CaseType,Content, 
             ISNULL(STUFF(VMNameList.VMNames, 1, 1, ''), '') AS VMNames, 
              ISNULL(STUFF(SMNameList.SMNames, 1, 1, ''), '') AS SMNames, 
               ISNULL(STUFF(WCNameList.WCNames, 1, 1, ''), '') AS WCNames, 
               ISNULL(STUFF(CenterNameList.CenterNames, 1, 1, ''), '') AS CenterNames, 
               ISNULL(STUFF(AgentNameList.AgentNames, 1, 1, ''), '') AS AgentNames, 
               ISNULL(STUFF(SUWCNameList.SUWCNames, 1, 1, ''), '') AS SUWCNames, 
			   ISNULL(STUFF(SUCenterNameList.SUCenterNames, 1, 1, ''), '') AS SUCenterNames, 
               ISNULL(STUFF(SUAgentNameList.SUAgentNames, 1, 1, ''), '') AS SUAgentNames, 
          	   ISNULL(STUFF(CCReplyDateList.CCReplyDates, 1, 1, ''), '') AS CCReplyDates, 
              CASE WHEN (ISNULL(CD.CCReplyDate,'')<>'') THEN DATEDIFF(d,CCC.ReceiveDateTime,CD.CCReplyDate) ELSE DATEDIFF(d,C.CreateTime,getdate()) END AS DoDay,
              ISNULL(STUFF(CRMEDoList.ContentValue, 1, 1, ''), '') AS Contents ,
              ISNULL(CDT2.Name,'') AS Result
              FROM( 
              select * from [H2O].[dbo].[CRMENo] where isnull(No,'')<>'' 
              and CODE=@code ";
                if (!string.IsNullOrWhiteSpace(condition.crm_no))
                    sql += @" and No=@crm_no ";
                if (condition.startdate.HasValue)
                    sql += @" and CreateTime >= @startdate ";
                if (condition.closedate.HasValue)
                    sql += @" and CreateTime <= @closedate +' 23:59:59' ";
                sql += @" ) C 
              LEFT JOIN [H2O].[dbo].[CRMECaseContent] CCC ON C.No=CCC.No 
              LEFT JOIN [H2O].[dbo].[CRMEInsurancePolicy] CIP ON C.No=CIP.No 
              LEFT JOIN [H2O].[dbo].[CRMEDiscipType] CDT ON CCC.SourceID=CDT.ID 
	          LEFT JOIN [H2O].[dbo].[CRMEDiscipType] CDT1 ON CCC.CaseTypeID=CDT1.ID 
              LEFT JOIN [CUF].[dbo].[v_member_account] V ON V.imember=C.Creator
              LEFT JOIN [CUF].[dbo].[v_member_account] V1 ON V1.imember=CCC.DoUser
              LEFT JOIN (SELECT No,ResultCode,Creator,MaxCCL.createtime  FROM CRMECloseLog,(SELECT MAX(createtime) AS createtime FROM CRMECloseLog GROUP BY No) AS MaxCCL WHERE CRMECloseLog.createtime = MaxCCL.createtime)  CCL ON C.No=CCL.No  
              LEFT JOIN [H2O].[dbo].[CRMEDiscipType] CDT2 ON CCL.ResultCode=CDT2.ID 
	           LEFT JOIN (SELECT NO ,MAX(CCReplyDate)  AS CCReplyDate FROM [H2O].[dbo].CRMEDo GROUP BY No)AS CD ON C.No=CD.No 
              OUTER APPLY (
                SELECT '|' + CD.Content 
                FROM dbo.CRMEDo CD
                WHERE C.No = CD.No
                FOR XML PATH ('')
            ) CRMEDoList (ContentValue)
            OUTER APPLY (
                SELECT distinct ',' + CIP.CompanyName 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) CompanyNameList (CompanyNames)
            OUTER APPLY (
                SELECT distinct ',' + CIP.Owner 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) OwnerList (Owners)
            OUTER APPLY (
                SELECT distinct ',' + CIP.Insured
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) InsuredList (Insureds)
            OUTER APPLY (
                SELECT distinct ',' + CIP.IssDate
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) IssDateList (IssDates)
             OUTER APPLY (
                SELECT distinct ',' + CIP.PolicyNo 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) PolicyList (PolicyValue)
             OUTER APPLY (
                SELECT distinct ',' +  convert(varchar,CIP.ModPrem) 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) ModPremList (ModPremValue)
             OUTER APPLY (
                SELECT	distinct ',' + CIP.VMName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) VMNameList (VMNames)
             OUTER APPLY (
                SELECT distinct ',' + CIP.SMName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SMNameList (SMNames)
             OUTER APPLY (
                SELECT	distinct ',' + CIP.WCName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) WCNameList (WCNames)
             OUTER APPLY (
                SELECT	distinct ',' + CIP.centerName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) CenterNameList (CenterNames)
            OUTER APPLY (
                SELECT 	distinct ',' + CIP.AgentName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) AgentNameList (AgentNames)
            OUTER APPLY (
                SELECT	distinct ',' + CIP.SUWCName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUWCNameList (SUWCNames)
			  OUTER APPLY (
                SELECT	distinct ',' + CIP.SUCenterName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUCenterNameList (SUCenterNames)
            OUTER APPLY (
                SELECT	distinct ',' + CIP.SUAgentName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUAgentNameList (SUAgentNames)
			OUTER APPLY (
                SELECT	distinct ',' +  CONVERT(VARCHAR,CD.CCReplyDate,111)
                FROM dbo.[CRMEDo] CD
                WHERE C.No = CD.No
                FOR XML PATH ('')
            ) CCReplyDateList (CCReplyDates)
            where 1=1 ";
                if (!string.IsNullOrWhiteSpace(condition.PolicyNo))
                    sql += @" and PolicyNo like '%'+@PolicyNo+'%' ";
                if (!string.IsNullOrWhiteSpace(condition.owner))
                    sql += @" and owner like '%'+@owner+'%' ";
                switch (condition.close_status) {
                    case 0:
                        sql += @" and isnull( CCL.CreateTime,'')='' ";
                        break;
                    case 1:
                        sql += @" and isnull( CCL.CreateTime,'')<>'' ";
                        break;
                }
                if (!string.IsNullOrWhiteSpace(condition.worker))
                    sql += @" and V1.nmember like '%'+@worker+'%' OR (ISNULL(CCC.DoUser,'')='' AND V.nmember like '%'+@worker+'%') " ;

                #endregion

                var models = DbHelper.Query<dynamic>(OldH2ORepository.ConnectionStringName, sql, condition).Select(m => (IDictionary<string, object>)m);
                MemoryStream ms = new MemoryStream();
                ExcelPackage excel = new ExcelPackage();

                // 頁籤名稱設定
                ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("客服業務報表");
                var rowPos = 0;
                var colPos = 0;
                rowPos++;

                string[] ccArr = new string[] { "序號", "類別", "類型", "照會單號", "來源","承辦人","受理月", "受理日","移交日","回文期限","保險公司", "要保人",
                    "被保險人", "生效日", "保單號碼","保費", "案件類別", "申訴內容", "副總團隊", "協理體系", "經手人實駐", "經手人處別", "原經手人","服務人員實駐","服務人員處別", "服務人員",
                    "回文日", "處理天數", "處理過程","處理結果" };



                foreach (string title in ccArr)
                {
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = title;
                }
                models.ForEach(m =>
                {
                    rowPos++;
                    colPos = 1;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = rowPos - 1;//序號     
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Category"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CODE"));
                    colPos++; ;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("No"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SourceType"));
                    colPos++;
                    //新增承辦人
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Creator"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("m"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("d"));
                    //新增移交日
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("ReceiveDateTime"));
                    //新增回文期限
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("ReplayDDLDateTime"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CompanyNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Owners")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Insureds")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("IssDates")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("PolicyNos")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("ModPrems")).Replace(",", "\n");//保費
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CaseType"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Content"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("VMNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SMNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("WCNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CenterNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("AgentNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUWCNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUCenterNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUAgentNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CCReplyDates")).Replace(",", "\n");//回文日
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("DoDay"));//處理天數
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Contents")).Replace("|", "\n");//處理流程
                    //新增處理結果
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Result"));//處理流程
                });

                // 自動調整欄位大小
                for (int i = 1; i <= ccArr.Length; i++)
                {
                    sheet.Column(i).AutoFit();
                    if (sheet.Column(i).Width > 75)
                        sheet.Column(i).Width = 75;
                }

                var cells = sheet.Cells[1, 1, rowPos, ccArr.Length];
                cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cells.Style.WrapText = true;
                excel.SaveAs(ms);
                excel.Dispose();
                ms.Position = 0;
                return ms;
            }
            catch (Exception ex)
            {
                Throw.BusinessError("Service Error");
            }
            return null;

        }

        /// <summary>
        /// SC報表
        /// </summary>
        public Stream GetSCReport(QueryReportCondition condition)
        {
            try
            {
                #region SQL 語法
                string sql = @" SELECT distinct
              CASE  WHEN (C.CODE='CS' or C.CODE='SS') THEN '客服'  ELSE '申訴' END AS Category, 
              C.CODE,C.No,CDT.Name as SourceType,
              CASE WHEN ISNULL(CCC.DoUser,'')<>'' THEN V1.nmember ELSE V.nmember END AS Creator,
              Month(C.CreateTime)as m,DaY(C.CreateTime) as d, 
               ISNULL(STUFF(CompanyNameList.CompanyNames, 1, 1, ''), '') AS CompanyNames, 
              ISNULL(STUFF(OwnerList.Owners, 1, 1, ''), '') AS Owners, 
              ISNULL(STUFF(InsuredList.Insureds, 1, 1, ''), '') AS Insureds, 
              ISNULL(STUFF(IssDateList.IssDates, 1, 1, ''), '') AS IssDates, 
              ISNULL(STUFF(PolicyList.PolicyValue, 1, 1, ''), '') AS PolicyNos, 
              ISNULL(STUFF(ModPremList.ModPremValue, 1, 1, ''), '') AS ModPrems,
               CDT1.Name as CaseType,Content, 
             ISNULL(STUFF(VMNameList.VMNames, 1, 1, ''), '') AS VMNames, 
              ISNULL(STUFF(SMNameList.SMNames, 1, 1, ''), '') AS SMNames, 
               ISNULL(STUFF(WCNameList.WCNames, 1, 1, ''), '') AS WCNames, 
               ISNULL(STUFF(CenterNameList.CenterNames, 1, 1, ''), '') AS CenterNames, 
               ISNULL(STUFF(AgentNameList.AgentNames, 1, 1, ''), '') AS AgentNames,
               ISNULL(STUFF(SUWCNameList.SUWCNames, 1, 1, ''), '') AS SUWCNames, 
			   ISNULL(STUFF(SUCenterNameList.SUCenterNames, 1, 1, ''), '') AS SUCenterNames, 
               ISNULL(STUFF(SUAgentNameList.SUAgentNames, 1, 1, ''), '') AS SUAgentNames, 
          	   CCL.CreateTime AS CloseDay, 
              CASE WHEN (ISNULL(CCL.CreateTime,'')<>'') THEN DATEDIFF(d,CCC.CreateTime,CCL.CreateTime) ELSE DATEDIFF(d,C.CreateTime,getdate()) END AS DoDay,
              ISNULL(STUFF(CRMEDoList.ContentValue, 1, 1, ''), '') AS Contents 
              FROM( 
              select * from [H2O].[dbo].[CRMENo] where isnull(No,'')<>'' 
              and CODE=@code ";
                if (!string.IsNullOrWhiteSpace(condition.crm_no))
                    sql += @" and No=@crm_no ";
                if (condition.startdate.HasValue)
                    sql += @" and CreateTime >= @startdate ";
                if (condition.closedate.HasValue)
                    sql += @" and CreateTime <= @closedate +' 23:59:59' ";
                sql += @" ) C 
              LEFT JOIN [H2O].[dbo].[CRMECaseContent] CCC ON C.No=CCC.No 
              LEFT JOIN [H2O].[dbo].[CRMEInsurancePolicy] CIP ON C.No=CIP.No 
              LEFT JOIN [H2O].[dbo].[CRMEDiscipType] CDT ON CCC.SourceID=CDT.ID 
	          LEFT JOIN [H2O].[dbo].[CRMEDiscipType] CDT1 ON CCC.CaseTypeID=CDT1.ID 
              LEFT JOIN [CUF].[dbo].[v_member_account] V ON V.imember=C.Creator
              LEFT JOIN [CUF].[dbo].[v_member_account] V1 ON V1.imember=CCC.DoUser
              LEFT JOIN (SELECT No,ResultCode,Creator,MaxCCL.createtime  FROM CRMECloseLog,(SELECT MAX(createtime) AS createtime FROM CRMECloseLog GROUP BY No) AS MaxCCL WHERE CRMECloseLog.createtime = MaxCCL.createtime)  CCL ON C.No=CCL.No  
              OUTER APPLY (
                SELECT '|' + CD.Content 
                FROM dbo.CRMEDo CD
                WHERE C.No = CD.No
                FOR XML PATH ('')
            ) CRMEDoList (ContentValue)
            OUTER APPLY (
                SELECT distinct ',' + CIP.CompanyName 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) CompanyNameList (CompanyNames)
            OUTER APPLY (
                SELECT distinct ',' + CIP.Owner 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) OwnerList (Owners)
            OUTER APPLY (
                SELECT distinct ',' + CIP.Insured
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) InsuredList (Insureds)
            OUTER APPLY (
                SELECT distinct ',' + CIP.IssDate
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) IssDateList (IssDates)
             OUTER APPLY (
                SELECT distinct ',' + CIP.PolicyNo 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) PolicyList (PolicyValue)
             OUTER APPLY (
                SELECT distinct ',' +  convert(varchar,CIP.ModPrem) 
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) ModPremList (ModPremValue)
             OUTER APPLY (
                SELECT	distinct ',' + CIP.VMName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) VMNameList (VMNames)
             OUTER APPLY (
                SELECT distinct ',' + CIP.SMName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SMNameList (SMNames)
             OUTER APPLY (
                SELECT	distinct ',' + CIP.WCName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) WCNameList (WCNames)
             OUTER APPLY (
                SELECT	distinct ',' + CIP.centerName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) CenterNameList (CenterNames)
            OUTER APPLY (
                SELECT 	distinct ',' + CIP.AgentName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) AgentNameList (AgentNames)
            OUTER APPLY (
                SELECT	distinct ',' + CIP.SUWCName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUWCNameList (SUWCNames)
			  OUTER APPLY (
                SELECT	distinct ',' + CIP.SUCenterName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUCenterNameList (SUCenterNames)
            OUTER APPLY (
                SELECT	distinct ',' + CIP.SUAgentName
                FROM dbo.[CRMEInsurancePolicy] CIP
                WHERE C.No = CIP.No
                FOR XML PATH ('')
            ) SUAgentNameList (SUAgentNames)
            where 1=1 ";
                if (!string.IsNullOrWhiteSpace(condition.PolicyNo))
                    sql += @" and PolicyNo like '%'+@PolicyNo+'%' ";
                if (!string.IsNullOrWhiteSpace(condition.owner))
                    sql += @" and owner like '%'+@owner+'%' ";
                switch (condition.close_status)
                {
                    case 0:
                        sql += @" and isnull( CCL.CreateTime,'')='' ";
                        break;
                    case 1:
                        sql += @" and isnull( CCL.CreateTime,'')<>'' ";
                        break;
                }
                if (!string.IsNullOrWhiteSpace(condition.worker))
                    sql += @" and V1.nmember like '%'+@worker+'%' OR (ISNULL(CCC.DoUser,'')='' AND V.nmember like '%'+@worker+'%') ";


                #endregion

                var models = DbHelper.Query<dynamic>(OldH2ORepository.ConnectionStringName, sql, condition).Select(m => (IDictionary<string, object>)m);
                MemoryStream ms = new MemoryStream();
                ExcelPackage excel = new ExcelPackage();

                // 頁籤名稱設定
                ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("客服業務報表");
                var rowPos = 0;
                var colPos = 0;
                rowPos++;

                string[] scArr = new string[] { "序號", "類別", "類型", "照會單號", "來源","承辦人", "受理月", "受理日", "保險公司", "要保人",
                    "被保險人", "生效日", "保單號碼","保費", "案件類別", "申訴內容", "副總團隊", "協理體系", "經手人實駐", "經手人處別", "原經手人", "服務人員實駐", "服務人員處別", "服務人員",
                    "結案日", "處理天數", "處理過程" };
                foreach (string title in scArr)
                {
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = title;
                }
                models.ForEach(m =>
                {
                    rowPos++;
                    colPos = 1;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = rowPos - 1;//序號     
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Category"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CODE"));
                    colPos++; ;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("No"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SourceType"));
                    colPos++;
                    //新增承辦人
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Creator"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("m"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("d"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CompanyNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Owners")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Insureds")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("IssDates")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("PolicyNos")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("ModPrems")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CaseType"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Content"));
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("VMNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SMNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("WCNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CenterNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("AgentNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUWCNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUCenterNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("SUAgentNames")).Replace(",", "\n");
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("CloseDay"));//結案日
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToInt32(m.GetOrDefault("DoDay"));//處理天數
                    colPos++;
                    sheet.Cells[rowPos, colPos, rowPos, colPos].Value = Convert.ToString(m.GetOrDefault("Contents")).Replace("|", "\n");//處理流程
                });

                // 自動調整欄位大小
                for (int i = 1; i <= scArr.Length; i++)
                {
                    sheet.Column(i).AutoFit();
                    if (sheet.Column(i).Width > 75)
                        sheet.Column(i).Width = 75;
                }

                var cells = sheet.Cells[1, 1, rowPos, scArr.Length];
                cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cells.Style.WrapText = true;
                excel.SaveAs(ms);
                excel.Dispose();
                ms.Position = 0;
                return ms;
            }
            catch (Exception ex)
            {

            }
            return null;

        }
       




        #endregion
    }
}
