using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EP.Platform.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework;
using Microsoft.CUF.Framework.Service;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    /// <summary>
    /// 立案相關服務
    /// </summary>
    public class CaseService : ICaseService
    {
        private static object caseNo = new object();

        /// <summary>
        /// 依要保人ID取得保單資料
        /// </summary>
        /// <param name="ownerID">要保人ID</param>
        /// <returns>要保人底下的所有產、壽保單資料清單</returns>
        public IEnumerable<CRMEInsurancePolicy> GetInsPolicyByOwnerID(string ownerID)
        {

            List<CRMEInsurancePolicy> result = new List<CRMEInsurancePolicy>();
            var condition = new { ownerID };

            // 先查壽險
            var sql = @"
                Select po_no2 PolicyNo
                    , Co_PolicyNo CoPolicyNo
                    , company_code CompanyCode
                    , company_name CompanyName
                    , modprem ModPrem
                    , status Status
                    , status_name StatusName
                    , modx Modx
                    , modx_name ModxName
                    , owner_id OwnerID
                    , owner Owner
                    , owner_mobile OwnerMobile
                    , insured_id InsuredID
                    , insured Insured
                    , issdate IssDate
                    , a_vm_code VMCode
                    , a_vm_name VMName
                    , a_sm_code SMCode
                    , a_sm_name SMName
                    , a_center_code CenterCode
                    , a_center_name CenterName
                    , a_wc_code WCCode
                    , a_wc_name WCName
                    , agent_code AgentCode
                    , agent_name AgentName
                    , a_vm_leader_id VMLeaderID
                    , a_vm_leader VMLeader
                    , a_sm_leader_id SMLeaderID
                    , a_sm_leader SMLeader
                    , a_center_leader_id CenterLeaderID
                    , a_center_leader CenterLeader
                    , su_vm_code SUVMCode
                    , su_vm_name SUVMName
                    , su_sm_code SUSMCode
                    , su_sm_name SUSMName
                    , su_center_code SUCenterCode
                    , su_center_name SUCenterName
                    , su_wc_code SUWCCode
                    , su_wc_name SUWCName
                    , sudirector_id SUAgentCode
                    , su_name SUAgentName
                    , su_vm_leader_id SUVMLeaderID
                    , su_vm_leader SUVMLeader
                    , su_sm_leader_id SUSMLeaderID
                    , su_sm_leader SUSMLeader
                    , su_center_leader_id SUCenterLeaderID
                    , su_center_leader SUCenterLeader
                    , status_type StatusType
                    , money_type MoneyType
                 From v_pbd_potrace_list
             Where owner_id = @ownerID
                Order by company_code, insured, po_no2
            ";
            var potraces =  DbHelper.Query<CRMEInsurancePolicy>(CRSRepository.ConnectionStringName, sql, condition);

            if (potraces.Any())
            {
                result.AddRange(potraces);
            }

            // 再查產險
            sql = @"
                Select po_no2 PolicyNo
                    , '' CoPolicyNo
                    , company_code CompanyCode
                    , company_name CompanyName
                    , gross_premium ModPrem
                    , '' Status
                    , status_name StatusName
                    , '' Modx
                    , '' ModxName
                    , owner_id OwnerID
                    , owner Owner
                    , owner_mobile OwnerMobile
                    , insured_id InsuredID
                    , insured Insured
                    , issdate IssDate
                    , a_vm_code VMCode
                    , a_vm_name VMName
                    , a_sm_code SMCode
                    , a_sm_name SMName
                    , a_center_code CenterCode
                    , a_center_name CenterName
                    , a_wc_code WCCode
                    , a_wc_name WCName
                    , agent_code AgentCode
                    , agent_name AgentName
                    , a_vm_leader_id VMLeaderID
                    , a_vm_leader VMLeader
                    , a_sm_leader_id SMLeaderID
                    , a_sm_leader SMLeader
                    , a_center_leader_id CenterLeaderID
                    , a_center_leader CenterLeader
                    , su_vm_code SUVMCode
                    , su_vm_name SUVMName
                    , su_sm_code SUSMCode
                    , su_sm_name SUSMName
                    , su_center_code SUCenterCode
                    , su_center_name SUCenterName
                    , su_wc_code SUWCCode
                    , su_wc_name SUWCName
                    , sudirector_id SUAgentCode
                    , su_name SUAgentName
                    , su_vm_leader_id SUVMLeaderID
                    , su_vm_leader SUVMLeader
                    , su_sm_leader_id SUSMLeaderID
                    , su_sm_leader SUSMLeader
                    , su_center_leader_id SUCenterLeaderID
                    , su_center_leader SUCenterLeader
                    , status_type StatusType
                    , '' MoneyType
                 From potrace_ins
             Where owner_id = @ownerID
                Order by company_code, insured, po_no2
            ";

            potraces = DbHelper.Query<CRMEInsurancePolicy>(CRSRepository.ConnectionStringName, sql, condition);

            if (potraces.Any())
            {
                result.AddRange(potraces);
            }

            // 回傳要保人的壽產險的保單清單資料
            return result;
        }


        /// <summary>
        /// 檢核保單號碼的狀態
        /// </summary>
        /// <param name="policyNo">保單號碼</param>
        /// <returns>
        /// 1. false:接續人為公司 true:有相同的保單號碼或是查無保單
        /// 2. 保單號碼的資料，如果1為true時，且沒有資料，代表查無保單
        /// </returns>
        public Tuple<bool, IEnumerable<CRMEInsurancePolicy>> CheckPolicyNo(string policyNo)
        {
            var check = true;
            List<CRMEInsurancePolicy> result = new List<CRMEInsurancePolicy>();
            var condition = new { policyNo };

            // 先查壽險
            var sql = @"
                Select po_no2 PolicyNo
                    , Co_PolicyNo CoPolicyNo
                    , company_code CompanyCode
                    , company_name CompanyName
                    , owner_id OwnerID
                    , owner Owner
                    , insured_id InsuredID
                    , insured Insured
                    , agent_code AgentCode
                    , sudirector_id SUAgentCode
                 From v_pbd_potrace_list
             Where po_no2 = @policyNo
                Group by po_no2, Co_PolicyNo, company_code, company_name, owner_id, owner, insured_id, insured, agent_code, sudirector_id
                Order by company_code, owner_id
            ";
            var potraces = DbHelper.Query<CRMEInsurancePolicy>(CRSRepository.ConnectionStringName, sql, condition);

            if (potraces.Any())
            {
                result.AddRange(potraces);
            }

            // 再查產險
            sql = @"
                Select po_no2 PolicyNo
                    , '' CoPolicyNo
                    , company_code CompanyCode
                    , company_name CompanyName
                    , owner_id OwnerID
                    , owner Owner
                    , insured_id InsuredID
                    , insured Insured
                    , agent_code AgentCode
                    , sudirector_id SUAgentCode
                 From potrace_ins
             Where po_no2 = @policyNo
                Group by po_no2, company_code, company_name, owner_id, owner, insured_id, insured, agent_code, sudirector_id
                Order by company_code, owner_id
            ";

            potraces = DbHelper.Query<CRMEInsurancePolicy>(CRSRepository.ConnectionStringName, sql, condition);

            if (potraces.Any())
            {
                result.AddRange(potraces);
            }

            if (result.Count() > 0 && result.Count() == result.Where(m => (string.IsNullOrWhiteSpace(m.SUAgentCode) && "Z99999999901".Equals(m.AgentCode ?? "", StringComparison.OrdinalIgnoreCase)) || "Z99999999901".Equals(m.SUAgentCode ?? "", StringComparison.OrdinalIgnoreCase)).Count())
            {
                check = false;
            }


            result = result.GroupBy(m => new CRMEInsurancePolicy { PolicyNo =  m.PolicyNo, CompanyCode = m.CompanyCode, CompanyName = m.CompanyName, OwnerID = m.OwnerID, Owner = m.Owner, InsuredID = m.InsuredID, Insured = m.Insured }).GroupBy(x => x.Key, new CRMEInsurancePolicyComparer()).Select(m => m.Key).ToList();

            return Tuple.Create(check, result?.AsEnumerable()); ;
        }

        /// <summary>
        /// 依服務申訴類型取得資料來源
        /// </summary>
        /// <param name="type">服務申訴類型</param>
        /// <returns>資料來源清單</returns>
        public IEnumerable<CRMEDiscipType> GetSourceListByType(DiscipTypeCode type)
        {
            var condition = new QueryDiscipTypeCondition() { Code = type, Kind = DiscipTypeKind.Source, Status = EnableStatus.Enabled };
            return new CommonService().QueryCRMEDiscipTypeDatas(condition);
        }

        /// <summary>
        /// 依服務申訴類型取得來電者
        /// </summary>
        /// <param name="type">服務申訴類型</param>
        /// <returns>來電者清單</returns>
        public IEnumerable<CRMEDiscipType> GetCallerListByType(DiscipTypeCode type)
        {
            var condition = new QueryDiscipTypeCondition() { Code = type, Kind = DiscipTypeKind.Caller, Status = EnableStatus.Enabled };
            return new CommonService().QueryCRMEDiscipTypeDatas(condition);
        }

        /// <summary>
        /// 依服務申訴類型取得案件類別
        /// </summary>
        /// <param name="type">服務申訴類型</param>
        /// <returns>案件類別清單</returns>
        public IEnumerable<CRMEDiscipType> GetCaseTypeListByType(DiscipTypeCode type)
        {
            var condition = new QueryDiscipTypeCondition() { Code = type, Kind = DiscipTypeKind.CaseType, Status = EnableStatus.Enabled };
            return new CommonService().QueryCRMEDiscipTypeDatas(condition);
        }

        /// <summary>
        /// 依服務申訴類型取得案件類型
        /// </summary>
        /// <param name="type">服務申訴類型</param>
        /// <returns>案件類型清單</returns>
        public IEnumerable<CRMEDiscipType> GetCaseCategoryListByType(DiscipTypeCode type)
        {
            var condition = new QueryDiscipTypeCondition() { Code = type, Kind = DiscipTypeKind.CaseCategory, Status = EnableStatus.Enabled };
            return new CommonService().QueryCRMEDiscipTypeDatas(condition);
        }

        /// <summary>
        /// 取得保險司清單(先壽後產)
        /// </summary>
        /// <returns>保險公司清單</returns>
        public IEnumerable<ValueText> GetInsCompanyList()
        {
            var result = new List<ValueText>();
            // 壽險公司清單
            var sql = @"
                    Select RTRIM(term_code) as Value,
                        RTRIM(term_meaning) as Text
                    From trmval
                        Where term_id = 'company_code'
                        And TRY_PARSE(term_code as int USING 'en-US') is not null
                        Order by term_sequence";
            var companys =  DbHelper.Query<ValueText>(MISRepository.ConnectionStringName, sql);
            if (companys.Any())
            {
                result.AddRange(companys);
            }

            // 產險公司清單
            sql = @"
                    Select RTRIM(term_code) as Value,
                        RTRIM(term_meaning) as Text
                    From trmval
                        Where term_id = 'company_code'
                        And TRY_PARSE(term_code as int USING 'en-US') is null
                        Order by term_sequence";
            companys = DbHelper.Query<ValueText>(MISRepository.ConnectionStringName, sql);
            if (companys.Any())
            {
                result.AddRange(companys);
            }

            return result;
        }

        /// <summary>
        /// 取得要保人手號碼
        /// </summary>
        /// <param name="ownerID">要保人ID</param>
        /// <returns>要保人手機號碼</returns>
        public string GetOwnerMobile(string ownerID)
        {
            var tempDatas = GetInsPolicyByOwnerID(ownerID);

            var result = tempDatas.Where(m => m.StatusType == "有效件" && !string.IsNullOrWhiteSpace(m.OwnerMobile)).OrderByDescending(m => m.IssDate).FirstOrDefault();

            if (result == null)
            {
                result = tempDatas.Where(m => m.StatusType != "有效件" && !string.IsNullOrWhiteSpace(m.OwnerMobile)).OrderByDescending(m => m.IssDate).FirstOrDefault();
            }

            return result?.OwnerMobile;
        }

        /// <summary>
        /// 新增立案
        /// </summary>
        /// <param name="content">立案的資料</param>
        /// <param name="policys">保單資料清單</param>
        /// <returns>立案後的受理編號</returns>
        public string CreateCaseContent(CRMECaseContent content, List<CRMEInsurancePolicy> policys)
        {

            // 取得受理編號
            content.No = CreateCRMENo(content.Type, content.Creator);
            // 填入Guid
            content.CaseGuid = Guid.NewGuid();

            // 保單資料填入對應的受理編號、建立時間和建立人員,要被保人姓名去除開頭和字尾空白和全型空白
            policys.ForEach(m =>
            {
                m.No = content.No;
                m.Owner = m.Owner.TrimEnd(new char[] { '　', ' ' }).TrimStart(new char[] { '　', ' ' });
                m.Insured = m.Insured.TrimEnd(new char[] { '　', ' ' }).TrimStart(new char[] { '　', ' ' });
                m.Creator = content.Creator;
                m.CreateTime = content.CreateTime;
            });

            policys = policys.GroupBy(x => x, new CRMEInsurancePolicyComparer()).Select(x => x.Key).ToList();

            // 新增的處理
            using (var ts = new TransactionScope())
            {
                content.Insert(new[] { H2ORepository.ConnectionStringName });
                policys.Inserts(new[] { H2ORepository.ConnectionStringName });
                ts.Complete();
            }

            // 回傳受理編號
            return content.No; 
        }

        /// <summary>
        /// 取得等待通知的案件資料清單
        /// </summary>
        /// <returns>等待通知的案件資料清單</returns>
        public IEnumerable<CRMECaseContent> GetWaitNofityDatas()
        {
            ///var result =  H2ORepository.Select<CRMECaseContent>(new { Status = 0 });
            var sql = @"
                       select ID,No,CaseGuid,Type ,SourceID ,SourceDesc ,CompanyCode,CallerID
                        ,CaseTypeID,CaseCategoryID,ReceiveDateTime,ReplayDDLDateTime,BizContract
                        ,SolicitingRpt,Complainant,Content,Status,DoUser,DoUserTelExt,ProcResultRadio
                        ,ProcResultCheckBox,OnlineReplay,Creator,CreateTime 
                        from CRMECaseContent where [Status] = 0 and [No] Not in(select No from CRMECloseLog group by No)
                        and [Type] <>'CC' 
                        union all
                        select ID,No,CaseGuid,Type ,SourceID ,SourceDesc ,CompanyCode,CallerID
                        ,CaseTypeID,CaseCategoryID,ReceiveDateTime,ReplayDDLDateTime,BizContract
                        ,SolicitingRpt,Complainant,Content,Status,DoUser,DoUserTelExt,ProcResultRadio
                        ,ProcResultCheckBox,OnlineReplay,Creator,CreateTime 
                        from CRMECaseContent where [Status] = 0 and [No] Not in(select No from CRMECloseLog group by No)
                        and [Type] ='CC' 
                        and [No] in(select [No] from CRMEAppealBy) ";

            var result = H2ORepository.Select<CRMECaseContent>(sql, new { });

            if (result != null)
            {
                result = result.ToList();
                var commonService = new CommonService();
                var types = commonService.GetCaseType(null).ToDictionary(m => m.Code.GetValue(), m => m.Name);
                var discipTypes = commonService.QueryCRMEDiscipTypeDatas(null).ToDictionary(m => m.ID, m => m.Name);
                result.ForEach(m =>
                {
                    m.CreatorName = Member.Get(m.Creator)?.Name;
                    if (!string.IsNullOrEmpty(m.Type))
                    {
                        m.TypeName = types[m.Type];
                    }

                    if (m.SourceID.HasValue)
                    {
                        m.SourceName = discipTypes[m.SourceID.Value];
                    }

                    if (m.CallerID.HasValue)
                    {
                        m.CallerName = discipTypes[m.CallerID.Value];
                    }

                    if (m.CaseCategoryID.HasValue)
                    {
                        m.CaseCategoryName = discipTypes[m.CaseCategoryID.Value];
                    }

                    if (m.CaseTypeID.HasValue)
                    {
                        m.CaseTypeName = discipTypes[m.CaseTypeID.Value];
                    }
                });

               
            }

            return result;
        }



        private string GetVarConfig(string varName)
        {
            var service = new DataSettingService();
            var config = service.GetVarConfigDetailModelByName(varName);
            return config.VarValue;
        }

        /// <summary>
        /// 取得等待通知申訴人的案件資料清單
        /// </summary>
        /// <returns>等待通知申訴人的案件資料清單</returns>
        public IEnumerable<CRMECaseContent> GetCCWaitNofityDatas()
        {
            ///var result =  H2ORepository.Select<CRMECaseContent>(new { Status = 0 });
            var DateStart = GetVarConfig("AppealBYDateStart");
            var sql = @"
                        select ID,No,CaseGuid,Type ,SourceID ,SourceDesc ,CompanyCode,CallerID
                        ,CaseTypeID,CaseCategoryID,ReceiveDateTime,ReplayDDLDateTime,BizContract
                        ,SolicitingRpt,Complainant,Content,Status,DoUser,DoUserTelExt,ProcResultRadio
                        ,ProcResultCheckBox,OnlineReplay,Creator,CreateTime 
                        from CRMECaseContent where [Type] ='CC'  
                        and [No] Not in(select No from CRMECloseLog where right(left(NO,5),2)='CC' group by No)
                        and [No] not in(select [No] from CRMEAppealBy)
                        and convert(nvarchar,CreateTime,111) > @DateStart ";

            var result = H2ORepository.Select<CRMECaseContent>(sql, new { DateStart });

            if (result != null)
            {
                result = result.ToList();
                var commonService = new CommonService();
                var types = commonService.GetCaseType(null).ToDictionary(m => m.Code.GetValue(), m => m.Name);
                var discipTypes = commonService.QueryCRMEDiscipTypeDatas(null).ToDictionary(m => m.ID, m => m.Name);
                result.ForEach(m =>
                {
                    m.CreatorName = Member.Get(m.Creator)?.Name;
                    if (!string.IsNullOrEmpty(m.Type))
                    {
                        m.TypeName = types[m.Type];
                    }

                    if (m.SourceID.HasValue)
                    {
                        m.SourceName = discipTypes[m.SourceID.Value];
                    }

                    if (m.CallerID.HasValue)
                    {
                        m.CallerName = discipTypes[m.CallerID.Value];
                    }

                    if (m.CaseCategoryID.HasValue)
                    {
                        m.CaseCategoryName = discipTypes[m.CaseCategoryID.Value];
                    }

                    if (m.CaseTypeID.HasValue)
                    {
                        m.CaseTypeName = discipTypes[m.CaseTypeID.Value];
                    }
                });


            }

            return result;
        }

        /// <summary>
        /// 檢核是否有未結案的保單號碼
        /// </summary>
        /// <param name="policyNo">保單號碼</param>
        /// <returns>true:有未結案 false:都結案</returns>
        public bool CheckNotClosedPolicyNo(string policyNo)
        {
            var sql = @"Select  a.* 
            From CRMECaseContent as a 
		            INNER JOIN CRMEInsurancePolicy as c on a.No = c.No
                Where 
                 NOT EXISTS (
		                Select Top 1 *
			                From CRMECloseLog as b 
			                Where a.No = b.No)
                And PolicyNo = @PolicyNo";

            var models = H2ORepository.Select<CRMECaseContent>(sql, new { PolicyNo = policyNo });

            return models.Any();
        }

        /// <summary>
        /// 取得業務員資訊
        /// </summary>
        /// <param name="agentCode">業務員代碼</param>
        /// <returns>業務員資訊</returns>
        public ExtraData GetAgentInfo(string agentCode)
        {
            var sql = @"SELECT 
                            * FROM orgin_full
                        WHERE agent_code_vlife = @AgentCode";
            var result = DbHelper.Query<dynamic>(MISRepository.ConnectionStringName, sql, new { AgentCode = agentCode }).Select( m => {
                var extraData = new ExtraData();
                ((IDictionary<string, object>)m).ForEach(item =>
                {
                    extraData.Add(item.Key, item.Value != null ? Convert.ToString(item.Value) : "");
                });
                return extraData;
            }).FirstOrDefault();
            
            return result;
        }

        /// <summary>
        /// 新增立案的受理編號
        /// </summary>
        /// <param name="type">客服申訴類型</param>
        /// <param name="creator">建立人員id</param>
        /// <returns>受理編號</returns>
        private string CreateCRMENo(string type, string creator)
        {
            var result = string.Empty;
            var no = DateTime.Now.GetString("CY") + type;
            var sql = @"Select max(No)
                            From CRMENo
                        Where Left(No, 5) = @no";
            lock (caseNo)
            {
                var maxNo = H2ORepository.Query<string>(sql, new { no }).FirstOrDefault();
                if (!string.IsNullOrEmpty(maxNo))
                {
                    result = no + Convert.ToString(Convert.ToInt32(maxNo.Substring(5, 4)) + 1).PadLeft(4, '0');
                }
                else
                {
                    result = no + "0001";
                }
                var model = new CRMENo() { Code = type, No = result, CreateTime = DateTime.Now, Creator = creator };
                model.Insert(new[] { H2ORepository.ConnectionStringName });
            }

            return result;
            
        }

    }
}
