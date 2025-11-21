using EP.H2OModels;
using EP.Platform.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Data;
using System.IO;
using EP.CUFModels;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    public class CUSCRMTX003Service : ICUSCRMTX003Service
    {
        /// <summary>
        /// 查詢待維護資料
        /// </summary>
        /// <param name="condition">查詢條件</param>
        /// <returns></returns>
        public List<MaintainInfo> GetMtnData(QueryMaintainCondition condition)
        {
            var sql = @"select distinct 0 as Serial
                            , c.[type] as Type
                            , (select [name] from CRMEDiscipType where status = 1 and kind = 1 and code = c.Type) as Kind
                            , c.[no] as No
                            , c.CreateTime
                            , p.Owner
                            , (case when isnull(p.SUAgentCode,'') = '' then p.AgentName else p.SUAgentName end) as AgentName 
                            , (select count(no) from CRMEAudit where no = c.no) as AuditCount
		                    , (select count(no) from CRMEDoS where no = c.no) as DosCount
							, max(dos.CreateTime) as dosTime
							, max(au.CreateTime) as auTime
							, max(n.CreateTime) as nTime
							, (case when c.No in (select l.no from CRMECloseLog l group by l.no) then '結案' else 
								case when  max(dos.CreateTime) is not null and (max(isnull(dos.CreateTime,'')) > max(isnull(au.CreateTime,''))) and (max(isnull(dos.CreateTime,'')) > max(isnull(n.CreateTime,''))) then '催辦中' else
									case when  max(au.CreateTime) is not null and (max(isnull(au.CreateTime,'')) > max(isnull(dos.CreateTime,''))) and (max(isnull(au.CreateTime,'')) > max(isnull(n.CreateTime,''))) then '稽催中' else 
										case when  max(n.CreateTime) is not null and (max(isnull(n.CreateTime,'')) > max(isnull(dos.CreateTime,''))) and (max(isnull(n.CreateTime,'')) > max(isnull(au.CreateTime,''))) then '照會中' else 
								'立案不通知' end end end end) as [Status]
                        from CRMECaseContent c 
                        inner join CRMEInsurancePolicy p on p.No = c.No 
						left join CRMENotifyTo n on n.No = c.No
						left join CRMEAudit au on au.No = c.No
						left join CRMEDoS dos on dos.No = c.No
                        where  1=1 
                        ";
            string filter = "";
            if (condition != null) 
            {
                if (!string.IsNullOrWhiteSpace(condition.No))
                {
                    filter += " and c.No = @No ";
                }
                if (!string.IsNullOrWhiteSpace(condition.Kind))
                {
                    filter += " and c.Type = @Kind ";
                }
                if (!string.IsNullOrWhiteSpace(condition.DateS))
                {
                    filter += " and c.CreateTime >= @DateS ";
                }
                if (!string.IsNullOrWhiteSpace(condition.DateE))
                {
                    filter += " and c.CreateTime <= @DateE ";
                }
                if (!string.IsNullOrWhiteSpace(condition.PolicyNo))
                {
                    filter += " and p.PolicyNo = @PolicyNo ";
                }
                if (!string.IsNullOrWhiteSpace(condition.Owner))
                {
                    filter += " and p.Owner = @Owner ";
                }
                if (!string.IsNullOrWhiteSpace(condition.Creator))
                {
                    string m = DbHelper.Query<string>(CUFRepository.ConnectionStringName, @"select imember from sc_member where nmember = @nmember", new { nmember = condition.Creator }).FirstOrDefault();
                    condition.Creator = m;
                    filter += " and c.DoUser = @Creator ";
                }
            }
            filter += @"group by c.[type], c.SourceID, c.[no], c.CreateTime, p.Owner, p.AgentName, p.SUAgentName, p.SUAgentCode
                        order by c.CreateTime asc ";
            sql += filter;

            var result = DbHelper.Query<MaintainInfo>(
                H2ORepository.ConnectionStringName,
                sql,
                param: condition
            ).ToList();

            for (int i = 0; i <= result.Count - 1; i++) {
                result[i].Serial = i + 1;
            }

            //2024.01.19 客訴系統維護優化第一階段，增加狀態欄位。 by peggy
            //2024.04.25 客訴系統維護優化第二階段，調整狀態欄位。 by peggy
            if (condition != null)
            {
                if (condition.IsClose)
                {
                    result = result.Where(x => x.Status == "結案").Select(x => x).ToList();
                }
                if (!condition.IsClose)
                {
                    result = result.Where(x => x.Status != "結案").Select(x => x).ToList();
                }
                switch (condition.Status)
                {
                    case StatusCategory.Audit:
                        //稽催
                        result = result.Where(x => x.Status == "稽催中").Select(x => x).ToList();
                        break;
                    case StatusCategory.Close:
                        //結案
                        result = result.Where(x => x.Status == "結案").Select(x => x).ToList();
                        break;
                    case StatusCategory.Dos:
                        //催辦
                        result = result.Where(x => x.Status == "催辦中").Select(x => x).ToList();
                        break;
                    case StatusCategory.Note:
                        //照會
                        result = result.Where(x => x.Status == "照會中").Select(x => x).ToList();
                        break;
                    case StatusCategory.Process:
                        //立案不通知
                        result = result.Where(x => x.Status == "立案不通知").Select(x => x).ToList();
                        break;
                    case StatusCategory.All:
                        //全部
                        break;
                }
            }

            return result;
        }

        /// <summary>
        /// 執行稽催
        /// </summary>
        /// <param name="cRMEAudit">稽催model</param>
        /// <param name="ip">user id</param>
        /// <param name="userMemberName">user name</param>
        /// <returns></returns>
        public bool DoAudit(CRMEAudit cRMEAudit, string ip, string userMemberName, string userMemberId)
        {
            var result = false;
            if (cRMEAudit != null) 
            {
                try
                {
                    // 1. 寫入稽催主檔
                    cRMEAudit.Insert(new[] { H2ORepository.ConnectionStringName });

                    CRMEInsurancePolicy policy = H2ORepository.Select<CRMEInsurancePolicy>().Where(x => x.No == cRMEAudit.No).FirstOrDefault();
                    string serviceName = String.IsNullOrWhiteSpace(policy.SUAgentName) ? policy.AgentName : policy.SUAgentName;
                    string serviceCode = String.IsNullOrWhiteSpace(policy.SUAgentName) ? policy.AgentCode : policy.SUAgentCode;
                    string maskOwner = GetPersonalMask(policy.Owner, "*");
                    string sendTitle = "稽催：保戶(" + maskOwner + "）服務照會單通知-" + policy.No + "(業務員：" + serviceName + ")";
                    string sendContent = "通知(服務人員：" + serviceName + ")保戶 " + maskOwner + "，照會單號：" + policy.No + "，尚未收到回覆，請於 EIP系統 業務發展 > 業務支援 > 客服業務 > 通知作業 下載稽催單，並於「二個工作日內」書面回覆總公司客服部，謝謝!";

                    //2. 取得要通知的人員名單  2024.01.19 客訴系統優化第一階段新增。 by peggy
                    var notifyToList = DbHelper.Query<CRMENotifyTo>(H2ORepository.ConnectionStringName, @"select no, NotifyType, MemberID +','+ m.nmember as MemberID from CRMENotifyTo o　inner join cuf.dbo.sc_member m on o.MemberID = m.imember where no = @no group by no, NotifyType, MemberID, m.nmember", new { no = cRMEAudit.No }).ToList();

                    Dictionary<string, string> keyValues = new Dictionary<string, string>();

                    foreach (CRMENotifyTo cRMENotifyTo in notifyToList)
                    {
                        var memberIdSplit = cRMENotifyTo.MemberID.Split(',');

                        //副本主管已終止: 移除留言對象(即不發留言推播)
                        if ((cRMENotifyTo.NotifyType == NotifyType.CC) && (CheckAgentIsExisted(memberIdSplit[0]) == false))
                        {
                            continue;
                        }

                        //行專已離職: 留言對象改為「實駐單位」的行專
                        if ((cRMENotifyTo.NotifyType == NotifyType.Employee) && (CheckAgentIsExisted(memberIdSplit[0]) == false))
                        {
                            List<string> list = DbHelper.Query<string>(H2ORepository.ConnectionStringName, @"select memExt.MemberID+','+peo.people_name from h2o.dbo.people peo 
                                                                                                                inner join cuf.dbo.v_MemberExtendID memExt on peo.people_orgid = memExt.ExtendID 
                                                                                                                where people_unit_code = @wcCode ", new { wcCode = policy.SUAgentCode == null ? policy.WCCode : policy.SUWCCode }).ToList();
                            foreach (string id in list)
                            {
                                var idSplit = id.Split(',');
                                if (keyValues.ContainsKey(idSplit[0]) == false)
                                {
                                    keyValues.Add(idSplit[0], idSplit[1]);
                                }
                            }
                        }
                        else
                        {
                            if (keyValues.ContainsKey(memberIdSplit[0]) == false)
                            {
                                keyValues.Add(memberIdSplit[0], memberIdSplit[1]);
                            }
                        }
                    }

                    foreach (KeyValuePair<string, string> kvp in keyValues)
                    {
                        MessageAndLineNotify(sendTitle, sendContent, kvp, ip, userMemberName, userMemberId, "CUSCRMTX003_Audit");
                    }

                    result = true;
                }
                catch (Exception ex)
                {
                    Throw.BusinessError(ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 執行催辦
        /// </summary>
        /// <param name="CRMEDoS">催辦model</param>
        /// <param name="ip">user id</param>
        /// <param name="userMemberName">user name</param>
        /// <returns></returns>
        public bool DoS(CRMEDoS cRMEDoS, string ip, string userMemberName, string userMemberId)
        {
            var result = false;
            if (cRMEDoS != null)
            {
                try
                {
                    // 1. 寫入催辦主檔
                    cRMEDoS.Insert(new[] { H2ORepository.ConnectionStringName });

                    CRMEInsurancePolicy policy = H2ORepository.Select<CRMEInsurancePolicy>().Where(x => x.No == cRMEDoS.No).FirstOrDefault();
                    string serviceName = String.IsNullOrWhiteSpace(policy.SUAgentName) ? policy.AgentName : policy.SUAgentName;
                    string serviceCode = String.IsNullOrWhiteSpace(policy.SUAgentName) ? policy.AgentCode : policy.SUAgentCode;
                    string maskOwner = GetPersonalMask(policy.Owner, "*");
                    string sendTitle = "催辦：保戶(" + maskOwner + "）服務照會單通知-" + policy.No + "(業務員：" + serviceName + ")";
                    string sendContent = "通知(服務人員：" + serviceName + ")保戶 " + maskOwner + " 有新增照會事項，請於 EIP系統 業務發展 > 業務支援 > 客服業務 > 通知作業 下載催辦單，照會單號 " + policy.No + "，並於「三個工作日內」書面回覆總公司客服部，謝謝!";

                    //2. 取得要通知的人員名單  2024.01.19 客訴系統優化第一階段新增。 by peggy
                    var notifyToList = DbHelper.Query<CRMENotifyTo>(H2ORepository.ConnectionStringName, @"select no, NotifyType, MemberID +','+ m.nmember as MemberID from CRMENotifyTo o　inner join cuf.dbo.sc_member m on o.MemberID = m.imember where no = @no group by no, NotifyType, MemberID, m.nmember", new { no = cRMEDoS.No }).ToList();
                    
                    Dictionary<string, string> keyValues = new Dictionary<string, string>();
                    
                    foreach (CRMENotifyTo cRMENotifyTo in notifyToList)
                    {
                        var memberIdSplit = cRMENotifyTo.MemberID.Split(',');

                        //副本主管已終止: 移除留言對象(即不發留言推播)
                        if ((cRMENotifyTo.NotifyType == NotifyType.CC) && (CheckAgentIsExisted(memberIdSplit[0]) == false))
                        {
                            continue;
                        }

                        //行專已離職: 留言對象改為「實駐單位」的行專
                        if ((cRMENotifyTo.NotifyType == NotifyType.Employee) && (CheckAgentIsExisted(memberIdSplit[0]) == false))
                        {
                            List<string> list = DbHelper.Query<string>(H2ORepository.ConnectionStringName, @"select memExt.MemberID+','+peo.people_name from h2o.dbo.people peo 
                                                                                                                inner join cuf.dbo.v_MemberExtendID memExt on peo.people_orgid = memExt.ExtendID 
                                                                                                                where people_unit_code = @wcCode ", new { wcCode = policy.SUAgentCode == null ? policy.WCCode : policy.SUWCCode }).ToList();
                            foreach (string id in list)
                            {
                                var idSplit = id.Split(',');
                                if (keyValues.ContainsKey(idSplit[0]) == false)
                                {
                                    keyValues.Add(idSplit[0], idSplit[1]);
                                }
                            }
                        }
                        else 
                        {
                            if (keyValues.ContainsKey(memberIdSplit[0]) == false)
                            {
                                keyValues.Add(memberIdSplit[0], memberIdSplit[1]);
                            }
                        }
                    }

                    foreach (KeyValuePair<string, string> kvp in keyValues) 
                    {
                        MessageAndLineNotify(sendTitle, sendContent, kvp, ip, userMemberName, userMemberId, "CUSCRMTX003_Dos");
                    }

                    result = true;
                }
                catch (Exception ex)
                {
                    Throw.BusinessError(ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 留言&ePower推播
        /// </summary>
        /// <param name="sendTitle">標題</param>
        /// <param name="sendContent">內文</param>
        /// <param name="serviceCode">留言對象員編</param>
        /// <param name="ip">user id</param>
        /// <param name="userMemberName">user name</param>
        public void MessageAndLineNotify(string sendTitle, string sendContent, KeyValuePair<string, string> valuePair, string ip, string userMemberName, string userMemberId, string msgNote)
        {
            Message message = new Message();
            using (TransactionScope ts = new TransactionScope())
            {
                //設定留言主檔
                message.MSGClass = 3;
                message.MSGPREVID = 0;
                message.MSGSubject = sendTitle;
                message.MSGDESC = sendContent;
                message.MSGCreater = userMemberId;
                message.MSGTime = DateTime.Now;
                message.MSGCreateIP = ip;
                message.MSGCreateName = userMemberName;
                message.MSGNote = msgNote;
                message.Insert(new[] { H2ORepository.ConnectionStringName });
                ts.Complete();
            }

            //推播給業務員
            MemberExtendService extendService = new MemberExtendService();

            //需求單20240711003，改只推播給業務員，不給行專行政人員
            if (extendService.GetMemberExtendDetail(valuePair.Key).MemberExtendInfo.CreateBy.ToUpper() == "VLIFE") 
            {
                var account = extendService.GetMemberAccount(valuePair.Key).FirstOrDefault();
                LineNotifyInfo lineNotifyInfo = new LineNotifyInfo();
                lineNotifyInfo.Subject = sendTitle;
                lineNotifyInfo.Content = sendContent;
                lineNotifyInfo.Sender = "cufadmin";
                string ePowerSql = @"exec dbo.usp_LINENotify_Send @ProgramType , @UserID, @Category , @Title , @Content";
                var ePowercondition = new
                {
                    ProgramType = "客訴系統維護推播",
                    UserID = account.iaccount,
                    Category = "CUSCRMTX003",
                    Title = sendTitle,
                    Content = sendContent
                };
                DbHelper.Execute(MISRepository.ConnectionStringName, ePowerSql, ePowercondition);
            }
  

            //設定留言對象
            string msgToSql = @"insert into MessageTo(MSGOBJMSGID, MSGOBJID, MSGOBJSendIP, MSGOBJCreateTime, MSGOBJSendID, MSGOBJName, MSGOBJNote)
                            values (@MSGOBJMSGID, @MSGOBJID, @MSGOBJSendIP, @MSGOBJCreateTime, @MSGOBJSendID, @MSGOBJName, @MSGOBJNote)";
            
            var msgTocondition = new
            {
                MSGOBJMSGID = message.MSGID,
                MSGOBJID = valuePair.Key,
                MSGOBJSendIP = ip,
                MSGOBJCreateTime = DateTime.Now,
                MSGOBJSendID = userMemberId,
                MSGOBJName = valuePair.Value,
                MSGOBJNote = msgNote
            };
            DbHelper.Execute(H2ORepository.ConnectionStringName, msgToSql, msgTocondition);
        }

        /// <summary>
        /// 產生處理單資料
        /// </summary>
        /// <param name="no">受理編號</param>
        /// <returns></returns>
        public ProcessFormData ProduceProcessFormData(string no)
        {
            //1. 取立案資料
            CRMECaseContent cRMECaseContent =  DbHelper.Query<CRMECaseContent>(H2ORepository.ConnectionStringName,
                                                @"Select * from CRMECaseContent where No=@no", new { No = no }).FirstOrDefault();

            //2. 取立案保單資料
            List<CRMEInsurancePolicy> cRMEInsurancePolicies = DbHelper.Query<CRMEInsurancePolicy>(H2ORepository.ConnectionStringName,
                                                @"Select * from CRMEInsurancePolicy where No=@no", new { No = no }).ToList();
            //3. 取處理過程
            List<CRMEDoInfo> cRMEDos = DbHelper.Query<CRMEDoInfo>(H2ORepository.ConnectionStringName,
                                                @"Select No, Content,
                                                        isnull(STUFF((
                                                            SELECT ',' + convert(varchar, d1.BUSContactCompanyDate, 111) FROM CRMEDo d1  WHERE  d1.No = do.No 
                                                            FOR XML PATH('')
                                                         ), 1, 1, ''),'') AS HistoryBUSContactCompanyDate
                                                         ,
                                                         isnull(STUFF((
                                                            SELECT ',' + convert(varchar, d2.ReplyCompanyDate, 111) FROM CRMEDo d2  WHERE  d2.No = do.No 
                                                            FOR XML PATH('')
                                                         ), 1, 1, ''),'') AS HistoryReplyCompanyDate
                                                         ,
                                                         isnull(STUFF((
                                                            SELECT ',' + convert(varchar, d3.ReplyYallowBillDate, 111) FROM CRMEDo d3  WHERE  d3.No = do.No 
                                                            FOR XML PATH('')
                                                         ), 1, 1, ''),'') AS HistoryReplyYallowBillDate
                                                from CRMEDo do where No=@no", new { No = no }).ToList();

            ProcessFormData formData = new ProcessFormData();
            //受理編號
            formData.No = cRMECaseContent.No;

            //承辦人
            formData.CaseHandler = cRMECaseContent.DoUser;

            //受理日期
            formData.CreatTime = cRMECaseContent.CreateTime.ToString("yyyy年MM月dd日");

            //來源
            formData.Source = cRMECaseContent.Type;

            //要保人
            formData.Owner = cRMEInsurancePolicies.Select(x=>x.Owner).FirstOrDefault();

            //組保單基本資料清單表格
            int count = 0;
            formData.PolicyHistory += @"<table style= 'width: 100 %;' border='1'>";
            foreach (CRMEInsurancePolicy policy in cRMEInsurancePolicies) 
            {
                #region 設定保單基本資料表頭
                if (count == 0) 
                {
                    formData.PolicyHistory += @"<tr>";
                    formData.PolicyHistory += @"<td>保單號碼</td>";
                    formData.PolicyHistory += @"<td>保險公司</td>";
                    formData.PolicyHistory += @"<td>要保人</td>";
                    formData.PolicyHistory += @"<td>被保人</td>";
                    formData.PolicyHistory += @"<td>生效日</td>";
                    formData.PolicyHistory += @"<td>保單狀態</td>";
                    formData.PolicyHistory += @"<td>原幣別保費</td>";
                    formData.PolicyHistory += @"<td>繳別</td>";
                    formData.PolicyHistory += @"<td>實駐</td>";
                    formData.PolicyHistory += @"<td>原招攬人</td>";
                    formData.PolicyHistory += @"<td>接續服務人員</td>";
                    formData.PolicyHistory += @"</tr>";
                }
                #endregion

                #region 設定保單資料
                formData.PolicyHistory += @"<tr>";
                formData.PolicyHistory += @"<td>" + policy.PolicyNo + @"</td>";
                formData.PolicyHistory += @"<td>" + policy.CompanyName + @"</td>";
                formData.PolicyHistory += @"<td>" + policy.Owner + @"</td>";
                formData.PolicyHistory += @"<td>" + policy.Insured + @"</td>";
                formData.PolicyHistory += @"<td>" + policy.IssDate + @"</td>";
                formData.PolicyHistory += @"<td>" + policy.StatusName + @"</td>";
                formData.PolicyHistory += @"<td>" + policy.ModPrem + @"</td>";
                formData.PolicyHistory += @"<td>" + policy.ModxName + @"</td>";
                formData.PolicyHistory += @"<td>" + (policy.SUCenterName ?? policy.CenterName) + @"</td>";
                formData.PolicyHistory += @"<td>" + policy.AgentName + @"</td>";
                formData.PolicyHistory += @"<td>" + policy.SUAgentName + @"</td>";
                formData.PolicyHistory += @"</tr>";
                #endregion
                count++;
            }
            formData.PolicyHistory += @"</table>";

            //組申訴及服務內容
            formData.ContentTXT = cRMECaseContent.Content;

            //處理過程
            int frequency = 1;
            StringBuilder processRecord = new StringBuilder();
            foreach (CRMEDoInfo dos in cRMEDos) 
            {
                if (frequency == 1) 
                {
                    formData.ProcessRecord += string.Format("{0}: {1}{2}", "業連保險公司日期", dos.HistoryBUSContactCompanyDate, "<br />");
                    formData.ProcessRecord += string.Format("{0}: {1}{2}", "保險公司回覆日", dos.HistoryReplyCompanyDate, "<br />");
                    formData.ProcessRecord += string.Format("{0}: {1}{2}", "回覆單位黃聯日", dos.HistoryReplyYallowBillDate, "<br />");
                }
                formData.ProcessRecord += string.Format("{1}{2}", frequency.ToString(), dos.Content, "<br />");
                frequency++;
            }

            //處理結果
            formData.CloseRecord = DbHelper.Query<string>(H2ORepository.ConnectionStringName,
                                                @"SELECT dis.Name FROM CRMEcloselog l left join CRMEdisciptype dis on l.ResultCode = dis.id  WHERE  l.No = @No order by l.id desc ", new { No = no }).FirstOrDefault();

            return formData;
        }

        /// <summary>
        /// 根據類型取得歷史維護清單
        /// </summary>
        /// <param name="no">受理編號</param>
        /// <returns></returns>
        public MtnHistoryInfo GetDoHistory(string no)
        {
            string code = "";
            if (no.Contains("CS")) {
                code = "CS";
            }else if (no.Contains("SS")) {
                code = "SS";
            }else if (no.Contains("CC")) {
                code = "CC";
            }else if (no.Contains("SC")) {
                code = "SC";
            }
            //整理歷次聯絡紀錄日期
            string hisSql = @"SELECT  convert(varchar, CreateTime, 111) AS CreateDate,
                                 isnull(STUFF((
                                    SELECT ',' + convert(varchar, a.CreateTime, 111) FROM CRMEAudit a  WHERE  a.No = c.No 
                                    FOR XML PATH('')
                                 ), 1, 1, ''),'') AS HistoryAuditDate
                                 ,
                                 isnull(STUFF((
                                    SELECT ',' + convert(varchar, s.CreateTime, 111) FROM CRMEDoS s  WHERE  s.No = c.No 
                                    FOR XML PATH('')
                                 ), 1, 1, ''),'') AS HistoryDoSDate
                                 ,
                                 isnull(STUFF((
                                    SELECT ',' + d.Content FROM CRMEDo d  WHERE  d.No = c.No 
                                    FOR XML PATH('')
                                 ), 1, 1, ''),'') AS HistoryContent
                                 ,
                                 isnull(STUFF((
                                    SELECT ',' + convert(varchar, d1.BUSContactCompanyDate, 111) FROM CRMEDo d1  WHERE  d1.No = c.No 
                                    FOR XML PATH('')
                                 ), 1, 1, ''),'') AS HistoryBUSContactCompanyDate
                                 ,
                                 isnull(STUFF((
                                    SELECT ',' + convert(varchar, d2.ReplyCompanyDate, 111) FROM CRMEDo d2  WHERE  d2.No = c.No 
                                    FOR XML PATH('')
                                 ), 1, 1, ''),'') AS HistoryReplyCompanyDate
                                 ,
                                 isnull(STUFF((
                                    SELECT ',' + convert(varchar, d3.ReplyYallowBillDate, 111) FROM CRMEDo d3  WHERE  d3.No = c.No 
                                    FOR XML PATH('')
                                 ), 1, 1, ''),'') AS HistoryReplyYallowBillDate
                                 ,
                                 isnull(STUFF((
                                    SELECT ',' + convert(varchar, d4.CCReplyDate, 111) FROM CRMEDo d4  WHERE  d4.No = c.No 
                                    FOR XML PATH('')
                                 ), 1, 1, ''),'') AS HistoryCCReplyDate
                                ,
                                 isnull(STUFF((
                                    SELECT ',' + dis.Name FROM CRMEcloselog l left join CRMEdisciptype dis on l.ResultCode = dis.id  WHERE  l.No = c.No 
                                    FOR XML PATH('')
                                 ), 1, 1, ''),'') AS HistoryCloseLog
                                ,
								 (select top 1 [Name] from CRMEDo left join CRMEDiscipType on CRMEDo.CompanyResult = CRMEDiscipType.ID where no = @no order by CRMEDo.CreateTime desc ) as HistoryCompanyResult
                                FROM CRMECaseContent c
                                where c.No = @No ";
            MtnHistoryInfo mtnInfo = DbHelper.Query<MtnHistoryInfo>(H2ORepository.ConnectionStringName, hisSql, new { No = no }).FirstOrDefault();

            //取得處理結果選單
            string disSql = @"select * from CRMEDiscipType where code = @code and (kind = 7 or kind = 8) and [Status] = 1  order by kind, sort";
            var disResult = DbHelper.Query<CRMEDiscipType>(H2ORepository.ConnectionStringName, disSql, new { code = code }).ToList();
            mtnInfo.ResultDiscipType = disResult;

            //取得保公決議選單
            string cmpSql = @"select * from CRMEDiscipType where code = @code and (kind = 9) and [Status] = 1 order by kind, sort";
            var cmpResult = DbHelper.Query<CRMEDiscipType>(H2ORepository.ConnectionStringName, cmpSql, new { code = code }).ToList();
            mtnInfo.CompanyDiscipType = cmpResult;

            //取得檔案附件
            List<CRMEFile> files = DbHelper.Query<CRMEFile>(H2ORepository.ConnectionStringName,
                                                @"Select * from CRMEFile where SourceType = 1 and FolderNo=@No", new { No = no }).ToList(); ;
            mtnInfo.Files = files;

            return mtnInfo;
        }

        /// <summary>
        /// 將上傳的檔案存入DB
        /// </summary>
        /// <param name="cRMEFile">檔案model</param>
        public bool AddUploadFile(CRMEFile cRMEFile)
        {
            return cRMEFile.Insert(new[] { H2ORepository.ConnectionStringName });
        }

        /// <summary>
        /// 存入一筆維護紀錄
        /// </summary>
        /// <param name="cRMEDo">維護紀錄model</param>
        public bool InserCRMEDo(CRMEDo cRMEDo)
        {
            return cRMEDo.Insert(new[] { H2ORepository.ConnectionStringName });
        }

        /// <summary>
        /// 依檔案ID取檔案資訊
        /// </summary>
        /// <param name="id">檔案ID</param>
        /// <returns></returns>
        public CRMEFile GetCRMEFileById(int id) 
        {
            return DbHelper.Query<CRMEFile>(H2ORepository.ConnectionStringName, @"select * from CRMEFile where ID = @ID", new { ID = id }).FirstOrDefault();
        }

        /// <summary>
        /// 依照檔案ID刪除實體檔案
        /// </summary>
        /// <param name="id">檔案ID</param>
        /// <returns></returns>
        public bool DeleteFileByCRMEFIleId(CRMEFile cRMEFile)
        {
            IDataSettingService settingService = new DataSettingService();
            var tempDirRoot = settingService.GetConfigValueByName("CUSCRMDir");
            string fullName = Path.Combine(tempDirRoot, cRMEFile.FolderNo, cRMEFile.FileMD5Name);
            if (File.Exists(fullName)) 
            {
                File.Delete(fullName);
            }
            return true;
        }

        /// <summary>
        /// 依檔案ID刪除實體檔案與檔案資訊
        /// </summary>
        /// <param name="list">檔案ID</param>
        /// <returns></returns>
        public bool DeleteCRMEFileDbAndFileById(List<int> list)
        {
            foreach (int i in list) 
            {
                CRMEFile cRMEFile = DbHelper.Query<CRMEFile>(H2ORepository.ConnectionStringName, @"select * from CRMEFile where ID = @ID", new { ID = i }).FirstOrDefault();
                if (cRMEFile != null) 
                {
                    DeleteFileByCRMEFIleId(cRMEFile);
                    H2ORepository.Delete<CRMEFile>(new { ID = cRMEFile.ID });
                }
            }
            return true;
        }

        /// <summary>
        /// 結案
        /// </summary>
        /// <param name="cRMECloseLog">結案model</param>
        /// <returns></returns>
        public bool InsertCloseLog(CRMECloseLog cRMECloseLog)
        {
            return cRMECloseLog.Insert(new[] { H2ORepository.ConnectionStringName });
        }

        #region 姓名隱碼
        private static string GetPersonalMask(string OriStr, string oPrefix, bool IsDis)
        {
            string oStr = "";
            OriStr = OriStr.Trim();
            char[] oStrArry = OriStr.ToCharArray();
            int[] oArray = new int[] { 0, OriStr.Trim().Length - 1 };

            if (Regex.IsMatch(OriStr.Trim(), @"^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$"))
            {
                oStr += GetPersonalMask(OriStr.Split('@')[0], oPrefix) + "@";

                for (int i = 0; i < OriStr.Split('@')[1].Split('.').Length; i++)
                {
                    string oStrL = OriStr.Split('@')[1].Split('.')[i].ToString();
                    if (i == 0)
                        oStr += GetPersonalMask(oStrL, oPrefix, false);
                    else
                        oStr += "." + GetPersonalMask(oStrL, oPrefix, false);
                }
                return oStr;
            }
            else if (Regex.IsMatch(OriStr.Trim(), "^(09([0-9]){8})$"))
            {
                oArray = new int[] { 0, 1, 2, 7, 8, 9 };
            }
            else if (Regex.IsMatch(OriStr.Trim(), "^[a-zA-Z][0-9]{9}$"))
            {
                oArray = new int[] { 0, 1, 2, 3, 9 };
            }

            for (int i = 0; i < oStrArry.Length; i++)
            {
                if (IsDis)
                    oStr += oArray.Contains(i) ? oStrArry[i].ToString() : oPrefix;
                else
                    oStr += oArray.Contains(i) ? oPrefix : oStrArry[i].ToString();
            }
            return oStr;
        }

        private static string GetPersonalMask(string OriStr, string oPrefix)
        {
            return GetPersonalMask(OriStr, oPrefix, true);
        }
        #endregion

        #region 稽催、催辦判斷點
        /// <summary>
        /// 判斷該案件當初實駐是否仍存在
        /// </summary>
        /// <param name="no"></param>
        /// <param name="imemberId"></param>
        /// <param name="wcCode"></param>
        /// <returns></returns>
        public bool CheckPeoleWcCode(string no)
        {
            //1. 取案件原始檔
            CRMEInsurancePolicy policy = H2ORepository.Select<CRMEInsurancePolicy>().Where(x => x.No == no).FirstOrDefault();

            //2. 判斷實駐是否還存在
            int count = DbHelper.Query<string>(H2ORepository.ConnectionStringName, @"select people_unit_code from people where people_unit_code = @wcCode", new { wcCode = policy.WCCode }).Count();
            if (count <= 0)
            {
                //實駐不存在
                return false;
            }
            else 
            {
                return true;
            }
        }

        /// <summary>
        /// 判斷人員是否在職
        /// </summary>
        /// <param name="imember">人員編號</param>
        /// <returns>true: 仍在職; false: 已離職</returns>
        public bool CheckAgentIsExisted(string imember)
        {
            int count = DbHelper.Query<string>(CUFRepository.ConnectionStringName, @"select imember from sc_account where imember = @imember and dexpire <= getdate() ", new { imember = imember }).Count();
            if (count > 0)
            {
                //已離職or報聘中
                return false;
            }
            else 
            {
                return true;
            }
        }

        #endregion
    }
}
