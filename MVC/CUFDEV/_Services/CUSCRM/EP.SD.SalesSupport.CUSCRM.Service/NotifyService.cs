using EP.CUFModels;
using EP.H2OModels;
using EP.Platform.Service;
using EP.SD.SalesSupport.CUSCRM.Service.Contracts;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
	public class NotifyService : INotifyService
	{
		#region 立案通知
		/// <summary>
		/// 取得立案通知資料的受文者、副本受文者、受理保單資料
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public List<NotifyInsuranceViewModel> GetCRMEInsurancePolicy(string No)
		{
			List<NotifyInsuranceViewModel> result = new List<NotifyInsuranceViewModel>();
			string sql = @"select * from CRMEInsurancePolicy
            where no = @no
            ";

			result = DbHelper.Query<NotifyInsuranceViewModel>(H2ORepository.ConnectionStringName, sql, new { No = No }).ToList();
			return result;
		}

		/// <summary>
		/// 處代理人
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public List<NotifyInsuranceViewModel> GetCRMEChinaOm(string No)
		{
			List<NotifyInsuranceViewModel> result = new List<NotifyInsuranceViewModel>();
			string sql = string.Empty;
			sql = @"select  distinct b.ag_proxy as CHAgentCode,b.ag_proxy_name as CHAgentName 
				from CRMEInsurancePolicy a join SL_EPDB_VLIFE.Vlife.dbo.v_inChina_Proxy b on (case when a.SUCenterLeaderID is not null then a.SUCenterLeaderID  else a.CenterLeaderID END) = b.agent_code  
				where Proxy_Type = 'om' and No = @No";

			result = DbHelper.Query<NotifyInsuranceViewModel>(H2ORepository.ConnectionStringName, sql, new { No = No }).ToList();
			return result;
		}

		/// <summary>
		/// 取得客服部人員
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public List<sc_member> GetCustomerUnit()
		{
			string sql = @"select * from sc_member m join sc_unit u on m.uunit = u.uunit where iunit in ('TEP019','TEP034','TEP033') and imember != 'CUSAC00001' and dleave_office is null";

			var result = new List<sc_member>();

			result = DbHelper.Query<sc_member>(CUFRepository.ConnectionStringName, sql).ToList();
			return result;
		}

		/// <summary>
		/// 取得實駐
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public Dictionary<string, string> GetWcCenterByAgentCode()
		{
			string sql = @" select agent_code as 'Key',wc_center as 'Value' from orgin
            ";

			Dictionary<string, string> result = new Dictionary<string, string>();

			DbHelper.Query<KeyValuePair<string, string>>(VLifeRepository.ConnectionStringName, sql).ForEach(x =>
			{
				result.Add(x.Key, x.Value);
			});
			return result;
		}

		/// <summary>
		/// 取得實駐名稱
		/// </summary>
		/// <returns></returns>
		public Dictionary<string, string> GetWcCenterNameByAgentCode()
		{
			string sql = @" select agent_code as 'Key',
			case when SUBSTRING(wc_center_name,0,CHARINDEX('(',wc_center_name)) = '' then wc_center_name 
			else SUBSTRING(wc_center_name, CHARINDEX('(', wc_center_name) + 1, CHARINDEX(')', wc_center_name) - CHARINDEX('(', wc_center_name) - 1) END as 'Value'  
			from orgin
            ";

			Dictionary<string, string> result = new Dictionary<string, string>();

			DbHelper.Query<KeyValuePair<string, string>>(VLifeRepository.ConnectionStringName, sql).ForEach(x =>
			{
				result.Add(x.Key, x.Value);
			});
			return result;
		}

		/// <summary>
		/// 取得職稱
		/// </summary>
		/// <returns></returns>
		public Dictionary<string, string> GetLevelNameChsByAgentCode()
		{
			string sql = @"select agent_code as 'Key',level_name_chs as 'Value' from  orgin c 
            join v_aglevel_occpind d on c.ag_level = d.AG_LEVEL and c.ag_occp_ind = d.ag_occp_ind
            ";

			Dictionary<string, string> result = new Dictionary<string, string>();

			DbHelper.Query<KeyValuePair<string, string>>(VLifeRepository.ConnectionStringName, sql).ForEach(x =>
			{
				result.Add(x.Key, x.Value);
			});
			return result;

		}

		/// <summary>
		/// 取得VLife人名
		/// </summary>
		/// <returns></returns>
		public string GetAgentName(string agentcode)
		{
			string sql = @"select name from orgin where agent_code = @agentcode
            ";
			string result = string.Empty;

			result = DbHelper.Query<string>(VLifeRepository.ConnectionStringName, sql, new { agentcode = agentcode }).FirstOrDefault();
			return result ?? string.Empty;

		}

		/// <summary>
		/// 取得人員ID
		/// </summary>
		/// <returns></returns>
		public string GetAccountId(string MemberId)
		{
			string sql = @"select b.iaccount from sc_member a join sc_account b on a.imember = b.imember where b.imember = @MemberId";
			string result = string.Empty;

			result = DbHelper.Query<string>(CUFRepository.ConnectionStringName, sql, new { MemberId = MemberId }).FirstOrDefault();
			return result ?? string.Empty;

		}

		/// <summary>
		/// 取得行專資料
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public List<NotifyCaseViewModel> GetCRMENotifyEmployee(string code)
		{
			List<NotifyCaseViewModel> result = new List<NotifyCaseViewModel>();
			string sql = @"select people_orgid,people_name from people
            where people_unit_code = @code
            ";

			result = DbHelper.Query<NotifyCaseViewModel>(H2ORepository.ConnectionStringName, sql,
			new
			{
				code = code
			}).ToList();
			return result;
		}

		/// <summary>
		/// 取得主表資料
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public CRMECaseContent GetCRMECaseContent(string No)
		{
			CRMECaseContent result = new CRMECaseContent();
			string sql = @"select * from CRMECaseContent
            where no = @no
            ";

			result = DbHelper.Query<CRMECaseContent>(H2ORepository.ConnectionStringName, sql, new { No = No }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 新增立案通知對象
		/// </summary>
		/// <param name="modeltoList"></param>
		public void InsertCRMENotifyTo(List<CRMENotifyTo> modeltoList)
		{
			for (int i = 0; i < modeltoList.Count; i++)
			{
				modeltoList[i].Insert(new[] { H2ORepository.ConnectionStringName });
			}
		}

		/// <summary>
		/// 更新主表
		/// </summary>
		/// <param name="model"></param>
		/// <returns></returns>
		public bool UpdateCaseContentType(CRMECaseContent model)
		{
			var result = false;
			if (model != null)
			{
				using (TransactionScope ts = new TransactionScope())
				{
					string Sql = @"update CRMECaseContent set 
                        Status = @Status,
                        DoUser = @DoUser,
                        DoUserTelExt = @DoUserTelExt
                        where No = @No
                    ";
					DbHelper.Execute(H2ORepository.ConnectionStringName, Sql, new
					{
						Status = model.Status,
						No = model.No,
						DoUser = model.DoUser,
						DoUserTelExt = model.DoUserTelExt
					});

					result = true;
					ts.Complete();
				}
			}
			return result;
		}

		/// <summary>
		/// 新增照會附件
		/// </summary>
		/// <param name="model">主檔</param>
		/// <param name="modelfileList">檔案附件物件</param>
		/// <param name="tabUniqueId">前端傳來唯一碼</param>
		/// <param name="RetMsg">回傳訊息</param>
		public void CreateCRMNotifyFile(CRMECaseContent model, List<CRMEFile> modelfileList, string tabUniqueId, out string RetMsg)
		{
			RetMsg = "";

			try
			{
				IDataSettingService dsService = new DataSettingService();
				var tempDirRoot = dsService.GetConfigValueByName("TempDir");
				string currRootDir = string.Empty;
				if (modelfileList.Count > 0)
				{
					currRootDir = Path.Combine(tempDirRoot, modelfileList[0].Creator + @"\" + tabUniqueId);
				}				 
				string newRootDir = dsService.GetConfigValueByName("CUSCRMDir");

				#region 執行資料
				string currefile = string.Empty;
				string extname = string.Empty;
				string newfilename = string.Empty;

				using (TransactionScope ts = new TransactionScope())
				{
					for (int i = 0; i < modelfileList.Count; i++)
					{
						currefile = Path.Combine(currRootDir, modelfileList[i].FileMD5Name);
						modelfileList[i].FileMD5Name = DateTime.Now.ToString("yyyyMM") + @"\" + modelfileList[i].FileMD5Name;
						System.IO.FileInfo fi = new System.IO.FileInfo(currefile);
						newfilename = Path.Combine(newRootDir, modelfileList[i].FileMD5Name);
						string uploaddir = Path.GetDirectoryName(newfilename);
						if (!Directory.Exists(uploaddir))
							Directory.CreateDirectory(uploaddir);

						modelfileList[i].Insert(new[] { H2ORepository.ConnectionStringName });
						fi.CopyTo(newfilename);
					}
					ts.Complete();
				}
				#endregion
			}
			catch (IOException ioe)
			{
				Throw.SQL("檔案儲存錯誤:" + ioe.Message);
				RetMsg = ioe.Message;
			}
			catch (Exception e)
			{
				Throw.SQL(e.Message);
				RetMsg = e.Message;
			}

		}

		/// <summary>
		/// 發送推播
		/// </summary>
		/// <param name="model"></param>
		/// <param name="Account"></param>
		/// <returns></returns>
		public void SendLineNotify(Message model, List<UserSimpleInfo> Account)
		{
			string sql = @"exec dbo.usp_LINENotify_Send @ProgramType , @UserID, @Category , @Title , @Content";
			using (IDbConnection Connectiton = MISRepository.CreateConnection())
			{
				for (int i = 0; i < Account.Count; i++)
				{
					if (!String.IsNullOrEmpty(Account[i].AccountId))
					{
						var condition = new
						{
							ProgramType = "客服業務系統",
							UserID = Account[i].AccountId,
							Category = "EIP",
							Title = model.MSGSubject,
							Content = model.MSGDESC
						};
						DbHelper.Query<decimal>(Connectiton, sql, condition).FirstOrDefault();
					}
				}
			}
		}
		#endregion

		#region 查詢通知作業
		/// <summary>
		/// 查詢通知作業-客服
		/// </summary>
		/// <param name="condition"></param>
		/// <returns></returns>
		public List<NotifyViewModel> QueryNotifyDataByCRM(QueryNotifyCondition condition)
		{
			string sql = @"SELECT
                        U.Type,
						U.CaseGuid,
                        U.No,
	                    U.CreateTime,
                        (Select Top 1 Owner from CRMEInsurancePolicy c where c.no = U.No order by ID) Owner,
                        ISNULL(STUFF(ToList.ToValue, 1, 1, ''), '') AS ToMemberID,
	                    (Select Top 1 d.CreateTime from CRMEDoS d where d.no = U.No order by ID desc) DoSCreateTime,
	                    (Select Top 1 e.CreateTime from CRMEAudit e where e.no = U.No order by ID desc) AuditCreateTime,
	                    ISNULL(STUFF(AllToList.AllToValue, 1, 1, ''), '')  AS AllToMemberID,
						ISNULL(STUFF(AllCCList.AllCCValue, 1, 1, ''), '')  AS AllCCMemberID,
						ISNULL(STUFF(FileList.FileValue, 1, 1, ''), '') AS CFID
                    FROM CRMECaseContent U 
                    OUTER APPLY (
                        SELECT
                            ',' + s.name
                        FROM CRMENotifyTo UR  join SL_EPDB_VLIFE.Vlife.dbo.orgin s on UR.MemberID = s.agent_code
                        WHERE U.No = UR.No and NotifyType = 0
                        FOR XML PATH ('')
                    ) ToList (ToValue)
                    OUTER APPLY (
                        SELECT
                            ',' + s.nmember					
                        FROM CRMENotifyTo UR  join cuf.dbo.sc_member s on UR.MemberID = s.imember and UR.NotifyType in (2,3,4)
                        WHERE U.No = UR.No
                        FOR XML PATH ('')
                    ) AllCCList (AllCCValue)
					OUTER APPLY (
                        SELECT
							',' + s.name						
                        FROM CRMENotifyTo UR  join SL_EPDB_VLIFE.Vlife.dbo.orgin s on UR.MemberID = s.agent_code and UR.NotifyType in (0,1)	
                        WHERE U.No = UR.No
                        FOR XML PATH ('')
                    ) AllToList (AllToValue)
                    OUTER APPLY (
                        SELECT
                            ',' + CONVERT(varchar, UR.ID) 
                        FROM CRMEFile UR  
                        WHERE U.No = UR.FolderNo and SourceType = 0
                        FOR XML PATH ('')
                    ) FileList (FileValue)
                    where 1=1";
			if (condition.Category == "0" && String.IsNullOrEmpty(condition.Type))
			{
				sql += " and (U.Type = 'CS' or U.Type = 'SS')";
			}
			else if (condition.Category == "1" && String.IsNullOrEmpty(condition.Type))
			{
				sql += " and (U.Type = 'CC' or U.Type = 'SC')";
			}
			if (condition.Status == 0)
			{				
				sql += " and U.no not in (select l.no from CRMECloseLog l group by l.no) ";
			}
			else
			{
				sql += " and U.no in (select l.no from CRMECloseLog l group by l.no) ";
			}

			if (!String.IsNullOrEmpty(condition.Type))
			{
				sql += " and U.Type = @Type";
			}
			sql += " order by CreateTime desc";
			var result = DbHelper.Query<NotifyViewModel>(H2ORepository.ConnectionStringName, sql, new { Status = condition.Status, Type = condition.Type }).ToList();
			return result;
		}

		/// <summary>
		/// 查詢通知作業-行專
		/// </summary>
		/// <param name="condition"></param>
		/// <returns></returns>
		public List<NotifyViewModel> QueryNotifyEPDataByCRM(QueryNotifyCondition condition)
		{
			string sql = @"SELECT
                        U.Type,
						U.CaseGuid,
                        U.No,
	                    U.CreateTime,
                        (Select Top 1 Owner from CRMEInsurancePolicy c where c.no = U.No order by ID) Owner,
                        ISNULL(STUFF(ToList.ToValue, 1, 1, ''), '') AS ToMemberID,
	                    (Select Top 1 d.CreateTime from CRMEDoS d where d.no = U.No order by ID desc) DoSCreateTime,
	                    (Select Top 1 e.CreateTime from CRMEAudit e where e.no = U.No order by ID desc) AuditCreateTime,
	                    ISNULL(STUFF(AllToList.AllToValue, 1, 1, ''), '')  AS AllToMemberID,
						ISNULL(STUFF(AllCCList.AllCCValue, 1, 1, ''), '')  AS AllCCMemberID,
						ISNULL(STUFF(FileList.FileValue, 1, 1, ''), '') AS CFID,
						T.MemberID
                    FROM CRMECaseContent U join CRMENotifyTo T on U.no = T.No and T.NotifyType in (2,3)
                    OUTER APPLY (
                        SELECT
                            ',' + s.name
                        FROM CRMENotifyTo UR  join SL_EPDB_VLIFE.Vlife.dbo.orgin s on UR.MemberID = s.agent_code
                        WHERE U.No = UR.No and NotifyType = 0
                        FOR XML PATH ('')
                    ) ToList (ToValue)
                    OUTER APPLY (
                        SELECT
                            ',' + s.nmember					
                        FROM CRMENotifyTo UR  join cuf.dbo.sc_member s on UR.MemberID = s.imember and UR.NotifyType in (2,3)
                        WHERE U.No = UR.No
                        FOR XML PATH ('')
                    ) AllCCList (AllCCValue)
					OUTER APPLY (
                        SELECT
							',' + s.name						
                        FROM CRMENotifyTo UR  join SL_EPDB_VLIFE.Vlife.dbo.orgin s on UR.MemberID = s.agent_code and UR.NotifyType in (0,1)	
                        WHERE U.No = UR.No
                        FOR XML PATH ('')
                    ) AllToList (AllToValue)
                    OUTER APPLY (
                        SELECT
                            ',' + CONVERT(varchar, UR.ID) 
                        FROM CRMEFile UR  
                        WHERE U.No = UR.FolderNo and SourceType = 0
                        FOR XML PATH ('')
                    ) FileList (FileValue)
                    where 1=1 and (U.Status = 1 or U.Status = 2) and U.no not in (select l.no from CRMECloseLog l group by l.no) and T.MemberID = @EPValue";
			if (condition.Category == "0" && String.IsNullOrEmpty(condition.Type))
			{
				sql += " and (U.Type = 'CS' or U.Type = 'SS')";
			}
			else if (condition.Category == "1" && String.IsNullOrEmpty(condition.Type))
			{
				sql += " and (U.Type = 'CC' or U.Type = 'SC')";
			}
			//if (condition.Status == 0)
			//{
			//	sql += " and (U.Status = 1 or U.Status = 2)";
			//}
			//else
			//{
			//	sql += " and U.Status = 4";
			//}
			if (!String.IsNullOrEmpty(condition.Type))
			{
				sql += " and U.Type = @Type";
			}
			sql += " order by CreateTime desc";
			var result = DbHelper.Query<NotifyViewModel>(H2ORepository.ConnectionStringName, sql, new { Status = condition.Status, Type = condition.Type, EPValue = condition.MemberID }).ToList();
			return result;
		}

		/// <summary>
		/// 查詢通知作業-業務員
		/// </summary>
		/// <param name="condition"></param>
		/// <returns></returns>
		public List<NotifyViewModel> QueryNotifyDataByAgent(QueryNotifyCondition condition)
		{
			string sql = @"SELECT
                        U.Type,
						U.CaseGuid,
                        U.No,
	                    U.CreateTime,
						 ISNULL(STUFF(ToList.ToValue, 1, 1, ''), '') AS ToMemberID,
                        (Select Top 1 Owner from CRMEInsurancePolicy c where c.no = U.No order by c.ID) Owner,
						ISNULL(STUFF(FileList.FileValue, 1, 1, ''), '') AS CFID,
						MemberID
                    FROM CRMECaseContent U　join  CRMENotifyTo T on U.No = T.No and MemberID = @MemberID					
					   OUTER APPLY (
                        SELECT
                            ',' + s.name
                        FROM CRMENotifyTo UR  join SL_EPDB_VLIFE.Vlife.dbo.orgin s on UR.MemberID = s.agent_code
                        WHERE U.No = UR.No and NotifyType = 0
                        FOR XML PATH ('')
                    ) ToList (ToValue)
					 OUTER APPLY (
                        SELECT
                            ',' + CONVERT(varchar, UR.ID) 
                        FROM CRMEFile UR  
                        WHERE U.No = UR.FolderNo and SourceType = 0
                        FOR XML PATH ('')
                    ) FileList (FileValue) where (U.Status = 1 or U.Status = 2) and U.no not in (select l.no from CRMECloseLog l group by l.no) order by CreateTime desc";

			var result = DbHelper.Query<NotifyViewModel>(H2ORepository.ConnectionStringName, sql, new { MemberID = condition.AgentCode }).ToList();
			return result;
		}

		/// <summary>
		/// 取得進度
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		public NotifyViewModel GetCRMESeq(string No)
		{
			string sql = @"select case when t1.CreateTime > t2.CreateTime then '催辦' when  t1.CreateTime < t2.CreateTime then '稽催' end State,
					 case when t1.CreateTime > t2.CreateTime then (select COUNT(*) from CRMEDoS  where no = @No)  
					 when  t1.CreateTime < t2.CreateTime then (select COUNT(*) from CRMEAudit  where no = @No) end Seq
					 from CRMEDoS t1 join CRMEAudit t2 on t1.no = t2.No
					 where t1.no = @No";
			var result = DbHelper.Query<NotifyViewModel>(H2ORepository.ConnectionStringName, sql, new { No = No }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 取得催辦進度
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		public CRMEDoS GetCRMEDoS(string No)
		{
			string sql = @"select top 1 * from CRMEDoS where No = @No	order by CreateTime desc";
			var result = DbHelper.Query<CRMEDoS>(H2ORepository.ConnectionStringName, sql, new { No = No }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 取得稽催進度
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		public CRMEAudit GetCRMEAudit(string No)
		{
			string sql = @"select top 1 * from CRMEAudit where No = @No order by CreateTime desc";
			var result = DbHelper.Query<CRMEAudit>(H2ORepository.ConnectionStringName, sql, new { No = No }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 取得催辦進度次數
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		public int GetCRMEDoSSep(string No)
		{
			string sql = @"select count(*) from CRMEDoS where No = @No";
			var result = DbHelper.Query<int>(H2ORepository.ConnectionStringName, sql, new { No = No }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 取得稽催進度次數
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		public int GetCRMEAuditSep(string No)
		{
			string sql = @"select count(*) from CRMEAudit where No = @No";
			var result = DbHelper.Query<int>(H2ORepository.ConnectionStringName, sql, new { No = No }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 取得客服業務檔案資料
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		public CRMEFile GetCRMEFileByMfid(int ID)
		{
			string sql = @"select * from  CRMEFile where ID = @ID";
			var result = DbHelper.Query<CRMEFile>(H2ORepository.ConnectionStringName, sql, new { ID = ID }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 取得客服業務檔案資料
		/// </summary>
		/// <param name="ID"></param>
		/// <returns></returns>
		public List<CRMEFile> GetCRMEFileLstByMfid(string ID)
		{
			string sql = @"select * from  CRMEFile where ID in (" + ID + ")";
			var result = DbHelper.Query<CRMEFile>(H2ORepository.ConnectionStringName, sql).ToList();
			return result;
		}
		#endregion

		#region 照會單
		/// <summary>
		/// 取得催辦內容
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public CRMEDoS QueryCRMEDoSByNo(string No)
		{
			string sql = @"select top 1 *  from CRMEDoS where No = @No order by ID desc";

			var result = DbHelper.Query<CRMEDoS>(H2ORepository.ConnectionStringName, sql, new { No = No }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 取得催辦內容次數
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public int QueryCRMEDoSCountByNo(string No)
		{
			string sql = @"select COUNT(*)  from CRMEDoS where No = @No";

			var result = DbHelper.Query<int>(H2ORepository.ConnectionStringName, sql, new { No = No }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 取得所有催辦內容
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public List<CRMEDoS> QueryListCRMEDoSByNo(string No)
		{
			string sql = @"select *  from CRMEDoS where No = @No order by ID asc";

			var result = DbHelper.Query<CRMEDoS>(H2ORepository.ConnectionStringName, sql, new { No = No }).ToList();
			return result;
		}

		/// <summary>
		/// 取得稽催內容
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public CRMEAudit QueryCRMEAuditByNo(string No)
		{
			string sql = @"select top 1 *  from CRMEAudit where No = @No order by ID desc";

			var result = DbHelper.Query<CRMEAudit>(H2ORepository.ConnectionStringName, sql, new { No = No }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 取得稽催內容次數
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public int QueryCRMEAuditCountByNo(string No)
		{
			string sql = @"select COUNT(*)  from CRMEAudit where No = @No";

			var result = DbHelper.Query<int>(H2ORepository.ConnectionStringName, sql, new { No = No }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 依Guid取得主表資料
		/// </summary>
		/// <param name="condition"></param>
		/// <returns></returns>
		public CRMECaseContent QueryNotifyDataByGuid(string CaseGuid)
		{
			string sql = @"select *  from CRMECaseContent where CaseGuid = @CaseGuid";

			var result = DbHelper.Query<CRMECaseContent>(H2ORepository.ConnectionStringName, sql, new { CaseGuid = CaseGuid }).FirstOrDefault();
			return result;
		}

		/// <summary>
		/// 依No取得受文者
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public List<CRMENotifyTo> QueryNotifyToByNo(string No)
		{
			string sql = @"select *  from CRMENotifyTo where No = @No and NotifyType in (0,1)";

			var result = DbHelper.Query<CRMENotifyTo>(H2ORepository.ConnectionStringName, sql, new { No = No }).ToList();
			return result;
		}

		/// <summary>
		/// 依No取得受文者
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public List<CRMENotifyTo> GetToByNo(string No)
		{
			string sql = @"select *  from CRMENotifyTo where No = @No and NotifyType in (0)";

			var result = DbHelper.Query<CRMENotifyTo>(H2ORepository.ConnectionStringName, sql, new { No = No }).ToList();
			return result;
		}

		/// <summary>
		/// 依No取得副本受文者
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public List<CRMENotifyTo> GetCCByNo(string No)
		{
			string sql = @"select *  from CRMENotifyTo where No = @No and NotifyType in (1)";

			var result = DbHelper.Query<CRMENotifyTo>(H2ORepository.ConnectionStringName, sql, new { No = No }).ToList();
			return result;
		}

		/// <summary>
		/// 依No取得處代理人
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public List<CRMENotifyTo> GetOMByNo(string No)
		{
			string sql = @"select *  from CRMENotifyTo where No = @No and NotifyType in (4)";

			var result = DbHelper.Query<CRMENotifyTo>(H2ORepository.ConnectionStringName, sql, new { No = No }).ToList();
			return result;
		}

		/// <summary>
		/// 依TypeID取得對應名稱
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public string GetDiscipTypeByID(int ID)
		{
			string sql = @"select Name from CRMEDiscipType where ID = @ID";

			var result = DbHelper.Query<string>(H2ORepository.ConnectionStringName, sql, new { ID = ID }).FirstOrDefault();
			return result;
		}
		#endregion

		#region 新增申訴通知
		/// <summary>
		/// 新增申訴通知
		/// </summary>
		/// <param name="No"></param>
		/// <returns></returns>
		public void InsertCRMEAppealBy(CRMEAppealBy model)
		{
			string sql = @"INSERT INTO CRMEAppealBy(No,AppealName,AppealMobile,AppealMobile_Content,AppealEmail,AppealEmail_Content,EntrustdName,EntrustdMobile,EntrustdMobile_Content,EntrustdEmail,EntrustdEmail_Content,Title,DoUser,DoUserFirstName,DoUserTelExt,Creator,CreateTime)
							Values (@No, @AppealName, @AppealMobile, @AppealMobile_Content, @AppealEmail, @AppealEmail_Content, @EntrustdName, @EntrustdMobile, @EntrustdMobile_Content, @EntrustdEmail, @EntrustdEmail_Content, @Title, @DoUser, @DoUserFirstName,@DoUserTelExt, @Creator, @CreateTime);";
			DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
			{
				No = model.No,
				AppealName = model.AppealName,
				AppealMobile = model.AppealMobile,
				AppealMobile_Content = model.AppealMobile_Content,
				AppealEmail = model.AppealEmail,
				AppealEmail_Content = model.AppealEmail_Content,
				EntrustdName = model.EntrustdName,
				EntrustdMobile = model.EntrustdMobile,
				EntrustdMobile_Content = model.EntrustdMobile_Content,
				EntrustdEmail = model.EntrustdEmail,
				EntrustdEmail_Content = model.EntrustdEmail_Content,
				Title = model.Title,
				DoUser = model.DoUser,
				DoUserFirstName = model.DoUserFirstName,
				DoUserTelExt = model.DoUserTelExt,
				Creator = model.Creator,
				CreateTime = model.CreateTime,
			});

		}
		#endregion


	}
}
