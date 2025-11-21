using Dapper;
using NLog;
using PolicyNoteShift.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PolicyNoteShift
{
	public class DBHelper
	{
        private static Logger logger = NLog.LogManager.GetCurrentClassLogger();
        public string ExecEnv = ConfigurationManager.AppSettings["ENV"];
        public string CmdTimeout = ConfigurationManager.AppSettings["CmdTimeout"];
        public string EPconstr = ConfigurationManager.ConnectionStrings["VisualBancas_EP"].ConnectionString;
        public string CUFconstr = ConfigurationManager.ConnectionStrings["CUF"].ConnectionString;
        public string Vlifeconstr = ConfigurationManager.ConnectionStrings["VLIFE"].ConnectionString;

        /// <summary>
        /// 取得排程Data
        /// </summary>
        /// <returns></returns>
        public List<FileTransInfo> setModelData()
        {
            try
            {
                using (IDbConnection conn = new SqlConnection(EPconstr))
                {
                    return conn.Query<FileTransInfo>("select * from FileTransInfo where FuncId <> ''  ").ToList();
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.StackTrace);
                logger.Error(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 照會資料檢核
        /// </summary>
        /// <param name="notePdfNames"></param>
        /// <returns></returns>
        public List<PbdNoteDataHistory> checkPbdNoteData(string notePdfNames)
        {
            try
            {
                using (var conn = new SqlConnection(EPconstr))
                {
                    DynamicParameters parameters = new DynamicParameters();
                    parameters.Add("notePdfNames", notePdfNames);
                    int Timeout = Convert.ToInt32(CmdTimeout);
                    //Execute stored procedure and map the returned result to a Customer object  
                    return conn.Query<PbdNoteDataHistory>("usp_pbd_checkpolicynote_update", parameters,  commandType: CommandType.StoredProcedure, commandTimeout: Timeout).ToList();
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 確認Half件
        /// </summary>
        /// <param name="notePdfNames"></param>
        /// <returns></returns>
        public string GetAgentId(string policy_no, string content_seq,int type = 1,string noteType = "NB")
        {
            try
            {   //202409 by Fion 全球照會要保書序號為空值，比對保單號碼
                using (var conn = new SqlConnection(EPconstr))
                {
                    DynamicParameters parameters = new DynamicParameters();

                    string sql = string.Empty;
                    if (noteType == "NB") 
                    {
                        parameters.Add("policy_no", policy_no);
                        parameters.Add("content_seq", content_seq);
                        if (type == 1)
                        {
                            //parameters.Add("agent_license_no", agent_license_no);
                            sql = $@"Select pr.agent1_id 
                            from PbdNoteData a WITH(NOLOCK)
                            JOIN policy p WITH(NOLOCK) on a.company_code = p.company_code and a.po_serial = p.po_serial
                            JOIN policyrelate pr WITH(NOLOCK) on p.policy_serial = pr.policy_serial 
                            Where 1=1
                            And a.note_type = 'NB' and a.po_serial <> ''
                            And ISNULL(pr.agent1_id,'') <> '' and a.policy_no = @policy_no and a.content_seq = @content_seq
                            UNION
                            Select pr.agent1_id 
                            from PbdNoteData a WITH(NOLOCK)
                            JOIN policy p WITH(NOLOCK) on a.company_code = p.company_code and a.policy_no = p.policy_no
                            JOIN policyrelate pr WITH(NOLOCK) on p.policy_serial = pr.policy_serial 
                            Where 1=1
                            And a.note_type = 'NB' and a.po_serial = ''
                            And ISNULL(pr.agent1_id,'') <> '' and a.policy_no = @policy_no and a.content_seq = @content_seq
                            ";
                        }
                        else if (type == 2)
                        {
                            parameters.Add("policy_no", policy_no);
                            parameters.Add("content_seq", content_seq);
                            //parameters.Add("agent_license_no", agent_license_no);
                            sql = $@"Select pr.agent2_id 
                            from PbdNoteData a WITH(NOLOCK)
                            JOIN policy p WITH(NOLOCK) on a.company_code = p.company_code and a.po_serial = p.po_serial
                            JOIN policyrelate pr WITH(NOLOCK) on p.policy_serial = pr.policy_serial 
                            Where 1=1
                            And a.note_type = 'NB' and a.po_serial <> ''
                            And ISNULL(pr.agent2_id,'') <> '' and a.policy_no = @policy_no and a.content_seq = @content_seq 
                            UNION
                            Select pr.agent2_id 
                            from PbdNoteData a WITH(NOLOCK)
                            JOIN policy p WITH(NOLOCK) on a.company_code = p.company_code and a.policy_no = p.policy_no
                            JOIN policyrelate pr WITH(NOLOCK) on p.policy_serial = pr.policy_serial 
                            Where 1=1
                            And a.note_type = 'NB' and a.po_serial = ''
                            And ISNULL(pr.agent2_id,'') <> '' and a.policy_no = @policy_no and a.content_seq = @content_seq 
                            ";
                        }
                    }
                    else 
                    {
                        parameters.Add("policy_no", policy_no);
                        parameters.Add("content_seq", content_seq);
                        //保全件只給一位
                        if (type == 1)
                        {
                            sql = $@"Select top 1 agent_code
                            from PbdNoteData a WITH(NOLOCK)
                            JOIN v_cmna v WITH(NOLOCK) on (a.agent_license_no = v.license_no OR a.agent_license_no = v.license_no_nonlife)
                            Where 1=1
                            And a.note_type <> 'NB' and a.policy_no = @policy_no and a.content_seq = @content_seq 
                            ";
                        }else
                        {
                            sql = $@"Select top 1 agent_code
                            from PbdNoteData a WITH(NOLOCK)
                            JOIN v_cmna v WITH(NOLOCK) on (a.agent_license_no = v.license_no OR a.agent_license_no = v.license_no_nonlife)
                            Where 1<>1
                            And a.note_type <> 'NB' and a.policy_no = @policy_no and a.content_seq = @content_seq 
                            ";
                        }
                    }

                    return conn.QueryFirstOrDefault<string>(sql, parameters);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 照會資料檢核
        /// </summary>
        /// <returns></returns>
        public void InsertPbdNoteDataHistory(List<PbdNoteDataHistory> list)
        {
            try
            {
				#region sql
				var sql = @"INSERT INTO [dbo].[PbdNoteDataHistory]
                ([note_type]
                ,[po_serial]
                ,[policy_no]
                ,[notice_date]
                ,[replay_date]
                ,[content_seq]
                ,[agent_license_no]
                ,[note_pdf_name]
                ,[zipfile_name]
                ,[company_code]
                ,[policy_serial]
                ,[result_flag]
                ,[result_desc]
                ,[batch_datetime])
                VALUES
                (@note_type
                , @po_serial
                , @policy_no
                , @notice_date
                , @replay_date
                , @content_seq
                , @agent_license_no
                , @note_pdf_name
                , @zipfile_name
                , @company_code
                , @policy_serial
                , @result_flag
                , @result_desc
                , GETDATE())";
				#endregion
				using (var conn = new SqlConnection(EPconstr))
                {
                    conn.Execute(sql, list);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 新增批次Log檔Layout
        /// </summary>
        /// <returns></returns>
        public void InsertBatchJobExecLog(string FuncId, string FileName)
        {
            try
            {
                #region sql
                var sql = @"Insert INTO BatchJobExecLog(batch_type,exe_name,func_id,file_name,batch_datetime) values('01','PolicyNoteShift',@FuncId,@FileName,GETDATE())";
                #endregion
                using (var conn = new SqlConnection(EPconstr))
                {
                    conn.Execute(sql, new { FuncId = FuncId, FileName = FileName});
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 檔案寫入DB
        /// </summary>
        /// <param name="ldt_import"></param>
        /// <returns></returns>
        public bool dtBulkCopy2DT(DataTable ldt_import)
        {
            bool rtnvalue = false;
            try
            {
                using (SqlBulkCopy lo_sbc = new SqlBulkCopy(EPconstr, SqlBulkCopyOptions.Default))
                {
                    lo_sbc.DestinationTableName = ldt_import.TableName;
                    foreach (DataColumn ldc in ldt_import.Columns)
                    {
                        lo_sbc.ColumnMappings.Add(ldc.ColumnName, ldc.ColumnName);
                    }
                    lo_sbc.WriteToServer(ldt_import);
                }
                rtnvalue = true;
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
            return rtnvalue;
        }

        public void TruncateTable(string tablename)
        {
            try
            {
                using (IDbConnection conn = new SqlConnection(EPconstr))
                {
                    conn.Execute("Truncate Table " + tablename);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 取得新契約掃描序號
        /// </summary>
        /// <param name="agent1_license_no">登錄證字號</param>
        /// <returns></returns>
        public string GetPolicySerial(string agent1_license_no)
        {
            try
            {
                using (IDbConnection conn = new SqlConnection(EPconstr))
                {
                    DynamicParameters parameters = new DynamicParameters();
                    parameters.Add("agent1_license_no", agent1_license_no);

                    string sql = $@"Select p.policy_serial 
                    from PbdNoteData a WITH(NOLOCK)
                    LEFT JOIN policy p WITH(NOLOCK) on a.company_code = p.company_code and a.policy_no = p.policy_no
                    LEFT JOIN policyrelate pr WITH(NOLOCK) on p.policy_serial = pr.policy_serial AND a.agent_license_no = pr.agent1_license_no
                    Where 1=1
                    And a.note_type = 'NB'
                    And ISNULL(pr.agent1_license_no,'')<>'' AND pr.agent1_license_no = @agent1_license_no";

                    return conn.QueryFirstOrDefault<string>(sql, parameters);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 依系統變數名稱回傳設定值
        /// </summary>
        /// <param name="VarName">變數名稱</param>
        /// <returns></returns>
        public string GetConfigValueByName(string VarName)
        {
            try
            {
                using (IDbConnection conn = new SqlConnection(CUFconstr))
                {
                    DynamicParameters parameters = new DynamicParameters();
                    parameters.Add("VarName", VarName);

                    return conn.QueryFirstOrDefault<string>("select VarValue from VarConfig where VarName = @VarName ", parameters);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
        }

        public AutoMailInfo getMailInfo(string programname)
        {
            try
            {
                AutoMailInfo result = new AutoMailInfo();
                using (IDbConnection conn = new SqlConnection(EPconstr))
                {
                    List<AutoMailInfo> oList = conn.Query<AutoMailInfo>(@"
	                            SELECT*
	                            ,(
	                            SELECT mail_address +'|'+mail_displayname+';'
		                            FROM auto_email_mailaddress
	                            where func_id =@funcid and from_to_type ='1'
	                            FOR XML PATH('')
	                            ) as mail_addr_TO
	                            ,(
	                            SELECT mail_address +'|'+mail_displayname+';'
		                            FROM auto_email_mailaddress
	                            where func_id =@funcid and from_to_type ='2'
	                            FOR XML PATH('')
	                            ) as mail_addr_CC
	                            ,(
	                            SELECT mail_address +'|'+mail_displayname+';'
		                            FROM auto_email_mailaddress
	                            where func_id =@funcid and from_to_type not in ('1','2')
	                            FOR XML PATH('')
	                            ) as mail_addr_BCC	
	                            FROM auto_email
	                            where func_id =@funcid
		                    ", new { funcid = programname }).ToList();

                    if (oList.Count == 1)
                    {
                        result = oList[0];
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 西元日期轉民國日期
        /// </summary>
        /// <param name="date">日期</param>
        /// <returns></returns>
        public string WDateToCDate(string date)
        {
            int year = Convert.ToInt32(date.Substring(0, 4));
            int month = Convert.ToInt32(date.Substring(4, 2));
            int day = Convert.ToInt32(date.Substring(6, 2));
            DateTime dt = new DateTime(year, month, day);
            CultureInfo culture = new CultureInfo("zh-TW");
            culture.DateTimeFormat.Calendar = new TaiwanCalendar();            
            return dt.ToString("0"+"yyyMMdd", culture);
        }
    }
}
