using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using NLog;
using System.Configuration;
using System.Linq;
using System.IO;
using MailReportProcess.Model;

namespace MailReportProcess
{
    public class DBHelper
    {
        public string ExecEnv = ConfigurationManager.AppSettings["ENV"];
        public string SysMailAdd = ConfigurationManager.AppSettings["SysMailAdd"];
        public string VISUALBANCAS_EP = ConfigurationManager.ConnectionStrings["VisualBancas_EP"].ConnectionString;
        public string MIS = ConfigurationManager.ConnectionStrings["MIS"].ConnectionString;
        public string VLIFE = ConfigurationManager.ConnectionStrings["VLIFE"].ConnectionString;
        public int execTimeOut = int.Parse(ConfigurationManager.AppSettings["ExecTimeOut"]);
        public string CatchMailAdd = ConfigurationManager.AppSettings["CatchMailAdd"];
        public SqlConnection SqlConn = new SqlConnection();

        public DBHelper()
        {
        }

        public DBHelper(string dbconn) 
        {
            switch (dbconn.ToUpper())
            {
                case "VISUALBANCAS_EP":
                    SqlConn = new SqlConnection(VISUALBANCAS_EP);
                    break;
                case "VLIFE":
                    SqlConn = new SqlConnection(VLIFE);
                    break;
                case "MIS":
                    SqlConn = new SqlConnection(MIS);
                    break;
            }
        }

        public void setParameter(string spparameters, string spparametersmemo,string _batchseq, ref SqlDataAdapter lda_pc)
        {
            string[] paremeterary = spparameters.Split(new char[] { ('|') }, StringSplitOptions.None);
            string[] memoary = spparametersmemo.Split(new char[] { ('|') }, StringSplitOptions.None);
            for (int i = 0; i < paremeterary.Count(); i++)
            {
                string paremetername = paremeterary[i].Trim();
                string memovlaue = memoary[i].Trim();
                switch (paremetername)
                {
                    default:
                        lda_pc.SelectCommand.Parameters.AddWithValue(paremetername, memovlaue);
                        break;
                }
            }
        }

        public DataSet execRptSqlString(string sqldata)
        {
            DataSet rtn_ds = new DataSet();
            using (SqlConn)
            {
                try
                {
                    //01121127001_增加保全照會;優化前端照會系統
                    using (SqlDataAdapter lda_pc = new SqlDataAdapter(sqldata, SqlConn))
                    {
                        lda_pc.SelectCommand.CommandType = CommandType.Text;
                        lda_pc.SelectCommand.CommandTimeout = execTimeOut;
                        lda_pc.Fill(rtn_ds);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(System.Reflection.MethodBase.GetCurrentMethod().Name+"："+ ex.Message);
                }
            }
            return rtn_ds;
        }

        public AutoMailInfo GetAutoMailInfo()
        {
            AutoMailInfo rtnMailInfo = new AutoMailInfo();
            rtnMailInfo.SenderAddress = ConfigurationManager.AppSettings["SenderAddress"];
            rtnMailInfo.SenderDisplayname = ConfigurationManager.AppSettings["SenderDisplayname"];
            rtnMailInfo.SmtpAddress = ConfigurationManager.AppSettings["SmtpAddress"];
            rtnMailInfo.SmtpPort = int.Parse(ConfigurationManager.AppSettings["SmtpPort"]);
            rtnMailInfo.SmtpCredentials = ConfigurationManager.AppSettings["SmtpCredentials"];
            rtnMailInfo.CredentialsId = ConfigurationManager.AppSettings["CredentialsId"];
            rtnMailInfo.CredentialsPwd = ConfigurationManager.AppSettings["CredentialsPwd"];

            return rtnMailInfo;
        }

        public void CreateJobLog(RptInfo rpt)
        {
            string sqlstr = @"Insert Into BatchJobExecLog ([batch_type] ,[exe_name] ,[func_id],[file_name],[batch_datetime]) 
                            values(@batch_type,@exe_name,@func_id,@file_name,@batch_datetime)";
            using (SqlConnection conn = new SqlConnection(VISUALBANCAS_EP))
            {
                using (SqlCommand command = new SqlCommand())
                {
                    command.Connection = conn;
                    command.CommandType = CommandType.Text;
                    command.CommandText = sqlstr;
                    command.Parameters.AddWithValue("@batch_type", "01");
                    command.Parameters.AddWithValue("@exe_name", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name);
                    command.Parameters.AddWithValue("@func_id", rpt.FuncId);
                    command.Parameters.AddWithValue("@file_name", "");
                    command.Parameters.AddWithValue("@batch_datetime", DateTime.Now);
                    command.Connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }
    }

}
