using Dapper;
using SACTAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace SACTAPI.Utilities
{
    public static class LogHelper
    {
        /// <summary>
        /// 寫入SACTAPI_LOG
        /// </summary>
        /// <param name="apiName"></param>
        /// <param name="introducer"></param>
        /// <param name="director"></param>
        /// <param name="registerNo"></param>
        /// <param name="agid"></param>
        /// <param name="responseCode"></param>
        /// <param name="logMsg"></param>
        /// <param name="logRequestData"></param>
        /// <param name="logResponseData"></param>
        public static void SaveLog(SACTAPILog log)
        {
            InsertLog(log);
        }

        /// <summary>
        /// 新增一筆 SACTAPI_LOG
        /// </summary>
        private static void InsertLog(SACTAPILog log)
        {
            using (IDbConnection dbConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["SACT"].ConnectionString))
            {
                string sql = @"
                INSERT INTO SACTAPI_LOG
                (
                    api_name,
                    Introducer,
                    Director,
                    RegisterNo,
                    AGID,
                    response_code,
                    log_msg,
                    log_request_data,
                    log_response_data
                )
                VALUES
                (
                    @ApiName,
                    @Introducer,
                    @Director,
                    @RegisterNo,
                    @AGID,
                    @ResponseCode,
                    @LogMsg,
                    @LogRequestData,
                    @LogResponseData
                );
                ";

                // 執行 SQL 並回傳剛插入的流水號
                dbConnection.Execute(sql, log);
            }
        }

    }
}