//20250326001-保險公司受理前十大商品年報排程 20250507 by Harrison
using Dapper;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Top10ReportProcess.Base;
using Top10ReportProcess.Model;

namespace Top10ReportProcess
{
    public class DBHelper
    {
        private readonly IDatabaseHelper _dbaseHelper;
        public int CmdTimeout { get; set; } = Convert.ToInt32(ConfigurationManager.AppSettings["CmdTimeout"]);
        public DBHelper(IDatabaseHelper dbHelper)
        {
            _dbaseHelper = dbHelper;
        }

        /// <summary>
        /// 取得核實業績年月
        /// </summary>
        /// <param name="FuncId"></param>
        /// <returns></returns>
        public string GetAgymProductionYm()
        {
            using (var connection = _dbaseHelper.GetConnection("VLIFE"))
            {
                string sql = @"	select top 1  
                CAST(CAST(LEFT(production_ym, 4) AS INT) + 1911 AS VARCHAR) + RIGHT(production_ym, 2) AS production_ym 
	            from agym
	            where 1=1 
	            and agbc_ind = '1'
                and [sequence] = 2
	            and [sequence] <> 88
	            order by production_ym desc,sequence desc";

                return connection.QuerySingleOrDefault<string>(sql, commandTimeout: CmdTimeout);
            }
        }

        /// <summary>
        /// 取得mailaddress
        /// </summary>
        /// <param name="FuncId"></param>
        /// <returns></returns>
        public List<AutoEmailMailaddress> GetMailAddresses(string FuncId)
        {
            using (var connection = _dbaseHelper.GetConnection("VISUALBANCAS_EP"))
            {
                string sql = @"SELECT * 
                FROM [auto_email_mailaddress]
                WHERE func_id = @FuncId";

                return connection.Query<AutoEmailMailaddress>(sql, new { FuncId = FuncId }, commandTimeout: CmdTimeout).ToList();
            }
        }

        /// <summary>
        /// 使用 Dapper 執行 Stored Procedure，並回傳資料清單
        /// </summary>
        /// <typeparam name="T">要回傳的資料型別</typeparam>
        /// <param name="storedProcedureName">Stored Procedure 名稱</param>
        /// <param name="param">輸入參數</param>
        /// <param name="connStr">連線字串代號</param>
        /// <returns></returns>
        public List<T> ExecuteStoredProc<T>(string storedProcedureName, string connStr, DynamicParameters param = null)
        {
            using (var conn = _dbaseHelper.GetConnection(connStr))
            {
                conn.Open();
                return conn.Query<T>(
                    storedProcedureName,
                    param,
                    commandType: CommandType.StoredProcedure,
                    commandTimeout: CmdTimeout
                ).ToList();
            }
        }

        /// <summary>
        /// 使用 Dapper 執行 Stored Procedure，並回傳多個資料清單
        /// </summary>
        /// <param name="storedProcedureName"></param>
        /// <param name="connStr"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public List<List<dynamic>> QueryMultipleDynamic(string storedProcedureName, string connStr, DynamicParameters parameters = null)
        {
            var results = new List<List<dynamic>>();

            using (var conn = _dbaseHelper.GetConnection(connStr))
            {
                conn.Open();

                using (var multi = conn.QueryMultiple(
                    storedProcedureName,
                    parameters,
                    commandType: CommandType.StoredProcedure,
                    commandTimeout: CmdTimeout))
                {
                    while (!multi.IsConsumed)
                    {
                        var table = multi.Read().ToList();
                        results.Add(table);
                    }
                }
            }

            return results;
        }

        /// <summary>
        /// 轉型為強型別
        /// </summary>
        /// <typeparam name="T">傳入Model</typeparam>
        /// <param name="dynamicList"></param>
        /// <returns></returns>
        public List<List<T>> ConvertToTypedList<T>(List<List<dynamic>> dynamicList)
        {
            var typedList = new List<List<T>>();

            foreach (var table in dynamicList)
            {
                var list = new List<T>();

                foreach (var row in table)
                {
                    //  dynamic 轉-> T
                    var json = JsonConvert.SerializeObject(row);
                    var obj = JsonConvert.DeserializeObject<T>(json);
                    list.Add(obj);
                }

                typedList.Add(list);
            }

            return typedList;
        }

        /// <summary>
        /// 不同Model 轉型為強型別
        /// var dynamicList = QueryMultipleDynamic(...); // List<List<dynamic>>
        /// var types = new Type[] { typeof(TopReprotModel), typeof(AnotherModel) };
        /// </summary>
        /// <param name="dynamicList"></param>
        /// <param name="types"></param>
        /// <returns></returns>
        public static List<object> ConvertToTypedList(List<List<dynamic>> dynamicList, Type[] types)
        {
            var result = new List<object>();

            for (int i = 0; i < dynamicList.Count; i++)
            {
                var table = dynamicList[i];
                var list = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(types[i]));

                foreach (var row in table)
                {
                    var json = JsonConvert.SerializeObject(row);
                    var obj = JsonConvert.DeserializeObject(json, types[i]);
                    list.Add(obj);
                }

                result.Add(list);
            }

            return result;
        }




    }
}
