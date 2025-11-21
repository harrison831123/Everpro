//20250326001-保險公司受理前十大商品年報排程 20250507 by Harrison
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Top10ReportProcess.Base
{
    public class DatabaseHelper : IDatabaseHelper
    {
        public SqlConnection GetConnection(string connStr)
        {
            string connectionString;
            switch (connStr.ToUpper())
            {
                case "VISUALBANCAS_EP":
                    connectionString = ConfigurationManager.ConnectionStrings["VisualBancas_EP"].ConnectionString;
                    break;
                case "VLIFE":
                    connectionString = ConfigurationManager.ConnectionStrings["VLIFE"].ConnectionString;
                    break;
                default:
                    throw new ArgumentException("無效的資料庫連線代號");
            }

            return new SqlConnection(connectionString);
        }
    }

}
