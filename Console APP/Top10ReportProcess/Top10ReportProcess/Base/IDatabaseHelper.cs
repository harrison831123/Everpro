//20250326001-保險公司受理前十大商品年報排程 20250507 by Harrison
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Top10ReportProcess.Base
{
    public interface IDatabaseHelper
    {
        SqlConnection GetConnection(string dbName);
    }

}
