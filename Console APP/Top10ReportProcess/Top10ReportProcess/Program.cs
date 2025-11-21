//20250326001-保險公司受理前十大商品年報排程 20250507 by Harrison
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Top10ReportProcess
{
    public class Program
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        public static void Main(string[] args)
        {
            logger.Info("Start");
            new Process();
            logger.Info("End");
        }
    }
}
