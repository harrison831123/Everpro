using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PolicyNoteShift
{
	class Program
	{
        private static Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private static void Main(string[] args)
        {
            logger.Info("Start");
            //new TestMail();
            new Process();
            logger.Info("End");
        }
    }
}
