//20250326001-保險公司受理前十大商品年報排程 20250507 by Harrison
using Dapper;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Top10ReportProcess.Base;
using Top10ReportProcess.Model;

namespace Top10ReportProcess
{
    public class Process
    {
        private readonly DBHelper _dbHelper;
        private readonly FileHelper _fileHelper;
        private readonly MailHelper _mailHelper;
        private static Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private static Boolean isSendLog = false;

        public Process()
        {
            //實作
            _dbHelper = new DBHelper(new DatabaseHelper());
            _fileHelper = new FileHelper();
            _mailHelper = new MailHelper();
            AnnualReportProcess();
        }

        /// <summary>
        /// 年報製作
        /// </summary>
        public void AnnualReportProcess()
        {
            try
            {
                logger.Info("AnnualReportProcess==START==");

                logger.Info("取得報表資料");
                var dynamicTables = _dbHelper.QueryMultipleDynamic(
                "usp_Top10AnnualReport",
                "VISUALBANCAS_EP",
                null
                );

                logger.Info("將報表資料轉成強型別TopReprotModel");
                var reportList = _dbHelper.ConvertToTypedList<TopReprotModel>(dynamicTables);

                logger.Info("取得範本檔案路徑");
                string annualReportTemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template", ConfigurationManager.AppSettings["AnnualReportTemplate"]);

                logger.Info("報表製作");             
                string startProductionYm = DateTime.Now.Year.ToString();
                string endProductionYm = _dbHelper.GetAgymProductionYm();
                string lastMonth = DateTime.Now.AddMonths(-1).Month.ToString(); //預收月份               
                string reportName = string.Empty;//EX:202401-2~202411-2核實+12月預收
                //endProductionYm = "202501";
                //endProductionYm = "202512";
                if (endProductionYm.Substring(4, 2) == "12")
                {
                    reportName = startProductionYm + "01-2~" + endProductionYm + "-2核實";
                }
                else if (endProductionYm.Substring(4, 2) == "01")
                {
                    reportName = startProductionYm + "01-2核實+" + lastMonth + "月預收";
                }
                else
                {
                    reportName = startProductionYm + "01-2~" + endProductionYm + "-2核實+" + lastMonth + "月預收";
                }
                logger.Info("報表名稱：" + reportName);

                var reportStream = _fileHelper.AnnualReportByExcelTemplate(annualReportTemplatePath, reportList, reportName, 4);

                logger.Info("取得Mail資訊，FuncId:Top10ReportProcess");
                var emailMailAddress = _dbHelper.GetMailAddresses("Top10ReportProcess");

                logger.Info("Mail發送");
                _mailHelper.SendReportByEmail(
                    reportStream,
                    "保險公司受理前十大產品年報_" + reportName + ".xlsx", 
                    "保險公司受理前十大產品年報_" + reportName,
                    emailMailAddress);
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                isSendLog = true;
            }
            finally
            {
                logger.Info("AnnualReportProcess==END==");
                if (isSendLog)
                {
                    logger.Info("SendLogMail");
                    _mailHelper.SendLogMail();
                }                
            }
        }
    }
}
