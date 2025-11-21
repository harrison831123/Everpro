using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using System.Xml.Serialization;
using System.IO;
using MailReportProcess.Model;
using NPOI;
using NPOI.HSSF.UserModel;
using System.Data;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Core;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using NPOI.SS.UserModel;

namespace MailReportProcess
{
    public class Program
    {

        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private static readonly string ExecENV = new DBHelper().ExecEnv;
        private static readonly string SysMailADD = new DBHelper().SysMailAdd;
        private static readonly string CatchMailADD = new DBHelper().CatchMailAdd;
        private static Boolean isSendLog = false;

        public static void Main(string[] args)
        {
            Program program = new Program();
            XmlSerializer serializer = new XmlSerializer(typeof(MailRpt));
            MailRpt mailRpt;
            logger.Info("MailReportProcess==START==");
            try
            {
                //取得SMTP資訊
                AutoMailInfo autoMailInfo = new DBHelper().GetAutoMailInfo();

                // 建立FileStream物件，讀取XML檔案
                using (FileStream fs = new FileStream(System.Environment.CurrentDirectory + "/RptInfo.xml", FileMode.Open))
                {
                    //反序列化XML
                    mailRpt = (MailRpt)serializer.Deserialize(fs);
                }

                foreach (RptInfo rptInfo in mailRpt.RptInfo)
                {
                    logger.Info($"==FuncId：{rptInfo.FuncId}；ExecuteType：{rptInfo.ExecuteType}；ExecuteDay：{rptInfo.ExecuteDay}；ExecuteTime：{rptInfo.ExecuteTime}==S==");
                    //檢核是否為執行時間
                    if (!program.isExecRptInfo(rptInfo))
                    {
                        continue;
                    }

                    //執行報表SQL(SP)
                    DBHelper dBHelper = new DBHelper(rptInfo.DbConn);
                    DataSet ds = new DataSet();
                    try
                    {
                        ds = dBHelper.execRptSqlString(rptInfo.Sqlstring);
                    }
                    catch(Exception e)
                    {
                        logger.Error(e.Message);
                        isSendLog = true;
                    }

                    //回傳DataSet中DataTable數量要與RptDetail設定數相同
                    if (rptInfo.RptDetail.Count != ds.Tables.Count)
                    {
                        logger.Error("RptDetail數量與資料表不符");
                        isSendLog = true;
                        continue;
                    }

                    List<string> conditionKeylist = rptInfo.RptKey.Split(',').ToList();
                    //ds.Tables[0]為MailList(報表smtp/email add資訊)，不產出
                    DataTable dtMailList = ds.Tables[0];

                    //報表資料集合
                    Dictionary<RptDetail, DataView> oDic = new Dictionary<RptDetail, DataView>();
                    string strRptTitle = "";
                    string strOriFileName = rptInfo.FileName;
                
                    if (dtMailList.Rows.Count == 0)
                    {
                        throw new Exception("dtMailList無資料");
                    }

                    for (int i = 0; i < dtMailList.Rows.Count; i++)
                    {
                        oDic.Clear();
                        DataRow dataMailRow = dtMailList.Rows[i];
                        string strCondition = "";
                        //依rptInfo.RptKey取得比對條件，報表datatable都要有rptInfo.RptKey的欄位
                        foreach (string conditionkey in conditionKeylist)
                        {
                            if (strCondition != "")
                            {
                                strCondition = strCondition + " AND " + conditionkey + " = '" + dataMailRow[conditionkey].ToString() + "'";
                            }
                            else
                            {
                                strCondition = conditionkey + " = '" + dataMailRow[conditionkey].ToString() + "'";
                            }
                        }

                        strRptTitle = dataMailRow["rpt_title"].ToString();

                        //報表資料從回傳DataSet中DataTable[1]開始
                        for (int j = 1; j < ds.Tables.Count; j++)
                        {
                            //依strCondition篩選DataTable資料，多個DataTable為一xls檔案，有不同sheet
                            DataView dataView = ds.Tables[j].DefaultView;
                            dataView.RowFilter = strCondition;
                            oDic.Add(rptInfo.RptDetail[j], dataView);
                        }
                        rptInfo.FileName = String.Format(strOriFileName, strRptTitle, (DateTime.Now.ToString("yyyyMMdd")));
                    
                        //產生檔案&寄信
                        using (MemoryStream outputms = program.GetReportMemoryStream(rptInfo, oDic))
                        {
                            //SmtpClient client = new SmtpClient();
                            using (SmtpClient client = new SmtpClient())
                            {
                            // 發送email
                            using (MailMessage mm = new MailMessage())
                            {
                                mm.From = new MailAddress(autoMailInfo.SenderAddress, autoMailInfo.SenderDisplayname);
                                try
                                {
                                        if (string.IsNullOrEmpty(dataMailRow["email_add"].ToString().Trim()))
                                        {
                                            throw new Exception();
                                        }
                                    //多個EMAIL處理方式
                                    if (dataMailRow["email_add"].ToString().IndexOf('|') >= 0)
                                    {
                                        string[] mailto = dataMailRow["email_add"].ToString().Trim().Split(new char[] { (',') }, StringSplitOptions.RemoveEmptyEntries);
                                        for (int m = 0; m < mailto.Length; m++)
                                        {
                                                mm.To.Add(new MailAddress(mailto[m].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[0]
                                                        , mailto[m].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[1]));
                                        }
                                    }
                                    else
                                    {
                                        mm.To.Add(new MailAddress(dataMailRow["email_add"].ToString().Trim(), dataMailRow["rpt_title"].ToString(), System.Text.Encoding.UTF8));
                                    }
                                }
                                catch (Exception)
                                {
                                    logger.Warn(string.Format("{0}-MailAddressError!!",dataMailRow["rpt_title"].ToString()));
                                        mm.To.Add(new MailAddress(CatchMailADD));
                                }

                                mm.Subject = String.Format(rptInfo.Title, strRptTitle, DateTime.Now.ToString("yyyyMMdd"));
                                if (ExecENV == "T")
                                {
                                    mm.Subject = "[TEST]" + String.Format(rptInfo.Title, strRptTitle, DateTime.Now.ToString("yyyyMMdd"));
                                }
                                mm.SubjectEncoding = System.Text.Encoding.UTF8;
                                mm.Body = "系統自動發送，請勿回信!";
                                mm.BodyEncoding = System.Text.Encoding.UTF8;
                                mm.IsBodyHtml = true;
                                mm.Priority = System.Net.Mail.MailPriority.High;

                                if (outputms.Length > 0)
                                {
                                    ContentType contentype = new ContentType();
                                    contentype.MediaType = MediaTypeNames.Application.Octet;
                                    contentype.Name = String.Format(rptInfo.ZipName, strRptTitle, DateTime.Now.ToString("yyyyMMdd"));
                                    Attachment att = new Attachment(outputms, contentype);
                                    mm.Attachments.Add(att);
                                }

                                System.Net.Mail.SmtpClient lo_smtp = new System.Net.Mail.SmtpClient(autoMailInfo.SmtpAddress);
                                lo_smtp.Port = autoMailInfo.SmtpPort; //設定Port
                                lo_smtp.UseDefaultCredentials = true;
                                lo_smtp.EnableSsl = true; //gmail預設開啟驗證

                                // Check Sender Mail ID and PWD
                                if (autoMailInfo.SmtpCredentials == "Y")
                                {
                                    lo_smtp.Credentials = new NetworkCredential(autoMailInfo.CredentialsId, autoMailInfo.CredentialsPwd);
                                }
                                lo_smtp.Send(mm);
                            }
                            }
                        }
                    }
                    logger.Info($"==FuncId：{rptInfo.FuncId}；ExecuteType：{rptInfo.ExecuteType}；ExecuteDay：{rptInfo.ExecuteDay}；ExecuteTime：{rptInfo.ExecuteTime}==E==");
                    //寫入批次執行記錄檔
                    dBHelper.CreateJobLog(rptInfo);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                isSendLog = true;
            }
            logger.Info("MailReportProcess==END==");

            if (isSendLog)
            {
                program.SendLogMail();
            }
        }

        /// <summary>
        /// 依設定資料檢核執行時間
        /// </summary>
        /// <param name="rptInfo"></param>
        /// <returns></returns>
        private bool isExecRptInfo(RptInfo rptInfo)
        {
            bool rtnval = true;
            string execHH = DateTime.Now.ToString("HH");
            string execDD = DateTime.Now.ToString("dd");
            string execWD = ((int)DateTime.Now.DayOfWeek).ToString();
            //99設定為一定執行，for測試/rerun使用
            if (rptInfo.ExecuteTime == "99")
            {
                rptInfo.ExecuteType = "D";
                rptInfo.ExecuteTime = execHH;
            }

            if (rptInfo.ExecuteType != "D" && rptInfo.ExecuteDay == "")
            {
                logger.Error($"未設定ExecuteDay!!");
                rtnval = false;
            }

            switch (rptInfo.ExecuteType)
            {
                case "M": //每月
                    rtnval = (rptInfo.ExecuteDay == execDD) ? true : false;
                    break;
                case "W": //每週
                    rtnval = (rptInfo.ExecuteDay == execWD) ? true : false;
                    break;
            }

            if (rptInfo.VaildFlag.ToUpper() != "Y" || rptInfo.ExecuteTime != execHH)
            {
                rtnval = false;
            }

            return rtnval;
        }

        /// <summary>
        /// 產生報表及壓縮檔案
        /// </summary>
        /// <param name="rptInfo">報表zip檔資訊</param>
        /// <param name="oDic">sheet名稱及報表資料的集合</param>
        /// <returns>MemoryStream</returns>
        public MemoryStream GetReportMemoryStream(RptInfo rptInfo, Dictionary<RptDetail, DataView> oDic)
        {
            MemoryStream ms = new MemoryStream();
            MemoryStream outputms = new MemoryStream();

            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet;
            HSSFRow row;
            Int32 rownumber = 0;
            MemoryStream _msexcel = new MemoryStream();
            HSSFCell cell;
            HSSFCellStyle borderstyle = (HSSFCellStyle)wb.CreateCellStyle();
            HSSFCellStyle green = (HSSFCellStyle)wb.CreateCellStyle();
            HSSFCellStyle red = (HSSFCellStyle)wb.CreateCellStyle();
            HSSFCellStyle headstyle = (HSSFCellStyle)wb.CreateCellStyle();

            try
            {
                foreach (var item in oDic)
                {
                    RptDetail rptDetail = item.Key;
                    DataView dataView = item.Value;

                    //產生報表頁籤
                    sheet = (HSSFSheet)wb.CreateSheet(rptDetail.RptName);
                    rownumber = 0;
                    row = (HSSFRow)sheet.CreateRow(rownumber);
                    DataTable dataTable = dataView.ToTable();
                    List<int> lsHeadColumnWidth = new List<int>();

                    // 產生報表表頭欄位說明    
                    foreach (DataColumn col in dataTable.Columns)
                    {
                        // 若是欄位名稱設定在xml的rpt_key中不處理，因為該欄位是用來判斷資料
                        if (!rptInfo.RptKey.Contains(col.ColumnName))
                        {
                            cell = (HSSFCell)row.CreateCell(col.Ordinal);
                            cell.SetCellValue(col.ColumnName);
                            headstyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightYellow.Index;
                            headstyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
                            cell.CellStyle = headstyle;
                            lsHeadColumnWidth.Add(Encoding.UTF8.GetBytes(cell.ToString()).Length);
                        }
                    }

                    // 產生報表資料
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        rownumber++;
                        row = (HSSFRow)sheet.CreateRow(rownumber);
                        foreach (DataColumn col in dataTable.Columns)
                        {
                            // 若是欄位名稱設定在xml的rpt_key中不處理，因為該欄位是用來判斷資料
                            if (!rptInfo.RptKey.Contains(col.ColumnName))
                            {
                                HSSFCell cells = (HSSFCell)row.CreateCell(col.Ordinal);
                                cells.CellStyle = borderstyle;
                                if (col.DataType.FullName == "System.String")
                                {

                                    if (col.ColumnName.Equals("被保人"))
                                    {
                                        cells.SetCellType(NPOI.SS.UserModel.CellType.String);
                                    }
                                    else
                                    {
                                        cells.SetCellType(NPOI.SS.UserModel.CellType.String);
                                        cells.SetCellValue(dataTable.Rows[i][col].ToString());
                                    }
                                }
                                else if (col.DataType.FullName == "System.Int32")
                                {
                                    cells.SetCellType(NPOI.SS.UserModel.CellType.Numeric);
                                    cells.SetCellValue(Convert.ToInt32(dataTable.Rows[i][col]));
                                }
                                else
                                {
                                    cells.SetCellType(NPOI.SS.UserModel.CellType.String);
                                    cells.SetCellValue(dataTable.Rows[i][col].ToString());
                                }
                            }
                        }
                        //設置自動調整欄寬
                        for (int columnNum = 0; columnNum < dataTable.Columns.Count; columnNum++)
                        {
                            int columnWidth = sheet.GetColumnWidth(columnNum) / 256;
                            for (int rowNum = 1; rowNum <= sheet.LastRowNum; rowNum++)
                            {
                                IRow currentRow = sheet.GetRow(rowNum);
                                if (currentRow.GetCell(columnNum) != null)
                                {
                                    ICell currentCell = currentRow.GetCell(columnNum);
                                    int length = Encoding.UTF8.GetBytes(currentCell.ToString()).Length;
                                    //先比對報表表頭欄寬
                                    if (columnWidth < lsHeadColumnWidth[columnNum])
                                    {
                                        columnWidth = lsHeadColumnWidth[columnNum];
                                    }
                                    if (columnWidth < length +1 )
                                    {
                                        columnWidth = length +1 ;
                                    }
                                }
                            }
                            sheet.SetColumnWidth(columnNum, columnWidth * 230);
                        }  
                    }
                }
                wb.Write(ms);
                ms.Flush();
                ms.Position = 0;
            
                // 檔案壓縮加密處理           
                // 被壓縮的檔案的名稱的編碼，若是沒指定有可能會發生亂碼的情況。
                Encoding ZFNEncoding = Encoding.GetEncoding("BIG5");
                ZipOutputStream zipstream = new ZipOutputStream(outputms);
                zipstream.Password = rptInfo.ZipPw;
                ZipEntry _zipentry = new ZipEntry(rptInfo.FileName);
                _zipentry.DateTime = DateTime.Now;
                zipstream.PutNextEntry(_zipentry);
                StreamUtils.Copy(ms, zipstream, new Byte[4096]);
                zipstream.CloseEntry();
                zipstream.IsStreamOwner = false;
                zipstream.Close();
                outputms.Position = 0;
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw;
            }
            return outputms;
        }

        public void SendLogMail(String FundID ="")
        {
            var logfile = (LogManager.Configuration.FindTargetByName("File") as NLog.Targets.FileTarget).FileName
                .Render(new LogEventInfo() { TimeStamp = DateTime.Now, LoggerName = "loggerName" });

            (LogManager.Configuration.FindTargetByName("File") as NLog.Targets.FileTarget).KeepFileOpen = false;
            SmtpClient client = new SmtpClient();
            try
            {
                AutoMailInfo Logmailinfo = new DBHelper().GetAutoMailInfo();
                using (MailMessage lo_mm = new MailMessage())
                {
                    lo_mm.From = new MailAddress(Logmailinfo.SenderAddress, Logmailinfo.SenderDisplayname);
                    lo_mm.Subject = String.Format("MailReportProcess_異常通知{0}", FundID);
                    if (ExecENV == "T")
                    {
                        lo_mm.Subject = "[TEST]" + lo_mm.Subject;
                        lo_mm.To.Add(new MailAddress("fioncheng@mail.everprobks.com.tw"));
                    }
                    else
                    {
                        lo_mm.To.Add(new MailAddress(SysMailADD));
                    }

                    lo_mm.SubjectEncoding = System.Text.Encoding.UTF8;
                    lo_mm.BodyEncoding = System.Text.Encoding.UTF8;
                    lo_mm.IsBodyHtml = true;
                    lo_mm.Priority = System.Net.Mail.MailPriority.High;
                    lo_mm.Attachments.Add(new Attachment(logfile));

                    client.Host = Logmailinfo.SmtpAddress; //設定smtp Server
                    client.Port = Logmailinfo.SmtpPort; //設定Port
                    client.UseDefaultCredentials = true;
                    client.EnableSsl = true; //gmail預設開啟驗證

                    // Check Sender Mail ID and PWD
                    if (Logmailinfo.SmtpCredentials == "Y")
                    {
                        client.Credentials = new NetworkCredential(Logmailinfo.CredentialsId, Logmailinfo.CredentialsPwd);
                    }
                    client.Send(lo_mm);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
            }
            finally
            {
                client.Dispose();
            }
        }
    }
}
