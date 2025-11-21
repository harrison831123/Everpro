using EP.H2OModels;
using EP.SD.SalesSupport.LAW.Models;
using EP.SD.SalesSupport.LAW.Service;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Controllers
{
    [Program("LAWQU003")]
    public class LAWQU003Controller : BaseController
    {
        // GET: LAW/LAWQU003
        private ILAWService _Service;
        public LAWQU003Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }

        /// <summary>
        /// 報表作業
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU003")]
        public ActionResult Index()
        {
            return View();
        }
        #region 法追統計
        [HttpGet]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU003")]
        public ActionResult _Statistics()
        {
            return PartialView("_Statistics");
        }

        /// <summary>
        /// 法追統計
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public JsonResult _Statistics(string year)
        {
            int set_year, y_str;
            string year_str = string.Empty;
            //set_year = (Convert.ToInt32(year) + 1911);
            set_year = Convert.ToInt32(year);
            List<StatisticsDetail> statistics = new List<StatisticsDetail>();

            for (int i = 95; i <= set_year; i++)
            {
                year_str = i == 95 ? "95" : year_str + "," + i;
            }
            y_str = DateTime.Now.Year - 1911;
            for (int i = set_year; i <= y_str; i++)
            {
                StatisticsDetail model = new StatisticsDetail();
                LawMasterReportLog reportLog = new LawMasterReportLog();
                LawContent content = new LawContent();
                LawRepaymentList repaymentList = new LawRepaymentList();
                model.year = year;
                if (i == set_year)
                {
                    reportLog = _Service.GetLawMasterReportLogByLawyear(y_str, year_str);
                    content = _Service.GetLawcontentByLawyears(year_str);
                    repaymentList = _Service.GetSumLawRepaymentList(year_str);
                    model.sumtype = 1;
                }
                else
                {
                    reportLog = _Service.GetLawMasterReportLogOtherByLawyear(y_str, i);
                    content = _Service.GetLawcontentByLawyear(i);
                    repaymentList = _Service.GetLawRepaymentList(i);
                    model.sumtype = 0;
                }

                if (reportLog != null)
                {
                    model.DateNow = (DateTime.Now.Year - 1911) + DateTime.Now.ToString("MMdd");
                    model.DateHave = (Convert.ToInt32(reportLog.CreateDate.Substring(0, 4)) - 1911).ToString() + reportLog.CreateDate.Substring(5, 2) + reportLog.CreateDate.Substring(8, 2);
                    model.Lawyear = reportLog.LawYear;
				}
				else
				{
                    model.DateNow = (DateTime.Now.Year - 1911) + DateTime.Now.ToString("MMdd");
                    model.DateHave = (DateTime.Now.Year - 1911) + DateTime.Now.ToString("MMdd");
                    model.Lawyear = y_str.ToString();
                }

                if (content != null)
                {
                    if (reportLog != null)
                    {
                        model.LawDueMoney = content.LawDueMoney;
                        model.LawTotalDue = reportLog.LawTotalDue;
                    }
                    else
                    {
                        model.LawDueMoney = content.LawDueMoney;
                    }
                }
                else
                {
                    model.LawDueMoney = 0;
                }

                if (repaymentList != null)
                {
                    if (reportLog != null)
                    {
                        model.LawTotalRepayment = reportLog.LawTotalRepayment;
                        model.LawRepaymentMoney = repaymentList.LawRepaymentMoney;
                    }
                    else
                    {
                        model.LawRepaymentMoney = repaymentList.LawRepaymentMoney;
                    }
                }
                else
                {
                    model.LawRepaymentMoney = 0;
                }

                if (model.LawDueMoney != 0 && model.LawRepaymentMoney != 0)
                {
                    decimal decimalValue = (model.LawRepaymentMoney / model.LawDueMoney) * 100;
                    model.pstr = Math.Round(decimalValue, 2) + "%";
                }
                else
                {
                    if (model.LawDueMoney != 0 && model.LawRepaymentMoney == 0)
                    {
                        model.pstr = "0%";
                    }
                    else
                    {
                        model.pstr = "";
                    }
                }

                if (reportLog != null)
                {
                    if (i == set_year)
                    {
                        decimal decimalValue = Convert.ToDecimal(reportLog.LawTotalRepayment) / Convert.ToDecimal(reportLog.LawTotalDue) * 100;
                        model.repayperstr = Math.Round(decimalValue, 2) + "%";
                    }
                    else
                    {
                        model.repayperstr = reportLog.LawRepayPercent;
                    }
                }
                else
                {
                    model.repayperstr = "";
                }

                statistics.Add(model);
            }

            //return PartialView("_Statistics", statistics);
            return Json(statistics, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 法追統計-系統紀錄log
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public void InsertLawMasterReportLog()
        {
            bool result = _Service.InsertLawMasterReportLog(User.MemberInfo.ID);
            if (result)
            {
                AppendMessage(PlatformResources.新增成功);
            }
            else
            {
                AppendMessage(PlatformResources.新增失敗);
            }
        }

        ///// <summary>
        ///// 依照條件抓取報表
        ///// </summary>
        ///// <param name="productionY">業績年</param>
        //[HttpPost]
        //public JsonResult GetReport(StatisticsDetail model)
        //{
        //    string fileName = string.Empty;
        //    byte[] data = null;
        //    var service = ServiceHelper.Create<ILAWService>();
        //    string datett = (DateTime.Now.Year - 1911) + DateTime.Now.ToString("MMdd");
        //    if (model.LawYearType == "95")
        //    {
        //        fileName = "法追系統-當月(" + datett + ")總表" + ".xls";
        //        data = service.GetStatisticReportList(model);
        //    }
        //    else
        //    {
        //        fileName = "法追系統-當月(" + datett + ")總表_年度合併" + ".xls";
        //        data = service.GetSumStatisticReportList(model);
        //    }

        //    string handle = Guid.NewGuid().ToString();
        //    TempData[handle] = data;
        //    return new JsonResult()
        //    {
        //        Data = new
        //        {
        //            FileGuid = handle
        //            ,
        //            FileName = fileName
        //        }
        //    };
        //}

        ///// <summary>
        ///// 下載輸出
        ///// </summary>
        ///// <param name="fileGuid">guid</param>
        ///// <param name="fileName">檔名</param>
        ///// <returns></returns>
        //[HttpGet]
        //public virtual ActionResult Download(string fileGuid, string fileName)
        //{
        //    if (TempData[fileGuid] != null)
        //    {
        //        byte[] data = TempData[fileGuid] as byte[];
        //        //return File(data, "application/vnd.ms-excel", fileName);
        //        return File(data, MimeMapping.GetMimeMapping(fileName), fileName);
        //    }
        //    else
        //    {
        //        // Problem - Log the error, generate a blank file,
        //        return new EmptyResult();
        //    }
        //}

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="selyear">年度</param>
        /// <param name="selyear">月份</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult GetReport(string LawYearType)
        {
            var service = ServiceHelper.Create<ILAWService>();
            string datett = (DateTime.Now.Year - 1911) + DateTime.Now.ToString("MMdd");
            string fileName = LawYearType == "95" ? "法追系統-當月(" + datett + ")總表" + ".xlsx" : "法追系統-當月(" + datett + ")總表_年度合併" + ".xlsx";
            var ms = LawYearType == "95" ? service.GetStatisticReportList(LawYearType) : service.GetSumStatisticReportList(LawYearType);
            var filename = Url.Encode(fileName);
            return File(ms, "application/octet-estream", filename);
        }
        #endregion

        #region  當月還款明細
        [HttpGet]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU003")]
        public ActionResult _MonthRepaymentReport()
        {
            return PartialView("_MonthRepaymentReport");
        }

        /// <summary>
        /// 查詢作業
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public JsonResult Query(LawMonthRepaymentReportDetail model)
        {
            //取資料
            var mService = new WebChannel<ILAWService>();
            List<LawMonthRepaymentReportDetail> Viewmodel = new List<LawMonthRepaymentReportDetail>();

            mService.Use(service => service
            .GetLawMonthRepaymentReport(model)
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new LawMonthRepaymentReportDetail();
                    item.LawId = d.LawId;
                    item.vmname = d.vmname;
                    item.smname = d.smname;
                    item.name = d.name;
                    item.LawRepaymentMoney = d.LawRepaymentMoney;
                    LawRepaymentList lawRepayment = _Service.GetSumLawRepaymentMoney(model);
                    item.LawSumRepaymentMoney = lawRepayment != null ? lawRepayment.LawRepaymentMoney : 0;
                    Viewmodel.Add(item);
                }
            }));

            return Json(Viewmodel, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="selyear">年度</param>
        /// <param name="selyear">月份</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult QueryMonthReport(string selyear, string selmonth, string chkm)
        {
            var service = ServiceHelper.Create<ILAWService>();
            var filename = Url.Encode(string.Format("法追系統-當月({0}年{1}月)還款明細(含續佣抵結欠).xlsx", Convert.ToInt32(selyear) - 1911, selmonth));

            //try 
            //{
            var ms = service.QueryLawMonthRepaymentReport(selyear, selmonth, chkm);
            return File(ms, "application/octet-estream", filename);
            //}
            //catch(Exception ex)
            //{
            //    AppendMessage("查詢無資料");
            //    return null;
            //}            
        }

        #endregion

        #region  團隊明細
        [HttpGet]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU003")]
        public ActionResult _TeamListReport()
        {
            return PartialView("_TeamListReport");
        }

        /// <summary>
        /// 團隊體系
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public JsonResult _TeamListReport(string year)
        {
            int set_year, y_str;
            string year_str = string.Empty;
            set_year = Convert.ToInt32(year);
            List<LawVmSmReport> lawVmSmReports = new List<LawVmSmReport>();
            List<LawVmSmReport> viewmodel = new List<LawVmSmReport>();
            LawVmSmReport model = new LawVmSmReport();
            model.LawYear = year;
            lawVmSmReports = _Service.GetLawVmSmReport(model);

            //for (int i = 95; i <= set_year; i++)
            //{
            //    year_str = i == 95 ? "95" : year_str + "," + i;
            //}
            //y_str = DateTime.Now.Year - 1911;


            //for (int i = set_year; i <= y_str; i++)
            //{
            //    LawVmSmReport item = new LawVmSmReport();
            //    lawVmSmReports[j].LawYear = i.ToString();
            //    if (i == set_year)
            //    {
            //        item = _Service.GetSumLawVmSmReportMoney(lawVmSmReports[j], year_str);
            //        item.VmName = lawVmSmReports[j].VmName;
            //        item.SmName = lawVmSmReports[j].SmName;
            //        item.LawYear = lawVmSmReports[j].LawYear;
            //    }
            //    else
            //    {
            //        item = _Service.GetLawVmSmReportMoney(lawVmSmReports[j]);
            //    }
            //}
            //viewmodel.Add(item);

            return Json(lawVmSmReports, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 團隊體系資料
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public JsonResult TeamListData(LawVmSmReport model)
        {
            int set_year, y_str;
            string year_str = string.Empty;
            set_year = Convert.ToInt32(model.LawYear);
            List<LawVmSmDetail> lawVmSmReports = new List<LawVmSmDetail>();
            for (int i = 95; i <= set_year; i++)
            {
                year_str = i == 95 ? "95" : year_str + "," + i;
            }
            y_str = DateTime.Now.Year - 1911;

            for (int i = set_year; i <= y_str; i++)
            {
                LawVmSmDetail item = new LawVmSmDetail();
                if (i == set_year)
                {
                    item = _Service.GetSumLawVmSmReportMoney(model, year_str);
                }
                else
                {
                    model.LawYear = i.ToString();
                    item = _Service.GetLawVmSmReportMoney(model);
                }
                if (item.RepayMoney == "0")
                {
                    item.pstr = string.Empty;
                }
                else
                {
                    decimal decimalValue = Convert.ToDecimal(item.RepayMoney) / Convert.ToDecimal(item.DueMoney) * 100;
                    item.pstr = Math.Round(decimalValue, 2).ToString();
                }
                item.DueMoney = item.DueMoney == null ? "0" : item.DueMoney;
                lawVmSmReports.Add(item);
            }
            return Json(lawVmSmReports, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="selyear">年度</param>
        /// <param name="selyear">月份</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult TeamReport(string LawYearType)
        {
            var service = ServiceHelper.Create<ILAWService>();
            var ms = service.QueryTeamReport(LawYearType);
            var filename = Url.Encode("團隊明細報表.xlsx");
            return File(ms, "application/octet-estream", filename);
        }
        #endregion

        #region  案件統計
        [HttpGet]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU003")]
        public ActionResult _YearReport()
        {
            return PartialView("_YearReport");
        }

        /// <summary>
        /// 查詢作業
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public JsonResult CaseQuery()
        {
            //取資料           
            Tuple<List<int>, List<int>, List<int>, List<int>, List<int>> Info = TupleGetInfo();

            return Json(Info, JsonRequestBehavior.AllowGet);
        }

        public static Tuple<List<int>, List<int>, List<int>, List<int>, List<int>> TupleGetInfo()
        {
            string mm;
            List<int> TopYear = new List<int>();
            List<int> NowYear = new List<int>();
            List<int> SumType = new List<int>();
            List<int> Outside = new List<int>();
            List<int> OwnSelf = new List<int>();
            var service = ServiceHelper.Create<ILAWService>();
            for (int i = 1; i <= 12; i++)
            {
                if (i.ToString().Length == 1)
                {
                    mm = "0" + i;
                }
                else
                {
                    mm = i.ToString();
                }
                List<LawContent> Top = service.GetLawTopYearReport();
                List<LawContent> Now = service.GetLawNowYearReport(mm);
                List<LawContent> Sum = service.GetLawSumNotTypeReport(mm);
                List<LawContent> Out = service.GetLawOutSideReport(mm);
                List<LawContent> Own = service.GetLawSelfReport(mm);

                TopYear.Add(Top.Count);
                NowYear.Add(Now.Count);
                SumType.Add(Sum.Count);
                Outside.Add(Out.Count);
                OwnSelf.Add(Own.Count);
            }

            return new Tuple<List<int>, List<int>, List<int>, List<int>, List<int>>(TopYear, NowYear, SumType, Outside, OwnSelf);
        }

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="selyear">年度</param>
        /// <param name="selyear">月份</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult CaseReport()
        {
            var service = ServiceHelper.Create<ILAWService>();
            var ms = service.QueryLawYearReport();
            var filename = Url.Encode("95-98年度案件統計表.xlsx");
            return File(ms, "application/octet-estream", filename);
        }
        #endregion

        #region  案件統計
        [HttpGet]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU003")]
        public ActionResult _NowYearReport()
        {
            return PartialView("_NowYearReport");
        }


        /// <summary>
        /// 查詢作業
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public JsonResult NowYearQuery(string year)
        {
            //取資料           
            Tuple<List<int>, List<int>, List<int>, List<int>, List<int>, List<int>> Info = NowYearTuple(year);

            return Json(Info, JsonRequestBehavior.AllowGet);
        }

        public static Tuple<List<int>, List<int>, List<int>, List<int>, List<int>, List<int>> NowYearTuple(string year)
        {
            string mm;
            List<int> Accpets = new List<int>();
            List<int> Outsources = new List<int>();
            List<int> Doselfs = new List<int>();
            List<int> NotOCs = new List<int>();
            List<int> OCs = new List<int>();
            List<int> SumTotal = new List<int>();
            var service = ServiceHelper.Create<ILAWService>();
            for (int i = 1; i <= 12; i++)
            {
                if (i.ToString().Length == 1)
                {
                    mm = "0" + i;
                }
                else
                {
                    mm = i.ToString();
                }
                List<LawContent> Accpet = service.GetLawAccpet(year, mm);
                List<LawContent> Outsource = service.GetLawOutsource(year, mm);
                List<LawContent> Doself = service.GetLawDoself(year, mm);
                List<LawContent> NotOC = service.GetLawNotOC(year, mm);
                List<LawContent> OC = service.GetLawOC(year, mm);

                Accpets.Add(Accpet.Count);
                Outsources.Add(Outsource.Count);
                Doselfs.Add(Doself.Count);
                NotOCs.Add(NotOC.Count);
                OCs.Add(OC.Count);
            }
            //總計資料
            List<LawContent> SumAccpet = service.GetLawSumAccpet(year);
            SumTotal.Add(SumAccpet.Count);
            List<LawContent> SumOutsource = service.GetLawSumOutsource(year);
            SumTotal.Add(SumOutsource.Count);
            List<LawContent> SumDoself = service.GetLawSumDoself(year);
            SumTotal.Add(SumDoself.Count);
            List<LawContent> SumNotOC = service.GetLawSumNotOC(year);
            SumTotal.Add(SumNotOC.Count);
            List<LawContent> SumOC = service.GetLawSumOC(year);
            SumTotal.Add(SumOC.Count);
            return new Tuple<List<int>, List<int>, List<int>, List<int>, List<int>, List<int>>(Accpets, Outsources, Doselfs, NotOCs, OCs, SumTotal);
        }

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="selyear">年度</param>
        /// <param name="selyear">月份</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult NowYearReport()
        {
            var service = ServiceHelper.Create<ILAWService>();
            var ms = service.QueryLawNowYearReport();
            var filename = Url.Encode("案件統計表.xlsx");
            return File(ms, "application/octet-estream", filename);
        }
        #endregion
    }
}