using EP.SD.SalesSupport.CUSCRM.Service;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using EP.Web;

namespace EP.SD.SalesSupport.CUSCRM.Web.Areas.CUSCRM.Controllers
{
    [Program("CUSCRMQU002")]
    public class CUSCRMQU002Controller : Controller
    {

        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMQU002")]
        public ActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// 查詢報表
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMQU002")]
        [PdLogFilter("EIP客服-報表查詢", PITraceType.Download)]

        public ActionResult QueryReport(QueryReportCondition condition)
        {
            var gridList = Enumerable.Empty<HistoryCSViewModel>().ToList();
            var mService = ServiceHelper.Create<IQueryService>();

            var ms =mService.GetReport(condition);
           
            return File(ms, "application/octet-estream", DataHelper.AddFileUniqueDownloadName("客服業務系統.xlsx"));
        }

    }
}