using EP.H2OModels;
using EP.SD.SalesSupport.LAW.Service;
using EP.SD.SalesSupport.LAW.Web.Areas.LAW.Model;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Controllers
{
    [Program("LAWTX007")]
    public class LAWTX007Controller : BaseController
    {
        // GET: LAW/LAWTX007
        private ILAWService _Service;
        public LAWTX007Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }
        /// <summary>
        /// 報表排序設定
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX007")]
        public ActionResult Index()
        {
            return View();
        }

        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX007")]
        public ActionResult UpdateVM(string yy)
        {
            _Service.CheckLawReportSortBySortYear(yy);
            LawDetailModel model = new LawDetailModel();
            model.yy = yy;
            return View(model);
        }

        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX007")]
        public ActionResult UpdateSM(string yy)
        {
            _Service.CheckLawReportSortBySortYear(yy);
            LawDetailModel model = new LawDetailModel();
            model.yy = yy;
            return View(model);
        }

        /// <summary>
        /// 查詢列表
        /// </summary>
        /// <param name="SortYear">年度</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult VMData(string SortYear)
        {
            List<LawReportSort> QResultList = new List<LawReportSort>();
            var mService = new WebChannel<ILAWService>();

            mService.Use(service => service
            .GetLawReportSortVM(SortYear)
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new LawReportSort();
                    item.SortYear = d.SortYear;
                    item.SortVm = d.SortVm;
                    item.SortVmName = d.SortVmName;

                    QResultList.Add(item);
                }
            }));

            return Json(QResultList, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 查詢列表
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult SMData(string SortYear)
        {
            List<LawReportSort> QResultList = new List<LawReportSort>();
            var mService = new WebChannel<ILAWService>();

            mService.Use(service => service
            .GetLawReportSortSM(SortYear)
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new LawReportSort();
                    item.SortYear = d.SortYear;
                    item.SortVm = d.SortVm;
                    item.SortVmName = d.SortVmName;
                    item.SortSm = d.SortSm;
                    item.SortSmName = d.SortSmName;

                    QResultList.Add(item);
                }
            }));

            return Json(QResultList, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 更新區塊排序
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public void UpdateVM(string SortVmName, int SortVm, string yy)
        {
            bool result = _Service.UpdateLawReportSortVM(SortVmName, SortVm, yy);
        }

        /// <summary>
        /// 更新體系排序
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public void UpdateSM(string SortVmSmName, int SortSm, string yy)
        {
            string[] sArray = SortVmSmName.Split('|');
            
            bool result = _Service.UpdateLawReportSortSM(sArray[0], sArray[1], SortSm, yy);
        }

        /// <summary>
        /// 刪除排序年度
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX007")]
        public void Delete(string yy)
        {
            bool result = _Service.DeleteLawReportSortBySortYear(yy);

            if (result)
                AppendMessage(PlatformResources.刪除成功, false);
            else
                AppendMessage(PlatformResources.刪除失敗, false);
        }
    }
}