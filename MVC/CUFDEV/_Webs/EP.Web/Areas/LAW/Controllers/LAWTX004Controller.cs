using EP.H2OModels;
using EP.SD.SalesSupport.LAW.Service;
using EP.SD.SalesSupport.LAW.Web.Areas.LAW.Model;
using EP.SD.SalesSupport.LAW.Web.Areas.LAW.Utilities;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Controllers
{
    [Program("LAWTX004")]
    public class LAWTX004Controller : BaseController
    {
        // GET: LAW/LAWTX004
        private ILAWService _Service;
        public LAWTX004Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX004")]
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 新增利率或律師服務費率
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX004")]
        public ActionResult Create(string LirType)
        {
            LawDetailModel model = new LawDetailModel();
            model.LirType = LirType;
            model.CreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
            model.CreateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            return View(model);
        }

        /// <summary>
        /// 列表
        /// </summary>
        /// <param name="cond"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult Query(string LirType)
        {
            if (LirType == "1") //利率
            {
                List<LawInterestRates> QResultList = new List<LawInterestRates>();
                var mService = new WebChannel<ILAWService>();
                mService.Use(service => service
                .GetLawInterestRates()
                .ForEach(d =>
                {
                    if (d != null)
                    {
                        var item = new LawInterestRates();
                        item.LawInterestRatesId = d.LawInterestRatesId;
                        item.InterestRates = d.InterestRates;
                        item.InterestRatesType = d.InterestRatesType;
                        item.CreateName = d.CreateName;
                        item.CreateDate = d.CreateDate;
                        QResultList.Add(item);
                    }
                }));

                return Json(QResultList, JsonRequestBehavior.AllowGet);
            }
            else //律師服務費率
            {
                List<LawLawyerServiceRates> QResultList = new List<LawLawyerServiceRates>();
                var mService = new WebChannel<ILAWService>();
                mService.Use(service => service
                .GetLawLawyerServiceRates()
                .ForEach(d =>
                {
                    if (d != null)
                    {
                        var item = new LawLawyerServiceRates();
                        item.LawLawyerServiceRatesId = d.LawLawyerServiceRatesId;
                        item.LawyerServiceRates = d.LawyerServiceRates;
                        item.LawyerServiceRatesType = d.LawyerServiceRatesType;
                        item.CreateName = d.CreateName;
                        item.CreateDate = d.CreateDate;
                        QResultList.Add(item);
                    }
                }));

                return Json(QResultList, JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// 寫入利率或律師服務費率
        /// </summary>
        /// <param name="model">系統設定view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Create(LawDetailModel model)
        {
            if(model.LirType == "1")
            {
                LawInterestRates lawInterest = new LawInterestRates();
                lawInterest.InterestRates = model.InterestRates;
                lawInterest.CreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
                lawInterest.CreateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                lawInterest.InterestRatesCreateID = User.MemberInfo.ID;
                lawInterest.InterestRatesType = 1;
                _Service.InsertLawInterestRates(lawInterest);
            }
            else
            {
                LawLawyerServiceRates lawLawyerService = new LawLawyerServiceRates();
                lawLawyerService.LawyerServiceRates = model.LawyerServiceRates;
                lawLawyerService.CreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
                lawLawyerService.CreateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                lawLawyerService.LawyerServiceRatesCreatorID = User.MemberInfo.ID;
                lawLawyerService.LawyerServiceRatesType = 1;
                _Service.InsertLawLawyerServiceRates(lawLawyerService);
            }

            return RedirectToAction("index", "LAWTX004");
        }

        /// <summary>
        /// 刪除一筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX004")]
        public void Delete(string ID, string LirType)
        {
            bool result = false;
            if (LirType == "1")
            {
                result = _Service.DeleteLawInterestRatesByID(ID);
            }
            else
            {
                result = _Service.DeleteLawLawyerServiceRatesByID(ID);
            }
            if (result)
                AppendMessage(PlatformResources.刪除成功, false);
            else
                AppendMessage(PlatformResources.刪除失敗, false);
        }

        /// <summary>
        /// 更新利率或律師服務費率狀態
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        public void ChangeStatusType(string StatusType, string ID,string LirType)
        {
            if(LirType == "1")
            {
                _Service.ChangeInterestRatesType(StatusType, ID);
            }
            else
            {
                _Service.ChangeLawyerServiceRatesType(StatusType, ID);
            }
        }

    }
}