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
    [Program("LAWTX003")]
    public class LAWTX003Controller : BaseController
    {
        // GET: LAW/LAWTX003
        private ILAWService _Service;
        public LAWTX003Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX003")]
        public ActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// 新增承辦單位
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX003")]
        public ActionResult Create()
        {
            LawDetailModel model = new LawDetailModel();
            model.CreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
            model.CreateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            return View(model);
        }

        /// <summary>
        /// 查詢列表
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult Query()
        {
            List<LawDoUnit> QResultList = new List<LawDoUnit>();
            var mService = new WebChannel<ILAWService>();

            mService.Use(service => service
            .GetLawDoUnit()
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new LawDoUnit();
                    item.LawDoUnitId = d.LawDoUnitId;
                    item.UnitName = d.UnitName;
                    item.StatusType = d.StatusType;
                    item.CreateName = d.CreateName;
                    item.CreateDate = d.CreateDate;

                    QResultList.Add(item);
                }
            }));

            return Json(QResultList, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 寫入承辦單位
        /// </summary>
        /// <param name="model">系統設定view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Create(LawDetailModel model)
        {
            LawDoUnit lawDoUnit = new LawDoUnit();
            lawDoUnit.UnitName = model.UnitName;
            lawDoUnit.StatusType = 1;
            lawDoUnit.CreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
            lawDoUnit.CreateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            lawDoUnit.DounitCreatorID = User.MemberInfo.ID;

            _Service.InsertLawDoUnit(lawDoUnit);

            return RedirectToAction("index", "LAWTX003");
        }

        /// <summary>
        /// 更新承辦單位
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX003")]
        public ActionResult Update(string id)
        {
            LawDoUnit model = new LawDoUnit();
            var mService = new WebChannel<ILAWService>();
            mService.Use(service => service
            .GetLawDoUnitByID(id)
            .ForEach(d =>
            {
                if (d != null)
                {
                    model.LawDoUnitId = d.LawDoUnitId;
                    model.UnitName = d.UnitName;
                    model.CreateName = d.CreateName;
                    model.CreateDate = d.CreateDate;
                }
            }));
            return View(model);
        }

        /// <summary>
        /// 更新承辦單位
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Update(LawDoUnit model)
        {
            _Service.UpdateLawDoUnit(model);

            return RedirectToAction("index", "LAWTX003");
        }

        /// <summary>
        /// 刪除一筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX003")]
        public void Delete(string ID)
        {
            bool result = _Service.DeleteLawDoUnitByID(ID);

            if (result)
                AppendMessage(PlatformResources.刪除成功, false);
            else
                AppendMessage(PlatformResources.刪除失敗, false);
        }

        /// <summary>
        /// 更新承辦狀態
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        //[HasPermission("EB.SL.PayRoll.PRTX002")]
        public void ChangeStatusType(string StatusType, string ID)
        {
            _Service.ChangeStatusType(StatusType, ID);
        }       
    }
}