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
    [Program("LAWTX005")]
    public class LAWTX005Controller : BaseController
    {
        // GET: LAW/LAWTX005
        private ILAWService _Service;
        public LAWTX005Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX005")]
        public ActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// 新增結案狀態
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX005")]
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
            List<LawCloseType> QResultList = new List<LawCloseType>();
            var mService = new WebChannel<ILAWService>();

            mService.Use(service => service
            .GetLawCloseType()
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new LawCloseType();
                    item.CloseTypeId = d.CloseTypeId;
                    item.CloseTypeName = d.CloseTypeName;
                    item.CountType = d.CountType;
                    item.StatusType = d.StatusType;
                    item.CreateName = d.CreateName;
                    item.CreateDate = d.CreateDate;

                    QResultList.Add(item);
                }
            }));

            return Json(QResultList, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 寫入結案狀態
        /// </summary>
        /// <param name="model">系統設定view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Create(LawDetailModel model)
        {
            LawCloseType LawCloseType = new LawCloseType();
            LawCloseType.CloseTypeName = model.CloseTypeName;
            LawCloseType.CountType = model.CountType;
            LawCloseType.StatusType = 1;
            LawCloseType.CreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
            LawCloseType.CreateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            LawCloseType.CloseTypeCreatorID = User.MemberInfo.ID;

            _Service.InsertLawCloseType(LawCloseType);

            return RedirectToAction("index", "LAWTX005");
        }

        /// <summary>
        /// 更新結案狀態
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX005")]
        public ActionResult Update(string id)
        {
            LawCloseType model = new LawCloseType();
            var mService = new WebChannel<ILAWService>();
            mService.Use(service => service
            .GetLawCloseTypeByID(id)
            .ForEach(d =>
            {
                if (d != null)
                {
                    model.CloseTypeId = d.CloseTypeId;
                    model.CloseTypeName = d.CloseTypeName;
                    model.CountType = d.CountType;
                    model.CreateName = d.CreateName;
                    model.CreateDate = d.CreateDate;
                }
            }));
            return View(model);
        }

        /// <summary>
        /// 更新結案狀態
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Update(LawCloseType model)
        {
            _Service.UpdateLawCloseType(model);

            return RedirectToAction("index", "LAWTX005");
        }

        /// <summary>
        /// 刪除一筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX005")]
        public void Delete(string ID)
        {
            bool result = _Service.DeleteLawCloseTypeByID(ID);

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
            _Service.ChangeLawCloseTypeStatusType(StatusType, ID);
        }
    }
}