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
    [Program("LAWTX006")]
    public class LAWTX006Controller : BaseController
    {
        // GET: LAW/LAWTX006
        private ILAWService _Service;
        public LAWTX006Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX006")]
        public ActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// 新增契撤變原因
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX006")]
        public ActionResult Create()
        {      
            return View();
        }

        /// <summary>
        /// 查詢列表
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult Query()
        {
            List<LawEvidType> QResultList = new List<LawEvidType>();
            var mService = new WebChannel<ILAWService>();

            mService.Use(service => service
            .GetLawEvidType()
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new LawEvidType();
                    item.EvidTypeId = d.EvidTypeId;
                    item.EvidTypeName = d.EvidTypeName;
                    item.EvidStatusType = d.EvidStatusType;
                    item.CreateName = d.CreateName;

                    QResultList.Add(item);
                }
            }));

            return Json(QResultList, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 寫入契撤變原因
        /// </summary>
        /// <param name="model">系統設定view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Create(LawDetailModel model)
        {
            LawEvidType LawEvidType = new LawEvidType();
            LawEvidType.EvidTypeName = model.EvidTypeName;
            LawEvidType.EvidStatusType = model.EvidStatusType;
            LawEvidType.CreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
            LawEvidType.CreateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            LawEvidType.EvidTypeCreatorID = User.MemberInfo.ID;

            _Service.InsertLawEvidType(LawEvidType);

            return RedirectToAction("index", "LAWTX006");
        }

        /// <summary>
        /// 更新契撤變原因
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX006")]
        public ActionResult Update(string id)
        {
            LawEvidType model = new LawEvidType();
            var mService = new WebChannel<ILAWService>();
            mService.Use(service => service
            .GetLawEvidTypeByID(id)
            .ForEach(d =>
            {
                if (d != null)
                {
                    model.EvidTypeId = d.EvidTypeId;
                    model.EvidTypeName = d.EvidTypeName;
                }
            }));
            return View(model);
        }

        /// <summary>
        /// 更新契撤變原因
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Update(LawEvidType model)
        {
            _Service.UpdateLawEvidType(model);

            return RedirectToAction("index", "LAWTX006");
        }

        /// <summary>
        /// 刪除一筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX006")]
        public void Delete(string ID)
        {
            bool result = _Service.DeleteLawEvidTypeByID(ID);

            if (result)
                AppendMessage(PlatformResources.刪除成功, false);
            else
                AppendMessage(PlatformResources.刪除失敗, false);
        }

        /// <summary>
        /// 更新契撤變原因狀態
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        //[HasPermission("EB.SL.PayRoll.PRTX002")]
        public void ChangeStatusType(string StatusType, string ID)
        {
            _Service.ChangeLawEvidTypeStatusType(StatusType, ID);
        }

        /// <summary>
        /// 確認契撤變原因有無重複
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        //[HasPermission("EB.SL.PayRoll.PRTX002")]
        public string chkEvidTypeNameRepeat(string EvidTypeName)
        {
            bool chk =  _Service.chkEvidTypeNameRepeat(EvidTypeName);
            string result = string.Empty;
            if (chk)
            {
                result = "OK";
            }
            else
            {
                result = "Repeat";
            }

            return result;
        }
        
    }
}