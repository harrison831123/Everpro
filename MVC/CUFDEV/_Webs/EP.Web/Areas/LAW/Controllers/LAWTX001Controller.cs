using EP.H2OModels;
using EP.Platform.Service;
using EP.SD.SalesSupport.LAW.Service;
using EP.SD.SalesSupport.LAW.Web.Areas.LAW.Model;
using EP.SD.SalesSupport.LAW.Web.Areas.LAW.Utilities;
using EP.Web;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Controllers
{
    [Program("LAWTX001")]
    public class LAWTX001Controller : BaseController
    {
        // GET: LAW/LAWTX001
        private ILAWService _Service;
        public LAWTX001Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }

        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX001")]
        /// <summary>
        /// 管理人員設定
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 新增管理人員
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX001")]
        public ActionResult Create()
        {
            return View();
        }

        /// <summary>
        /// 寫入單筆調整
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult Query()
        {
            List<LawSys> QResultList = new List<LawSys>();
            var mService = new WebChannel<ILAWService>();

            mService.Use(service => service
            .GetLawSys()
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new LawSys();
                    item.ID = d.ID;
                    item.SysName = d.SysName;

                    QResultList.Add(item);
                }
            }));

            return Json(QResultList, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 寫入設定成員
        /// </summary>
        /// <param name="model">系統設定view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Create(LawDetailModel model)
        {
            //設定人員
            List<AccountGroupJsonString> JsonClassLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(model.LawimemberToJson);
            List<UserSimpleInfo> idnameLst = PlatformHelper.GetRecipientUsersList(JsonClassLst);
            List<LawSys> LawSyslist = GetIdsToList(idnameLst);
            List<LawSys> LawSysnewlist = new List<LawSys>();
            List<LawSys> QResultList = _Service.GetLawSys();
            List<string> vs = new List<string>();

            //轉為 Dictionary 物件
            var odDictionary = QResultList.ToDictionary(x => x.SysMemberID, x => x.SysMemberID);
            var memDictionary = LawSyslist.ToDictionary(x => x.SysMemberID, x => x.SysMemberID);

            //找出已是管理員的人
            var target = memDictionary.Where(x => !odDictionary.ContainsKey(x.Key)).ToList();
            vs = target.Select(x => x.Value).ToList();

            for (int i = 0; i < vs.Count; i++)
            {
                for (int j = 0; j < LawSyslist.Count; j++)
                {
                    if (vs[i].Equals(LawSyslist[j].SysMemberID))
                    {
                        var item = new LawSys();
                        item.SysName = LawSyslist[i].SysName;
                        item.SysMemberID = LawSyslist[i].SysMemberID;
                        item.SysCreatorID = LawSyslist[i].SysCreatorID;
                        item.Createname = LawSyslist[i].Createname;
                        item.CreateDate = LawSyslist[i].CreateDate;

                        LawSysnewlist.Add(item);
                    }
                }
            }

            _Service.InsertLawSys(LawSysnewlist);

            return RedirectToAction("index", "LAWTX001");
        }

        /// <summary>
        /// 刪除一筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX001")]
        public void Delete(string ID)
        {
            bool result = _Service.DeleteLawSysByID(ID);

            if (result)
                AppendMessage(PlatformResources.刪除成功, false);
            else
                AppendMessage(PlatformResources.刪除失敗, false);
        }

        /// <summary>
        /// 組成成員清單 
        /// </summary>
        /// <param name="idNamelist">成員資料</param>
        /// <returns></returns>
        private List<LawSys> GetIdsToList(List<UserSimpleInfo> idNamelist)
        {
            List<LawSys> list = new List<LawSys>();

            for (int i = 0; i < idNamelist.Count; i++)
            {
                LawSys Ls = new LawSys();

                Ls.Createname = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
                Ls.SysCreatorID = User.MemberInfo.ID;
                Ls.SysMemberID = idNamelist[i].MemberId;
                Ls.SysName = LAWHelper.ChangeUnitName(idNamelist[i].MemberId) + " " + idNamelist[i].Name;
                Ls.CreateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                list.Add(Ls);
            }
            return list;
        }
    }
}