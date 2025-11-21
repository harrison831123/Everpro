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
    [Program("LAWTX002")]
    public class LAWTX002Controller : BaseController
    {
        // GET: LAW/LAWTX002
        private ILAWService _Service;
        public LAWTX002Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX002")]
        /// <summary>
        /// 經辦人員設定
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 新增經辦人員
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX002")]
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
            List<LawDoUser> QResultList = new List<LawDoUser>();
            var mService = new WebChannel<ILAWService>();

            mService.Use(service => service
            .GetLawDouser()
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new LawDoUser();
                    item.LawDouserId = d.LawDouserId;
                    item.DouserName = d.DouserName;
                    item.DouserPhoneExt = d.DouserPhoneExt;
                    item.DouserSort = d.DouserSort;

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
            List<LawDoUser> LawDouserlist = GetIdsToList(idnameLst);
            List<LawDoUser> LawDousernewlist = new List<LawDoUser>();
            List<LawDoUser> QResultList = _Service.GetLawDouser();

            var QMemberList = QResultList.Select(x => x.DouserMemberID).ToList();
            var LMemberList = LawDouserlist.Select(x => x.DouserMemberID).ToList();
            var ExList = LMemberList.Except(QMemberList).ToList();
            if (ExList != null)
            {
                for (int i = 0; i < ExList.Count; i++)
                {
                    for (int j = 0; j < LawDouserlist.Count; j++)
                    {
                        if (ExList[i].Equals(LawDouserlist[j].DouserMemberID))
                        {
                            var item = new LawDoUser();
                            item.DouserName = LawDouserlist[i].DouserName;
                            item.DouserPhoneExt = model.DouserPhoneExt;
                            item.DouserSort = model.DouserSort;
                            item.DouserMemberID = LawDouserlist[i].DouserMemberID;
                            item.DouserCreatorID = LawDouserlist[i].DouserCreatorID;
                            item.CreateName = LawDouserlist[i].CreateName;
                            item.CreateDate = LawDouserlist[i].CreateDate;

                            LawDousernewlist.Add(item);
                        }
                    }
                }
            }


            _Service.InsertLawDouser(LawDousernewlist);

            return RedirectToAction("index", "LAWTX002");
        }

        /// <summary>
        /// 更新經辦人員
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX002")]
        public ActionResult Update(string id)
        {
            LawDoUser douser = new LawDoUser();
            var mService = new WebChannel<ILAWService>();
            mService.Use(service => service
            .GetLawDouserByID(id)
            .ForEach(d =>
            {
                if (d != null)
                {
                    douser.LawDouserId = d.LawDouserId;
                    douser.DouserName = d.DouserName;
                    douser.DouserPhoneExt = d.DouserPhoneExt;
                    douser.DouserSort = d.DouserSort;
                }
            }));
            return View(douser);
        }

        /// <summary>
        /// 更新經辦人員
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Update(LawDoUser model)
        {
            _Service.UpdateLawDouser(model);

            return RedirectToAction("index", "LAWTX002");
        }

        /// <summary>
        /// 刪除一筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX002")]
        public void Delete(string ID)
        {
            bool result = _Service.DeleteLawDouserByID(ID);

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
        private List<LawDoUser> GetIdsToList(List<UserSimpleInfo> idNamelist)
        {
            List<LawDoUser> list = new List<LawDoUser>();

            for (int i = 0; i < idNamelist.Count; i++)
            {
                LawDoUser Ls = new LawDoUser();

                Ls.DouserName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
                Ls.DouserCreatorID = User.MemberInfo.ID;
                Ls.DouserMemberID = idNamelist[i].MemberId;
                Ls.DouserName = LAWHelper.ChangeUnitName(idNamelist[i].MemberId) + " " + idNamelist[i].Name;
                Ls.CreateDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                Ls.CreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
                list.Add(Ls);
            }
            return list;
        }

        /// <summary>
        /// 確認經辦順位
        /// </summary>
        /// <param name="DouserSort">業績年月</param>
        /// <returns></returns>
        [HttpPost]
        public string CheckDouserSort(string DouserSort)
        {
            string result = _Service.CheckDouserSort(DouserSort);

            return result = result == "0" ? "OK" : result;
        }

        /// <summary>
        /// 確認經辦順位
        /// </summary>
        /// <param name="DouserSort">業績年月</param>
        /// <returns></returns>
        [HttpPost]
        public string CheckDouserSortByID(string DouserSort, string LawDouserId)
        {
            //先確認經辦順位資料是否和原本一樣
            string res = _Service.CheckDouserSortByID(DouserSort, LawDouserId);
            string result;

            if (res == "0")
            {
                //在確認經辦順位有無重覆
                result = _Service.CheckDouserSort(DouserSort);
            }
            else
            {
                return "OK";
            }

            return result = result == "0" ? "OK" : result;
        }
    }
}