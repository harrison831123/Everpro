using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using EP.SD.SalesZone.AGUPG.Models;
using EP.SD.SalesZone.AGUPG.Service;
using EP.SD.SalesZone.AGUPG.Web.Areas.AGUPG.Utilities;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;

namespace EP.SD.SalesZone.AGUPG.Web.Areas.AGUPG.Controllers
{
    [Program("AGUPGQU001")]
    public class AGUPGQU001Controller : BaseController
    {
        private IAGUPGService _service;
        private static string _programID = "AGUPGQU001";
        public AGUPGQU001Controller()
        {
            _service = ServiceHelper.Create<IAGUPGService>();
        }

        // GET: AGUPG/AGUPGQU001
        [HasPermission("EP.SD.SalesZone.AGUPG.AGUPGQU001.*")]
        public ActionResult Index(string code, string YYYYSeason)
        {
            var list = _service.GetHrUpg25Rst();
            var vm = new HrUpg25RstViewModel
            {
                hrUpg25List = list,
                AgentCode = code ?? string.Empty,
                IsAdmin = AGUPGHelper.IsAdmin(User, _programID),
                YYYYSeason = YYYYSeason
            };

            return View(vm);
        }

        #region 查詢
        /// <summary>
        /// 取得查詢頁面資料
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesZone.AGUPG.AGUPGQU001.*")]
        public JsonResult Query(HrUpg25QueryCondition model)
        {
            Account info = ChkAccountInfo(model.AgentCode);

            var channel = new WebChannel<IAGUPGService>();
            //取資料
            HrUpg25Dto Result = new HrUpg25Dto();
            List<HrUpg25RstGrid1> QResultList1 = new List<HrUpg25RstGrid1>();
            List<HrUpg25RstGrid2> QResultList2 = new List<HrUpg25RstGrid2>();
            List<HrUpg25RstGrid3> QResultList3 = new List<HrUpg25RstGrid3>();

            try
            {
                model.AgentCode = info.MemberID;
                Result = _service.GetQueryHrUpg25(model);
                if (Result.HrUpg25RstTitle != null)
                {
                    Result.HrUpg25RstTitle.QueryDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                }
                if (Result.HrUpg25RstGrid1.Count > 0) { QResultList1 = Result.HrUpg25RstGrid1; }
                if (Result.HrUpg25RstGrid2.Count > 0) { QResultList2 = Result.HrUpg25RstGrid2; }
                if (Result.HrUpg25RstGrid3.Count > 0) { QResultList3 = Result.HrUpg25RstGrid3; }
            }
            catch (Exception ex)
            {
                AppendMessage("查詢資料時發生錯誤");
                Throw.LogError("查詢資料時發生錯誤: " + ex.Message);
            }

            var gridKey1 = channel.DataToCache(QResultList1.AsEnumerable());
            SetGridKey("HrUpg25RstGrid1", gridKey1);
            var gridKey2 = channel.DataToCache(QResultList2.AsEnumerable());
            SetGridKey("HrUpg25RstGrid2", gridKey2);
            var gridKey3 = channel.DataToCache(QResultList3.AsEnumerable());
            SetGridKey("HrUpg25RstGrid3", gridKey3);
            return Json(Result);
        }
        #endregion

        #region 明細
        /// <summary>
        /// 轄下第一代區級
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesZone.AGUPG.AGUPGQU001.*")]
        public ActionResult RightEmpOM01(HrUpg25QueryCondition model)
        {
            return View(model);
        }

        /// <summary>
        /// 轄下第一代區級Data
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesZone.AGUPG.AGUPGQU001.*")]
        public JsonResult RightEmpOM01Detail(HrUpg25QueryCondition model)
        {
            return GetDetail(model, r => r.HrUpgGet25Detail1, "HrUpgGet25Detail1");
        }

        /// <summary>
        /// 連續四季明細
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesZone.AGUPG.AGUPGQU001.*")]
        public ActionResult FourSeason(HrUpg25QueryCondition model)
        {
            return View(model);
        }

        /// <summary>
        /// 連續四季明細Data
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesZone.AGUPG.AGUPGQU001.*")]
        public JsonResult FourSeasonDetail(HrUpg25QueryCondition model)
        {
            return GetDetail(model, r => r.HrUpgGet25Detail2, "HrUpgGet25Detail2");
        }

        /// <summary>
        /// 加計被推介人業績
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesZone.AGUPG.AGUPGQU001.*")]
        public ActionResult IntroduceReturn(HrUpg25QueryCondition model)
        {
            return View(model);
        }

        /// <summary>
        /// 加計被推介人業績Data
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesZone.AGUPG.AGUPGQU001.*")]
        public JsonResult IntroduceReturnDetail(HrUpg25QueryCondition model)
        {
            return GetDetail(model, r => r.HrUpgGet25Detail3, "HrUpgGet25Detail3");
        }

        /// <summary>
        /// 遞延未核實
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesZone.AGUPG.AGUPGQU001.*")]
        public ActionResult VBPolicy(HrUpg25QueryCondition model)
        {
            return View(model);
        }

        /// <summary>
        /// 遞延未核實Data
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesZone.AGUPG.AGUPGQU001.*")]
        public JsonResult VBPolicyDetail(HrUpg25QueryCondition model)
        {
            return GetDetail(model, r => r.HrUpgGet25Detail4, "HrUpgGet25Detail4");
        }

        /// <summary>
        /// 共用明細綁定Jqgrid Key
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="model"></param>
        /// <param name="selector"></param>
        /// <param name="gridKey"></param>
        /// <returns></returns>
        private JsonResult GetDetail<T>(HrUpg25QueryCondition model, Func<HrUpgGet25Dto, List<T>> selector, string gridKey)
        {
            Account info = ChkAccountInfo(model.AgentCode);

            var channel = new WebChannel<IAGUPGService>();
            HrUpgGet25Dto result = new HrUpgGet25Dto();
            List<T> detailList = new List<T>();

            try
            {
                model.AgentCode = info.MemberID;
                result = _service.GetQueryHrUpgGet25WebShowDetail(model);
                // 使用 selector 找出對應的 Detail 清單
                detailList = selector(result);
                // List<T> 轉成 IEnumerable<T>並綁定
                var gridKey1 = channel.DataToCache(detailList.AsEnumerable());
                SetGridKey(gridKey, gridKey1);
            }
            catch (Exception ex)
            {
                AppendMessage("查詢資料時發生錯誤");
                Throw.LogError("查詢資料時發生錯誤: " + ex.Message);
            }

            return Json(result);
        }
        #endregion

        #region 私用函式
        private Account ChkAccountInfo(string AgentCode)
        {
            //當 model.AgentCode 不為空，表示是從 AGUPGQU002 或 明細 導頁過來的。
            var code10 = AgentCode?.Length >= 10 ? AgentCode.Substring(0, 10) : AgentCode;
            Account info = !String.IsNullOrEmpty(AgentCode) ? Account.Get(code10) : (Account)Session["AccountInfo"];

            if (string.IsNullOrEmpty(info?.ID))
            {
                Throw.LogError("請重新整理頁面，並重新輸入查詢條件");
            }

            return info;
        }
        #endregion
    }
}