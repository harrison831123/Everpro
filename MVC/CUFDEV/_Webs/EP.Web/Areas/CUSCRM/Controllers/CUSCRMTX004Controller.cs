using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using EP.SD.SalesSupport.CUSCRM.Service;
using Microsoft.CUF.Framework;
using Microsoft.CUF.Framework.Service;

namespace EP.SD.SalesSupport.CUSCRM.Web
{
    /// <summary>
    /// 資料設定
    /// </summary>
    [Program("CUSCRMTX004")]
    public class CUSCRMTX004Controller : BaseController
    {
        /// <summary>
        /// 共用相關處理服務
        /// </summary>
        private static ICommonService _service;

        /// <summary>
        /// 建構子
        /// </summary>
        public CUSCRMTX004Controller()
        {
            _service = ServiceHelper.Create<ICommonService>();
        }

        /// <summary>
        /// 查詢頁面
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            MyPageStatus = null;
            return View();
        }

        /// <summary>
        /// 查詢的處理
        /// </summary>
        /// <param name="condition">查詢的條件</param>
        [HttpPost]
        public void QueryDiscipTypeGridDatas(QueryDiscipTypeCondition condition)
        {
            var channel = new WebChannel<ICommonService>();
            IEnumerable<DiscipTypeGridModel> gridList = null;

            channel.Use(proxy =>
            {
                var dataList = proxy.QueryCRMEDiscipTypeDatas(condition);
                gridList = dataList.Select(m =>
                {
                    return new DiscipTypeGridModel() {
                        ID = m.ID,
                        Code = m.Code.GetName(),
                        Kind = m.Kind.GetName(),
                        Status = m.Status.GetName(),
                        Name = m.Name
                    };
                });
            });

            var cacheKey = channel.DataToCache(gridList);
            SetGridKey("BindDiscipTypeDatas", cacheKey);
        }

        /// <summary>
        /// 回傳查詢結果
        /// </summary>
        /// <param name="jqParams"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult BindDiscipTypeDatas(jqGridParam jqParams)
        {
            var cacheKey = GetGridKey("BindDiscipTypeDatas");
            return BaseGridBinding<DiscipTypeGridModel>(jqParams,
                () => new WebChannel<ICommonService, DiscipTypeGridModel>().Get(cacheKey));
        }

        /// <summary>
        /// 明細資料取得與頁面呈現
        /// </summary>
        /// <param name="id">自動編號</param>
        /// <returns>明細頁面的呈現</returns>
        public ActionResult Detail(int id)
        {
            MyPageStatus = PageStatus.View;
            var model = _service.GetCRMEDiscipTypeByID(id);
            return View(model);
        }

        /// <summary>
        /// 新增的頁面呈現
        /// </summary>
        /// <returns>新增的頁面呈現</returns>
        public ActionResult Create()
        {
            MyPageStatus = PageStatus.Create;
            // 狀態預設為啟用, 代碼為SYS
            var model = new CRMEDiscipType() { Status = EnableStatus.Enabled, Code = DiscipTypeCode.SYS };
            return View("Detail", model);
        }

        /// <summary>
        /// 新增資料的處理
        /// </summary>
        /// <param name="model">要新增的資料</param>
        /// <returns>新增成功:顯示新增成功後的明細頁面 新增失敗:顯示新增前的資料頁面</returns>
        [HttpPost]
        public ActionResult Create(CRMEDiscipType model)
        {
            MyPageStatus = PageStatus.Create;
            TempData["model"] = model;

            if (!model.Kind.HasValue)
            {
                AppendMessage("請勾選類別");
                return View("Detail", model);
            }

            if (!model.Status.HasValue)
            {
                AppendMessage("請勾選狀態");
                return View("Detail", model);
            }

            model.Creator = User.MemberInfo.ID;
            model.CreateTime = DateTime.Now;

            var item = _service.CreateCRMEDiscipType(model);
            if (item.ID > 0)
            {
                AppendMessage("新增成功", true);
                return RedirectToAction("Detail", new { id = item.ID });
            }
            else
            {
                AppendMessage("新增失敗");
                return View("Detail", model);
            }
        }

        /// <summary>
        /// 編輯資料取得與頁面呈現
        /// </summary>
        /// <param name="id">自動編號</param>
        /// <returns>編輯頁面的呈現</returns>
        public ActionResult Edit(int id)
        {
            MyPageStatus = PageStatus.Edit;
            var model = _service.GetCRMEDiscipTypeByID(id);
            return View("Detail", model);
        }

        /// <summary>
        /// 修改資料的處理
        /// </summary>
        /// <param name="model">要修改的資料</param>
        /// <returns>修改成功:顯示修改成功後的明細頁面 修改失敗:顯示修改前的資料頁面</returns>
        [HttpPost]
        public ActionResult Edit(CRMEDiscipType model)
        {
            MyPageStatus = PageStatus.Edit;
            TempData["model"] = model;
            
            if (!model.Kind.HasValue)
            {
                AppendMessage("請勾選類別");
                return View("Detail", model);
            }

            if (!model.Status.HasValue)
            {
                AppendMessage("請勾選狀態");
                return View("Detail", model);
            }

            model.UpdateTime = DateTime.Now;
            model.Updator = User.MemberInfo.ID;


            model = _service.UpdateCRMEDiscipType(model);

            AppendMessage("更新成功", true);
            return RedirectToAction("Detail", new { id = model.ID });

        }

        /// <summary>
        /// 刪除的處理
        /// </summary>
        /// <param name="model">要刪除的資料</param>
        /// <returns>刪除成功:關閉頁面 刪除失敗:顯示明細頁面，並且提示訊息</returns>
        [HttpPost]
        public ActionResult Delete(CRMEDiscipType model)
        {
            _service.DeleteCRMEDiscipType(model.ID);
            AppendMessage("刪除成功");
            return View("Close");
        }

        /// <summary>
        /// 取消動作的處理
        /// </summary>
        /// <param name="id">自動編號</param>
        /// <returns>取消目前的動作，回到明細頁面的呈現</returns>
        public ActionResult Cancel(int id)
        {
            return RedirectToAction("Detail", new { id});
        }
    }
}