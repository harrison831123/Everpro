//需求單號：20241210003 現售商品佣獎查詢。 2024.12 BY VITA 
//需求單號：20250312002 增加險種中文、險種代碼查詢功能。 2025.03 BY VITA
//需求單號：20250312002 將下拉選單「請選擇」更改為「請下拉選單，選擇險種」。 2025.04 BY VITA 
using EP.Platform.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using EP.SD.Collections.PlanSet.Service;
using EP.SD.Collections.PlanSet.Models;

namespace EP.SD.Collections.PlanSet.Web.Controllers
{
    [Program("PLANSETQU001")]
    public class PlanSetQU001Controller : BaseController
    {
        private IPlanSetService _service;
        public PlanSetQU001Controller()
        {
            _service = ServiceHelper.Create<IPlanSetService>();
        }
        public ActionResult Index()
        {
            ViewBag.CompanyCodeList = GetCompanyCode();
            return View();
        }

        public List<SelectListItem> GetCompanyCode()
        {
            List<ValueText> list = _service.GetCompanyCode();

            List<SelectListItem> items = new List<SelectListItem>();
            items.Add(new SelectListItem() { Text = "請選擇", Value = "0" });

            foreach (var data in list)
            {
                items.Add(new SelectListItem() { Text = data.Text, Value = data.Value });
            }
            return items;
        }

        #region 查詢
        [HasPermission("EP.SD.Collections.PlanSet.PlanSetQU001")]
        public JsonResult Query(string CompanyCode, string plan_code, string collect)
        {
            Account info = (Account)Session["AccountInfo"];
            if (info.ID == "" || info.ID == null)
            {
                Throw.LogError("請重新整理頁面，並重新輸入查詢條件");
            }
            var channel = new WebChannel<IPlanSetService>();
            //取資料
            PlanSetAll Result = new PlanSetAll();
            List<PlanSetArea1> QResultList1 = new List<PlanSetArea1>();
            List<PlanSetArea2> QResultList2 = new List<PlanSetArea2>();
            List<PlanSetArea3> QResultList3 = new List<PlanSetArea3>();
            PlanTitle title = new PlanTitle();

            PlanSetMainInput model = new PlanSetMainInput();
            if (collect == "") { collect = null; }
            model.collect = collect;
            model.company_name = CompanyCode;
            model.plan_code = plan_code;

            try
            {
                Result = _service.GetQueryArea(model);//全部的area
                if (Result.PlanSetArea1.Count > 0) { QResultList1 = Result.PlanSetArea1; }
                if (Result.PlanSetArea2.Count > 0) { QResultList2 = Result.PlanSetArea2; }
                if (Result.PlanSetArea3.Count > 0) { QResultList3 = Result.PlanSetArea3; }
                if (Result.PlanSetArea1.Count > 0)
                {
                    title.company_name = QResultList1[0].company_name;
                    title.plan_title = QResultList1[0].plan_title;
                    title.plan_code = QResultList1[0].plan_code;
                    Result.PlanTitle = title;
                }
            }
            catch (Exception ex)
            {
                AppendMessage(ex.Message);
                Throw.LogError("查詢資料時發生錯誤: " + ex.Message);
            }

            var gridKey1 = channel.DataToCache(QResultList1.AsEnumerable());
            SetGridKey("PlanSetGrid1", gridKey1);
            var gridKey2 = channel.DataToCache(QResultList2.AsEnumerable());
            SetGridKey("PlanSetGrid2", gridKey2);
            var gridKey3 = channel.DataToCache(QResultList3.AsEnumerable());
            SetGridKey("PlanSetGrid3", gridKey3);
            var gridKey4 = channel.DataToCache(title);
            SetGridKey("PlanTitle", gridKey4);
            return Json(Result);
        }
        #endregion

        #region 查詢險種中文名稱
        /// <summary>
        /// 需求單號：20250312002 增加險種中文、險種代碼查詢功能。 2025.03 BY VITA
        ///                       修改「請選擇」呈現的文字。 2025.04 BY VITA
        /// 查詢險種中文名稱
        /// </summary>
        /// <param name="CompanyCode">保險公司代碼</param>
        /// <param name="plan_title">險種中文名稱 / 險種代碼</param>
        /// <param name="chktype">0:險種中文名稱、1:險種代碼</param>
        /// <returns></returns>
        public ActionResult GetPlanTitle(string CompanyCode, string plan_title, string chktype)
        {
            List<PlanTitle> result = _service.GetPlanTitle(CompanyCode, plan_title, chktype);
            List<string> PlanTitleList = new List<string>();

            if (result.Count > 0)
            {
                PlanTitleList.Add("- 請下拉選單，選擇險種 -");
                for (int i = 0; i < result.Count; i++)
                {
                    PlanTitleList.Add(result[i].plan_code + " " + result[i].plan_title);
                }
            }
            return Json(PlanTitleList, JsonRequestBehavior.AllowGet);
        }
        #endregion
    }
}