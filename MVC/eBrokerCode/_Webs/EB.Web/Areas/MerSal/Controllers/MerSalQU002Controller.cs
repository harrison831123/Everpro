using EB.Common;
using EB.SL.MerSal.Models;
using EB.SL.MerSal.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EB.SL.MerSal.Web.Areas.MerSal.Controllers
{
    [Program("MERSALQU002")]
    public class MerSalQU002Controller : BaseController
    {
        private IMerSalService _service;
        private static string _programID = "MerSalQU002";
        // GET: MerSal/MerSalQU002
        public MerSalQU002Controller()
        {
            _service = ServiceHelper.Create<IMerSalService>();
        }
        [HasPermission("EB.SL.MerSal.MerSalQU002")]
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 原始檔系統保留暨人工調帳查詢
        /// </summary>
        /// <param name="MerSalViewModel"></param>
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalQU002")]
        public void Query(MerSalCutViewModel model)
        {
            //取資料
            List<MerSalCutViewModel> list = new List<MerSalCutViewModel>();
            WebChannel<IMerSalService> _channelService = new WebChannel<IMerSalService>();
            model.ProductionYM = StringExtension.WYearMonthToCYearMonth(model.ProductionYM);
            _channelService.Use(service => list = service.GetMerSalCut(model));

            var gridKey = _channelService.DataToCache(list.AsEnumerable());
            SetGridKey("QueryGrid", gridKey);
        }

        public JsonResult BindGrid(jqGridParam jqParams)
        {
            var cacheKey = GetGridKey("QueryGrid");
            return BaseGridBinding<MerSalCutViewModel>(jqParams,
                () => new WebChannel<IMerSalService, MerSalCutViewModel>().Get(cacheKey));
        }

        /// <summary>
        /// 依照條件抓取報表
        /// </summary>
        /// <param name="productionY">業績年</param>
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalQU002")]
        public JsonResult GetMerSalCutReport(MerSalCutViewModel model)
        {
            var service = ServiceHelper.Create<IMerSalService>();
            string fileName = "";
            model.ProductionYM = StringExtension.WYearMonthToCYearMonth(model.ProductionYM);

            //檔名
            if (model.BtnType == "btnReport")
            {
                fileName = "原始檔系統保留暨人工調帳報表_" +model.ProductionYM.Replace("/","")+"_"+model.CompanyCode+ ".xlsx";
            }
            else
            {
                fileName = "人工調帳轉大批格式報表_" + model.ProductionYM.Replace("/", "") + "_" + model.CompanyCode + ".xlsx";
            }

            //to Service
            byte[] data = service.GetMerSalCutReportList(model,fileName);

            string handle = Guid.NewGuid().ToString();
            TempData[handle] = data;
            if (data != null)
            {
                TempData[handle] = data;
                return new JsonResult()
                {
                    Data = new
                    {
                        FileGuid = handle
                        ,
                        FileName = fileName
                    }
                };
            }
            else
            {
                AppendMessage("查無資料");
                return Json("Error", JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// 下載輸出
        /// </summary>
        /// <param name="fileGuid">guid</param>
        /// <param name="fileName">檔名</param>
        /// <returns></returns>
        [HttpGet]
        [HasPermission("EB.SL.MerSal.MerSalQU002")]
        public virtual ActionResult Download(string fileGuid, string fileName)
        {
            if (TempData[fileGuid] != null)
            {
                byte[] data = TempData[fileGuid] as byte[];
                //return File(data, "application/vnd.ms-excel", fileName);
                return File(data, MimeMapping.GetMimeMapping(fileName), fileName);
            }
            else
            {
                // Problem - Log the error, generate a blank file,
                return new EmptyResult();
            }
        }
    }
}