using EB.Common;
using EB.Platform.Service;
using EB.SL.MerSal.Models;
using EB.SL.MerSal.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EB.SL.MerSal.Web.Areas.MerSal.Controllers
{
    [Program("MERSALQU001")]
    public class MerSalQU001Controller : BaseController
    {
        private IMerSalService _service;
        private IVlifeService _vlifeService;
        private static string _programID = "MerSalQU001";

        public MerSalQU001Controller()
        {
            _service = ServiceHelper.Create<IMerSalService>();
            _vlifeService = ServiceHelper.Create<IVlifeService>();
        }
        // GET: MerSal/MerSalQU001
        [HasPermission("EB.SL.MerSal.MerSalQU001")]
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 原始檔報表-入佣資料報表
        /// </summary>
        /// <param name="MerSalViewModel"></param>
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalQU001")]
        public JsonResult GetMerSalDReport(MerSalViewModel model)
        {
            //Excel檔名
            string fileName = "入佣資料報表" + "_" + model.ProductionYM.Replace("/","") + model.Sequence.ToString() + "_" + model.CompanyCode + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            model.ProductionYM = StringExtension.WYearMonthToCYearMonth(model.ProductionYM); //西元年轉民國年
            model.QueryUser = User.MemberInfo.Name;
            model.QueryDate = DateTime.Now;
            byte[] data = _service.GetMerSalDReportList(model);
            string handle = Guid.NewGuid().ToString();

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
        /// 原始檔報表-保險公司佣酬檢核表
        /// </summary>
        /// <param name="selyear">年度</param>
        /// <param name="selyear">月份</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult GetCompaneyMerSalDReport(MerSalViewModel model)
        {
            //Excel檔名
            string fileName = "保險公司佣酬檢核表" + "_" + model.ProductionYM.Replace("/", "") + model.Sequence.ToString() + "_" + model.CompanyCode + "_" +DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            model.ProductionYM = StringExtension.WYearMonthToCYearMonth(model.ProductionYM); //西元年轉民國年

            var ms = _service.GetCompanyMerSalDReportList(model);
            var filename = Url.Encode(fileName);

            if (ms != null)
            {
                return File(ms, "application/octet-estream", filename);
            }
            else
            {
                Throw.BusinessError("查無資料");
                return RedirectToAction("Index");
            }

        }

        /// <summary>
        /// 下載輸出
        /// </summary>
        /// <param name="fileGuid">guid</param>
        /// <param name="fileName">檔名</param>
        /// <returns></returns>
        [HttpGet]
        [HasPermission("EB.SL.MerSal.MerSalQU001")]
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