//需求單號：20250707001 競賽計C商品清單。 2025/07/03 BY Harrison 
using EP.Common;
using EP.SD.Collections.PlanSet.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.Collections.PlanSet.Web.Controllers
{
    [Program("PLANSETQU003")]
    public class PlanSetQU003Controller : BaseController
    {
        private IPlanSetService _service;
        public PlanSetQU003Controller()
        {
            _service = ServiceHelper.Create<IPlanSetService>();
        }

		// GET: PlanSet/PlanSetQU003
		[HasPermission("EP.SD.Collections.PlanSet.PlanSetQU003")]
		public ActionResult Index()
        {
			ViewBag.WorkDate = _service.GetWorkDate();
			var model = new PlanSetWarptSetCondition
            {
                SelectedProduct = ProductList.All // 預設選擇
            };
            return View(model);
        }

		/// <summary>
		/// 競賽計C商品清單報表
		/// </summary>
		/// <param name="condition"></param>
		/// <returns></returns>
		[HttpPost]
		[HasPermission("EP.SD.Collections.PlanSet.PlanSetQU003")]
		public JsonResult GetPlanSetWarptSetReport(PlanSetWarptSetCondition condition)
		{
			string dateTime = DateTime.Now.ToString("yyyyMMdd");
			//Excel檔名
			string fileName = "競賽計C商品清單_" + dateTime + ".xlsx";
			var Check = _service.GetPlanSetWarptSet(condition);
			//var filename = Url.Encode(fileName);

			if (Check != null || Check.Count != 0)
			{
				var ms = _service.GetPlanSetWarptSetReport(condition);
				string handle = Guid.NewGuid().ToString();
				TempData[handle] = ms;
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
				Throw.BusinessError("無該資料，無法產出報表");
				return Json("ERROR", JsonRequestBehavior.AllowGet);
			}
		}

		/// <summary>
		/// 下載輸出
		/// </summary>
		/// <param name="fileGuid">guid</param>
		/// <param name="fileName">檔名</param>
		/// <returns></returns>
		[HttpGet]
		public virtual ActionResult Download(string fileGuid, string fileName)
		{
			if (TempData[fileGuid] != null)
			{
				Stream data = (Stream)TempData[fileGuid];
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