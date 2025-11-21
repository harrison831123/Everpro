using EB.Common;
using EB.SL.MerSal.Models;
using EB.SL.MerSal.Service;
using EB.WebBrokerModels;
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
    [Program("MERSALTX003")]
    public class MerSalTX003Controller : BaseController
    {
        private IMerSalService _service;
        private static string _programID = "MerSalTX003";
        // GET: MerSal/MerSalTX003
        public MerSalTX003Controller()
        {
            _service = ServiceHelper.Create<IMerSalService>();
        }

        [HasPermission("EB.SL.MerSal.MerSalTX003")]
        public ActionResult Index(string msg = "")
        {
            MerSalCheckSPruleViewModel model = new MerSalCheckSPruleViewModel();
            model.YmClose = _service.GetYmClose();
            model.SeqClose = _service.GetSeqClose();
            if (msg != "")
            {
                AppendMessage(msg, false);
            };
            return View(model);
        }

        /// <summary>
        /// 查詢特殊資料設定
        /// </summary>
        /// <param name="model"></param>
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalTX003")]
        public JsonResult Query(MerSalCheckSPruleViewModel model)
        {
            //取資料
            List<MerSalCheckSPruleViewModel> dataList = new List<MerSalCheckSPruleViewModel>();
            WebChannel<IMerSalService> _channelService = new WebChannel<IMerSalService>();
            _channelService.Use(service => dataList = service.GetMerSalCheckSPrule(model,"Query"));

            //try
            //{
            //    for (int i = 0; i < dataList.Count; i++)
            //    {
            //        dataList[i].CreateUserCode = Member.Get(Account.Get(dataList[i].CreateUserCode).MemberID).Name;
            //        dataList[i].UpdateUserCode = !String.IsNullOrEmpty(dataList[i].UpdateUserCode) ? Member.Get(Account.Get(dataList[i].UpdateUserCode).MemberID).Name : "";
            //    }
            //}
            //catch (Exception ex)
            //{

            //}

            return Json(dataList);
        }

        /// <summary>
        /// 更新停用狀態
        /// </summary>
        /// <param name="model"></param>
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalTX003")]
        public JsonResult UpdateMerSalCheckSPruleIsDelete(MerSalCheckSPruleViewModel model)
        {
            model.UpdateUserCode = User.AccountInfo.ID;
            bool check = _service.UpdateMerSalCheckSPruleIsDelete(model);
            int checknumber;
            return Json(checknumber = check ? 1 : 0);
        }

        /// <summary>
        /// 新增特殊資料起始畫面
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [HasPermission("EB.SL.MerSal.MerSalTX003")]
        public ActionResult Create()
        {
            return View();
        }

        /// <summary>
        /// 寫入特殊資料設定
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalTX003")]
        public ActionResult Create(MerSalCheckSPruleViewModel model)
        {

            MerSalCheckSPrule merSalCheckSPrule = new MerSalCheckSPrule();
            if (!String.IsNullOrEmpty(model.ProductionYmS))
            {
                string[] sArray = model.ProductionYmS.Split('-');
                merSalCheckSPrule.ProductionYmS = sArray[0];
                merSalCheckSPrule.SequenceS = sArray[1];
            }
            else
            {
                merSalCheckSPrule.ProductionYmS = "";
                merSalCheckSPrule.SequenceS = "";
            }
            if (!String.IsNullOrEmpty(model.ProductionYmE))
            {
                string[] sArray = model.ProductionYmE.Split('-');
                merSalCheckSPrule.ProductionYmE = sArray[0];
                merSalCheckSPrule.SequenceE = sArray[1];
            }
            else
            {
                merSalCheckSPrule.ProductionYmE = "";
                merSalCheckSPrule.SequenceE = "";
            }
            merSalCheckSPrule.CompanyCode = model.CompanyCode;
            merSalCheckSPrule.FileSeq = model.FileSeq;
            merSalCheckSPrule.AmountType = model.AmountType;
            merSalCheckSPrule.ChkCodeActType = model.ChkCodeActType;

            string[] ChkCodeActTypeArray = model.ChkCodeActType.Split('|');
            merSalCheckSPrule.ChkCode = ChkCodeActTypeArray[0];
            merSalCheckSPrule.ActType = ChkCodeActTypeArray[1];
            merSalCheckSPrule.ActName = !String.IsNullOrEmpty(model.ActName) ? model.ActName : "";

            merSalCheckSPrule.Rule01 = model.Rule01;
            merSalCheckSPrule.Rule02 = !String.IsNullOrEmpty(model.Rule02) ? model.Rule02 : "";
            merSalCheckSPrule.Rule03 = !String.IsNullOrEmpty(model.Rule03) ? model.Rule03 : "";
            merSalCheckSPrule.Rule04 = !String.IsNullOrEmpty(model.Rule04) ? model.Rule04 : "";
            merSalCheckSPrule.Remark = !String.IsNullOrEmpty(model.Remark) ? model.Remark : "";
            merSalCheckSPrule.CreateUserCode = User.AccountInfo.ID;

            //_service.InsertMerSalCheckSPrule(merSalCheckSPrule);
            string result = _service.InsertMerSalCheckSPrule(merSalCheckSPrule);

            if (result == "OK")
            {
                return RedirectToAction("Index", new { msg = "新增成功" });

            }
            else if (result == "Duplicate")
            {
                AppendMessage("新增失敗，資料重複", false);
                return View(model);
            }
            else
            {
                return RedirectToAction("Index", new { msg = "新增失敗" });
            }

            //return RedirectToAction("Index");
        }

        /// <summary>
        /// 依照條件抓取報表
        /// </summary>
        /// <param name="MerSalCheckSPruleViewModel"></param>
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalTX003")]
        public JsonResult GetMerSalCheckSPruleReport(MerSalCheckSPruleViewModel model)
        {
            string fileName = "檢核特殊資料設定報表" + "_" + model.CompanyCode + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            byte[] data = _service.GetMerSalCheckSPruleReportList(model);
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
        /// 下載輸出
        /// </summary>
        /// <param name="fileGuid">guid</param>
        /// <param name="fileName">檔名</param>
        /// <returns></returns>
        [HttpGet]
        [HasPermission("EB.SL.MerSal.MerSalTX003")]
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