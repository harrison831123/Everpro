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
    [Program("MERSALTX002")]
    public class MerSalTX002Controller : BaseController
    {
        private IMerSalService _service;
        private static string _programID = "MerSalTX002";

        public MerSalTX002Controller()
        {
            _service = ServiceHelper.Create<IMerSalService>();
        }
        // GET: MerSal/MerSalTX002
        [HasPermission("EB.SL.MerSal.MerSalTX002")]
        public ActionResult Index()
        {
            var model = new MerSalCheckViewModel();
            return View(model);
        }

        /// <summary>
        /// 佣酬發放調整查詢
        /// </summary>
        /// <param name="MerSalViewModel"></param>
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalTX002")]
        public void Query(MerSalCheckViewModel model)
        {
            //取資料
            List<MerSalCheckViewModel> list = new List<MerSalCheckViewModel>();
            WebChannel<IMerSalService> _channelService = new WebChannel<IMerSalService>();
            model.ProductionYM = !String.IsNullOrEmpty(model.ProductionYM) ? StringExtension.WYearMonthToCYearMonth(model.ProductionYM) : model.ProductionYM;//轉民國年
            model.PayMonth = !String.IsNullOrEmpty(model.PayMonth) ? StringExtension.WYearMonthToCYearMonth(model.PayMonth) : model.PayMonth;//轉民國年
            model.NotPayYMSS = !String.IsNullOrEmpty(model.NotPayYMSS) ? StringExtension.WYearMonthToCYearMonth(model.NotPayYMSS.Substring(0, 7)) + model.NotPayYMSS.Substring(7, 2) : model.NotPayYMSS;
            model.NotPayYMSE = !String.IsNullOrEmpty(model.NotPayYMSE) ? StringExtension.WYearMonthToCYearMonth(model.NotPayYMSE.Substring(0, 7)) + model.NotPayYMSE.Substring(7, 2) : model.NotPayYMSE;
            _channelService.Use(service => list = service.GetMerSalCheck(model));
            var gridKey = _channelService.DataToCache(list.AsEnumerable());
            SetGridKey("QueryGrid", gridKey);
        }

        public JsonResult BindGrid(jqGridParam jqParams)
        {
            var cacheKey = GetGridKey("QueryGrid");
            return BaseGridBinding<MerSalCheckViewModel>(jqParams,
                () => new WebChannel<IMerSalService, MerSalCheckViewModel>().Get(cacheKey));
        }

        /// <summary>
        /// 更新佣酬發放調整
        /// </summary>
        /// <param name="MerSalViewModel"></param>
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalTX002")]
        public string UpdatePayTypes(string[] MainChkValueAll, string ProductionYMNow)
        {
            //取資料 業績年月==>發佣年月 & 發佣序號
            string[] sArray = ProductionYMNow.Split('-');
            string PayMonth = sArray[0];
            string PaySeq = sArray[1];
            bool result = false;

            string[] sArrayAll;
            string sAction1 = "";
            string sAction2 = "";
            string sAction3 = "";
            string sIden = "";
            string sMemo = "";

            List<string> MainChkValueAllList = MainChkValueAll == null ? new List<string>() : MainChkValueAll.ToList();

            foreach (var item in MainChkValueAllList)
            {

                sArrayAll = item.Split('|');
                sAction1 = sArrayAll[0];//pay_type有改否
                sAction2 = sArrayAll[1];//rpt_include_flag
                sAction3 = sArrayAll[2];//memo iden
                sIden = sArrayAll[3];   //Iden值
                sMemo = sArrayAll[4];//memo值


                MerSalCheckViewModel model = new MerSalCheckViewModel();
                //model.Iden = Convert.ToInt32(item);
                model.Iden = Convert.ToInt32(sIden);

                // 1-------------------------------------------------佣酬發放
                if (sAction1 == "V1")       //發
                {
                    model.CheckInd = "WarHuman02P";//(當工作月)人為註記發
                    model.PayType = "Pay";
                    model.PayMonth = PayMonth;
                    model.PaySeq = PaySeq;
                    model.NotPayYMS = "";
                    model.ChkCode = model.ChkCode + "|WarHuman02";
                }
                else if (sAction1 == "N1")  //不發
                {
                    model.CheckInd = "WarHuman01NP";//(當工作月)人為註記不發
                    model.PayType = "NotPay-02";
                    model.PayMonth = "";
                    model.PaySeq = "";
                    model.NotPayYMS = ProductionYMNow;
                    model.ChkCode = model.ChkCode + "|WarHuman01";

                    //if (sAction2 == "X2")//人為註記不發，一起改成結轉下期
                    //{
                    //    //model.RptIncludeFlag = "N";
                    //    sAction2 = "V2";
                    //}
                }
                // 2-------------------------------------------------結轉下期
                if (sAction2 == "V2")
                {
                    model.RptIncludeFlag = "Y";
                }
                else if (sAction2 == "N2")
                {
                    model.RptIncludeFlag = "N";

                    //不結轉下期，就一定不發20240422琍婷，所以自動將未註記不發的資料註記成不發(簽收回條未回又契撤的含在此規則中)
                    if (sAction1 == "X1")
                    {
                        sAction1 = "N1";
                        model.CheckInd = "WarHuman01NP";//(當工作月)人為註記不發
                        model.PayType = "NotPay-02";
                        model.PayMonth = "";
                        model.PaySeq = "";
                        model.NotPayYMS = ProductionYMNow;
                        model.ChkCode = model.ChkCode + "|WarHuman01";
                    }
                }
                // 3-------------------------------------------------人工備註
                if (sAction3 == "V3")
                {
                    model.Memo = sMemo;
                }

                model.ProcessUserCode = User.AccountInfo.ID;
                result = _service.UpdateMerSalCheck(model, sAction1, sAction2, sAction3);

            }

            if (result)
            {
                return "OK";
            }
            else
            {
                return "ERROR";
            }
        }

        /// <summary>
        /// 依照條件抓取報表
        /// </summary>
        /// <param name="MerSalViewModel"></param>
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalTX002")]
        public JsonResult GetMerSalCheckReport(MerSalCheckViewModel model,string ShowTitle)
        {
            string fileName = "佣酬發放調整報表" + "_" + model.CompanyCode + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            model.ProductionYM = !String.IsNullOrEmpty(model.ProductionYM) ? StringExtension.WYearMonthToCYearMonth(model.ProductionYM) : model.ProductionYM;
            model.PayMonth = !String.IsNullOrEmpty(model.PayMonth) ? StringExtension.WYearMonthToCYearMonth(model.PayMonth) : model.PayMonth;
            model.NotPayYMSS = !String.IsNullOrEmpty(model.NotPayYMSS) ? StringExtension.WYearMonthToCYearMonth(model.NotPayYMSS.Substring(0,7))+model.NotPayYMSS.Substring(7,2) : model.NotPayYMSS;
            model.NotPayYMSE = !String.IsNullOrEmpty(model.NotPayYMSE) ? StringExtension.WYearMonthToCYearMonth(model.NotPayYMSE.Substring(0,7))+model.NotPayYMSE.Substring(7,2) : model.NotPayYMSE;
            byte[] data = _service.GetMerSalCheckReportList(model, ShowTitle);
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
        [HasPermission("EB.SL.MerSal.MerSalTX002")]
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