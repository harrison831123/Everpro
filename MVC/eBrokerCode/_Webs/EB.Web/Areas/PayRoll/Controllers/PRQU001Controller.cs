///========================================================================================== 
/// 程式名稱：人工調帳
/// 建立人員：Harrison
/// 建立日期：2022/07
/// 修改記錄：（[需求單號]、[修改內容]、日期、人員）
/// 需求單號:20240122004-因現有VLIFE系統(核心系統)使用已長達20多年，架構老舊，已不敷使用，且為提升資訊安全等級，故計劃執行VLIFE系統改版(新核心系統:eBroker系統)。; 修改內容:上線; 修改日期:20240613; 修改人員:Harrison;
/// 需求單號:20240807001-調整人工調帳系統產出之相關畫面及報表修改等功能。; 修改日期:20240807; 修改人員:Harrison;
///==========================================================================================
using EB.SL.PayRoll.Models;
using EB.SL.PayRoll.Service;
using EB.SL.PayRoll.Web.Areas.PayRoll.Utilities;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EB.SL.PayRoll.Web.Areas.PayRoll.Controllers
{
    [Program("PRQU001")]
    public class PRQU001Controller : BaseController
    {
        private IPayRollService _service;

        public PRQU001Controller()
        {
            _service = ServiceHelper.Create<IPayRollService>();
        }
        // GET: PayRoll/PRQU001
        [HasPermission("EB.SL.PayRoll.PRQU001")]
        public ActionResult Index()
        {
            AgentBonusAdjustViewModel model = new AgentBonusAdjustViewModel();
            model.nmember = User.MemberInfo.Name;
            return View(model);
        }

        /// <summary>
        /// AG078報表明細查詢
        /// </summary>
        /// <param name="AgentBonusAdjustViewModel"></param>
        [HttpPost]
        [HasPermission("EB.SL.PayRoll.PRQU001")]
        public void Query(AgentBonusAdjustViewModel model)
        {
            string MemberID = User.MemberInfo.ID;

            //取資料
            List<AgentBonusAdjustViewModel> list = new List<AgentBonusAdjustViewModel>();
            WebChannel<IPayRollService> _channelService = new WebChannel<IPayRollService>();

            _channelService.Use(service => list = service.GetAgentBonusAdjust(model));

            for (int i = 0; i < list.Count; i++)
            {
                list[i].CreateDatetimeString = list[i].CreateDatetime.ToString("yyyy/MM/dd");
                list[i].ProductionYMSequence = list[i].ProductionYM + "-" + list[i].Sequence;
                list[i].AmountTS = list[i].Amount != 0 ? list[i].Amount.ToString("###,###") : "0";
                list[i].FYCTS = list[i].FYC != 0 ? list[i].FYC.ToString("###,###") : "0";
                list[i].FYPTS = list[i].FYP != 0 ? list[i].FYP.ToString("###,###") : "0";
            }

            var gridKey = _channelService.DataToCache(list.AsEnumerable());
            SetGridKey("QueryGrid", gridKey);
        }

        [HasPermission("EB.SL.PayRoll.PRQU001")]
        public JsonResult BindGrid(jqGridParam jqParams)
        {
            var cacheKey = GetGridKey("QueryGrid");
            return BaseGridBinding<AgentBonusAdjustViewModel>(jqParams,
                () => new WebChannel<IPayRollService, AgentBonusAdjustViewModel>().Get(cacheKey));
        }

        /// <summary>
        /// 依照年分撈取業績年月
        /// </summary>
        /// <param name="productionY">業績年</param>
        [HttpPost]
        public JsonResult GetproductionYM(string productionY)
        {
            List<string> result = new List<string>();
            productionY = (Convert.ToInt32(productionY) - 1911).ToString();
            result = _service.GetAllYMData(productionY);
            if (result != null)
            {
                for (int i = 0; i < result.Count; i++)
                {
                    result[i] = PayRollHelper.RocToWY(result[i]);
                }
            }

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="selyear">年度</param>
        /// <param name="selyear">月份</param>
        /// <returns></returns>
        [HttpPost]
        [HasPermission("EB.SL.PayRoll.PRQU001")]
        public ActionResult GetAg078Report(AgentBonusAdjustViewModel model)
        {
            string fileName = "人工調整明細報表" + ".xlsx";
            var List = _service.GetAgentBonusAdjust(model);

            if (List.Count != 0)
            {
                var ms = _service.GetAG078ReportList(model);
                var filename = Url.Encode(fileName);
                return File(ms, "application/octet-estream", filename);
            }
            else
            {
                TempData["message"] = "查無資料!";
                return RedirectToAction("Index");
            }
        }

	}
}