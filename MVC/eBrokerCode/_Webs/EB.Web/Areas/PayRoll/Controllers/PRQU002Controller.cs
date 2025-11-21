using EB.SL.PayRoll.Models;
using EB.SL.PayRoll.Service;
using EB.SL.PayRoll.Service.Contracts;
using EB.SL.PayRoll.Web.Areas.PayRoll.Utilities;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EB.SL.PayRoll.Web.Areas.PayRoll.Controllers
{
    [Program("PRQU002")]
    public class PRQU002Controller : BaseController
    {
        private IPayRollService _service;

        public PRQU002Controller()
        {
            _service = ServiceHelper.Create<IPayRollService>();
        }
        // GET: PayRoll/PRQU002
        [HasPermission("EB.SL.PayRoll.PRQU002")]
        public ActionResult Index()
        {
			AgentBonusAdjustViewModel model = new AgentBonusAdjustViewModel();
			model.nmember = User.MemberInfo.Name;
			return View(model);
        }

        /// <summary>
        /// 依照年分撈取業績年月
        /// </summary>
        /// <param name="productionY">業績年</param>
        [HttpPost]
        [HasPermission("EB.SL.PayRoll.PRQU002")]
        public JsonResult GetproductionYM(string productionY)
        {
            List<string> result = new List<string>();
            string unid = PayRollHelper.ChangeUnitID(User.MemberInfo.ID);
            if (User.HasPermission("EB.SL.PayRoll.PRTX002.Admin") || User.HasPermission("EB.SL.PayRoll.PRTX002.FIN"))
			{
                result = _service.GetAgentBonusAdjustMainProductionYm(productionY, unid);
            }
			else
			{
                result = _service.GetAgentBonusAdjustProductionYm(productionY, unid);
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
        [HasPermission("EB.SL.PayRoll.PRQU002")]
        public ActionResult GetTotalReasonCodeReport(AgentBonusAdjustCondition Condition)
        {
            string fileName = "調帳原因碼總表" + ".xlsx";
            Condition.CreateReportUserName = User.MemberInfo.Name;
            Condition.CreateReportUnitName = PayRollHelper.ChangeUnitName(User.MemberInfo.ID);
            var List = _service.CheckTotalReasonCodeReport(Condition);

            if (List.Count != 0)
            {
                var ms = _service.GetTotalReasonCodeReport(Condition);
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