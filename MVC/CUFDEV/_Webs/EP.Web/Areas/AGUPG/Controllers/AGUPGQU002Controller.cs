using EP.SD.SalesZone.AGUPG.Service;
using EP.SD.SalesZone.AGUPG.Web.Areas.AGUPG.Utilities;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static EP.SD.SalesZone.AGUPG.Models.Enumerations;

namespace EP.SD.SalesZone.AGUPG.Web.Areas.AGUPG.Controllers
{
    [Program("AGUPGQU002")]
    public class AGUPGQU002Controller : BaseController
    {
        private IAGUPGService _service;
        private static string _programID = "AGUPGQU002";
        public AGUPGQU002Controller()
        {
            _service = ServiceHelper.Create<IAGUPGService>();
        }

        // GET: AGUPG/AGUPGQU002
        [HasPermission("EP.SD.SalesZone.AGUPG.AGUPGQU002.*")]
        public ActionResult Index()
        {
            HrUpg25QueryCondition model = new HrUpg25QueryCondition();
            model.AgentCode = User.AccountInfo.MemberID;
            model.UserType = AGUPGHelper.GetUserType(User, _programID);

            var data = _service.GetHrUpg25RstFamilyTree(model);

            if (data == null)
                return RedirectToAction("Index", "AGUPGQU001");

            return View(data);
        }
    }
}