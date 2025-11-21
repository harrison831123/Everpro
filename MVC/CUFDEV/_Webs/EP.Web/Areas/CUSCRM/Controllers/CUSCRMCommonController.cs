using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using EP.SD.SalesSupport.CUSCRM.Service;
using Microsoft.CUF.Framework.Service;

namespace EP.SD.SalesSupport.CUSCRM.Web
{
    public class CUSCRMCommonController : BaseController
    {

        public ActionResult GetCaseType(Category? category)
        {
            var result = CUSCRMHelper.GetCaseTypeList(category);
            return Json(result);
        }
    }
}