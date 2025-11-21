using EP.H2OModels;
using EP.SD.SalesSupport.LAW.Models;
using EP.SD.SalesSupport.LAW.Service;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Controllers
{
    [Program("LAWTX010")]
    public class LAWTX010Controller : BaseController
    {
        // GET: LAW/LAWTX010
        private ILAWService _Service;
        public LAWTX010Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX010")]
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 電話催告通知
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public void Query(string LawSearchType)
        {
            //取資料
            var mService = new WebChannel<ILAWService>();
            List<LawPhoneCallLogDetail> Viewmodel = new List<LawPhoneCallLogDetail>();
            WebChannel<ILAWService> _channelService = new WebChannel<ILAWService>();

            if (LawSearchType == "1")
            {
                _channelService.Use(service => Viewmodel = service.GetLawPhoneCallLog());
            }
            else
            {
                _channelService.Use(service => Viewmodel = service.GetThirtyDayLawContent());
            }

            var gridKey = _channelService.DataToCache(Viewmodel.AsEnumerable());
            SetGridKey("QueryGrid", gridKey);
        }

        public JsonResult BindGrid(jqGridParam jqParams)
        {
            var cacheKey = GetGridKey("QueryGrid");
            return BaseGridBinding<LawPhoneCallLogDetail>(jqParams,
                () => new WebChannel<ILAWService, LawPhoneCallLogDetail>().Get(cacheKey));
        }

        /// <summary>
        /// 更新電話催告通知
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX010")]
        public void Update()
        {
            _Service.UpdateLawPhoneCallLog(User.MemberInfo.ID);
        }
    }
}