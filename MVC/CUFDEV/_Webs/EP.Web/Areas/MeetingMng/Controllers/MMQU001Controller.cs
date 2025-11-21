using EP.H2OModels;
using EP.PSL.WorkResources.MeetingMng.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.PSL.WorkResources.MeetingMng.Web.Areas.MeetingMng.Controllers
{
    
    public class MMQU001Controller : BaseController
    {
        // GET: MeetingMng/MMQU001
        private IMeetingMngService _Service;
        public MMQU001Controller()
        {
            _Service = ServiceHelper.Create<IMeetingMngService>();
        }

        /// <summary>
        /// 會議列表
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.PSL.WorkResources.MeetingMng.MMQU001")]
        public ActionResult Index()
        {
            //更新會議狀態
            _Service.UpdateMeetingTime();
            //取資料
            var condition = new QueryMeetingCondition();

            //會議類別的下拉選單
            condition.MeetingReadType = "1";

            return View(condition);
        }

        /// <summary>
        /// 會議列表JQGrid
        /// </summary>
        /// <param name="cond"></param>
        /// <returns></returns>
        [HasPermission("EP.PSL.WorkResources.MeetingMng.MMQU001")]
        public void Query(QueryMeetingCondition cond)
        {
            cond.imember = User.MemberInfo.ID;
            List<Meeting> QResultList = new List<Meeting>();
            var mService = new WebChannel<IMeetingMngService>();
            mService.Use(service => service
            .GetMeetingList(cond)
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new Meeting();
                    item.MTID = d.MTID;
                    item.MTName = d.MTName;
                    item.MTStartDate = d.MTStartDate;
                    item.MTConvenerName = d.MTConvenerName;
                    item.MeetingReadType = cond.MeetingReadType;
                    item.MTActive = d.MTActive;
                    QResultList.Add(item);
                }
            }));
            var gridKey = mService.DataToCache(QResultList.AsEnumerable());
            SetGridKey("MMQU001Grid", gridKey);
        }
        /// <summary>
        /// 查詢結果資料綁定jquery處理
        /// </summary>
        /// <param name="jqParams"></param>
        /// <returns></returns>
        [HasPermission("EP.PSL.WorkResources.MeetingMng.MMQU001")]
        public JsonResult BindGrid(jqGridParam jqParams)
        {
            //取得CacheKey
            var key = GetGridKey("MMQU001Grid");
            return BaseGridBinding<Meeting>(jqParams,
                () => new WebChannel<IMeetingMngService, Meeting>().Get(key));
        }
    }
}