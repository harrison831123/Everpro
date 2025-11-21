using EP.H2OModels;
using EP.Platform.Service;
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
    [Program("LAWQU001")]
    public class LAWQU001Controller : BaseController
    {
        // GET: LAW/LAWQU001
        private ILAWService _Service;
        public LAWQU001Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }

        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU001")]
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 通知作業查詢
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU001")]
        public void Query(LawNote model)
        {
            string MemberID = User.MemberInfo.ID;
            string accountID = User.AccountInfo.ID;
            int CheckSys;
            CheckSys = _Service.CheckLawNoteByMemberID(MemberID);
            int sys = CheckSys > 0 ? 1 : 0;
            //產生新的chkid
            Random outerRnd = new Random(Guid.NewGuid().GetHashCode());
            string chkid = DateTime.Now.ToString("yyyyMMddHHmmss") + outerRnd.Next(9000, 10001).ToString();

            //寫入BPM隨機參數
            var homeService = Microsoft.CUF.Framework.Service.ServiceHelper.Create<IHomeService>();
            homeService.UpdateBPMRanNum(chkid, Session["orgID"].ToString());

            //取資料
            List<LawNote> list = new List<LawNote>();
            List<LawNoteDetail> Viewmodel = new List<LawNoteDetail>();
            WebChannel<ILAWService> _channelService = new WebChannel<ILAWService>();
            var memberService = new WebChannel<IMemberExtendService>();
            int? orgID = null;
            memberService.Use(service => orgID = service.GetOrgidById(accountID));
            _channelService.Use(service => list = service.GetLawNote(model.LawNoteNo, sys, MemberID, model.LawNoteName, orgID.ToString()));

            //照會單列表
            for(int i =0;i< list.Count; i++)
            {
                LawNoteDetail noteDetail = new LawNoteDetail();
                string LawNoteNoYear;
                if (list[i].LawNoteNo.Length == 8)
                {
                    LawNoteNoYear = list[i].LawNoteNo.Substring(0, 2);
                    noteDetail.BPMURL = GetNoteExport(LawNoteNoYear, list[i].LawNoteNo, chkid, sys);
                }
                else
                {
                    LawNoteNoYear = list[i].LawNoteNo.Substring(0, 3);
                    noteDetail.BPMURL = GetNoteExport(LawNoteNoYear, list[i].LawNoteNo, chkid, sys);
                }
                noteDetail.LawNoteNo = list[i].LawNoteNo;
                noteDetail.LawNoteName = list[i].LawNoteName;
                noteDetail.LawNoteCreatedate = list[i].LawNoteCreatedate;
                noteDetail.LawNoteId = list[i].LawNoteId;
                Viewmodel.Add(noteDetail);
            }

            var gridKey = _channelService.DataToCache(Viewmodel.AsEnumerable());
            SetGridKey("QueryGrid", gridKey);
        }

        public JsonResult BindGrid(jqGridParam jqParams)
        {
            var cacheKey = GetGridKey("QueryGrid");
            return BaseGridBinding<LawNoteDetail>(jqParams,
                () => new WebChannel<ILAWService, LawNoteDetail>().Get(cacheKey));
        }

        /// <summary>
        /// 取得照會單連結
        /// </summary>
        /// <param name="LawNoteNoYear"></param>
        /// <param name="LawNoteNo"></param>
        /// <param name="chkid"></param>
        /// <param name="sys"></param>
        /// <returns></returns>
        public string GetNoteExport(string LawNoteNoYear,string LawNoteNo,string chkid,int sys)
        {
            string BPMURL = System.Web.Configuration.WebConfigurationManager.AppSettings["BPMLAWNoteURL"];
            string BPMURL2 = System.Web.Configuration.WebConfigurationManager.AppSettings["BPMLAWNoteURL2"];
            if(Convert.ToInt32(LawNoteNoYear) <= 107)
            {
                BPMURL = BPMURL + "&crm_no={0}" + "&chkid={1}" + "&sys={2}";               
                return string.Format(BPMURL, LawNoteNo, chkid, sys); ;
            }
            else
            {
                BPMURL2 = BPMURL2 + "&crm_no={0}" + "&chkid={1}" + "&sys={2}";                
                return string.Format(BPMURL2, LawNoteNo, chkid, sys); ;
            }
        }
    }
}