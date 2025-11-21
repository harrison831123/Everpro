using EP.H2OModels;
using EP.SD.SalesSupport.LAW.Models;
using EP.SD.SalesSupport.LAW.Service;
using EP.SD.SalesSupport.LAW.Web.Areas.LAW.Utilities;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Controllers
{
    [Program("LAWQU004")]
    public class LAWQU004Controller : BaseController
    {
        // GET: LAW/LAWQU004
        private ILAWService _Service;
        public LAWQU004Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }
        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU004")]
        public ActionResult Index()
        {
            //LawSys lawSys = _Service.GetLawSysBySysMemberID(User.MemberInfo.ID);
            //LawSearchDetail searchDetail = new LawSearchDetail();
            //searchDetail.SysType = lawSys == null ? 0 : 1;
            //searchDetail.UserName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
            //return View(searchDetail);
            return View();
        }

        /// <summary>
        /// 法追明細
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU004")]
        public void Query(LawSearchDetail model)
        {
            //取資料
            var mService = new WebChannel<ILAWService>();

            List<LawSearchDetail> Viewmodel = new List<LawSearchDetail>();
            WebChannel<ILAWService> _channelService = new WebChannel<ILAWService>();
            OrgVm orgvm = _Service.GetOrfVm(User.AccountInfo.ID);
            if (orgvm == null)
            {
                orgvm = new OrgVm();
                orgvm.vm_flag = 0;
                orgvm.vsm_flag = 0;
                orgvm.vmleaderid = User.AccountInfo.ID;
            }

            if (model.LawSearchType == "1")
            {
                mService.Use(service => service
                .GetLawSearchByNotClose(model, orgvm)
                .ForEach(d =>
                {
                    if (d != null)
                    {
                        //var item = new LawSearchDetail();
                        LawRepaymentList lawRepaymentList = _Service.GetLawRepaymentListByAgID(d.LawDueAgentId.Substring(0, 10), d.LawId.ToString());
                        d.LawRepaymentMoney = lawRepaymentList != null ? lawRepaymentList.LawRepaymentMoney : 0;

                        LawEvidenceDesc lawEvidenceDesc = _Service.GetLawEvidenceDescByLawId(d.LawId.ToString());
                        d.LawEvidencedesc = lawEvidenceDesc != null ? lawEvidenceDesc.LawEvidencedesc : "";
                        Viewmodel.Add(d);
                    }
                }));
            }
            else
            {
                mService.Use(service => service
                .GetLawSearchByClose(model, orgvm)
                .ForEach(d =>
                {
                    if (d != null)
                    {
                        //var item = new LawSearchDetail();
                        LawRepaymentList lawRepaymentList = _Service.GetLawRepaymentListByAgID(d.LawDueAgentId.Substring(0, 10), d.LawId.ToString());
                        d.LawRepaymentMoney = lawRepaymentList != null ? lawRepaymentList.LawRepaymentMoney : 0;

                        LawEvidenceDesc lawEvidenceDesc = _Service.GetLawEvidenceDescByLawId(d.LawId.ToString());
                        d.LawEvidencedesc = lawEvidenceDesc != null ? lawEvidenceDesc.LawEvidencedesc : "";
                        Viewmodel.Add(d);
                    }
                }));
            }

            var gridKey = _channelService.DataToCache(Viewmodel.AsEnumerable());
            SetGridKey("QueryGrid", gridKey);
        }

        public JsonResult BindGrid(jqGridParam jqParams)
        {
            var cacheKey = GetGridKey("QueryGrid");
            return BaseGridBinding<LawSearchDetail>(jqParams,
                () => new WebChannel<ILAWService, LawSearchDetail>().Get(cacheKey));
        }

        /// <summary>
        /// 訴訟程序
        /// </summary>
        /// <param name="LawId"></param>
        /// <param name="LawNoteNo"></param>
        /// <returns></returns>
        public string GetLawLitigationProgressByLawIdNotoNo(string LawId, string LawNoteNo)
        {
            List<LawLitigationProgress> lawLitigationProgresses = _Service.GetLawLitigationProgressByLawIdNotoNo(LawId, LawNoteNo);
            string result = string.Empty;
            for (int j = 0; j < lawLitigationProgresses.Count; j++)
            {
                result = j == 0 ? "●" + lawLitigationProgresses[j].LawLitigationprogress + "<br />" : result + "●" + lawLitigationProgresses[j].LawLitigationprogress + "<br />";
            }
            return result;
        }
    }
}