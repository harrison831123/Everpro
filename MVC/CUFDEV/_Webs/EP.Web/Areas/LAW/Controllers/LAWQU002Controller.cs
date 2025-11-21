using EP.H2OModels;
using EP.SD.SalesSupport.LAW.Models;
using EP.SD.SalesSupport.LAW.Service;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Controllers
{
    [Program("LAWQU002")]
    public class LAWQU002Controller : BaseController
    {
        // GET: LAW/LAWQU002
        private ILAWService _Service;
        public LAWQU002Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }

        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU002")]
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 查詢作業
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWQU002")]
        public void Query(LawSearchDetail model)
        {
            //取資料
            var mService = new WebChannel<ILAWService>();
            List<LawSearchDetail> Viewmodel = new List<LawSearchDetail>();
            WebChannel<ILAWService> _channelService = new WebChannel<ILAWService>();
            if (model.LawSearchType == "1")
            {
                _channelService.Use(service => Viewmodel = service.GetLawContentBySearch(model));
                for (int i = 0; i < Viewmodel.Count; i++)
                {
                    if (Viewmodel[i].LawStatusType == 2)
                    {
                        List<LawCloseType> lawCloseTypes = _Service.GetLawCloseTypeByLawCloseType(Viewmodel[i].LawCloseType);
                        for (int j = 0; j < lawCloseTypes.Count; j++)
                        {
                            Viewmodel[i].CloseTypeName = j == 0 ? "●" + lawCloseTypes[j].CloseTypeName : Viewmodel[i].CloseTypeName + "●" + lawCloseTypes[j].CloseTypeName;
                        }
                    }
                }
            }
            else if (model.LawSearchType == "2")
            {
                LawSys lawSys = _Service.GetLawSysBySysMemberID(User.MemberInfo.ID);
                OrgVm orgvm = _Service.GetOrfVm(User.AccountInfo.ID);
                if (orgvm == null)
                {
                    orgvm = new OrgVm();
                    orgvm.vm_flag = 0;
                    orgvm.vsm_flag = 0;
                    orgvm.vmleaderid = User.AccountInfo.ID;
                }

                mService.Use(service => service
                .GetLawContentByNotClose(model, orgvm, lawSys)
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
                LawSys lawSys = _Service.GetLawSysBySysMemberID(User.MemberInfo.ID);
                OrgVm orgvm = _Service.GetOrfVm(User.AccountInfo.ID);
                if (orgvm == null)
                {
                    orgvm = new OrgVm();
                    orgvm.vm_flag = 0;
                    orgvm.vsm_flag = 0;
                    orgvm.vmleaderid = User.AccountInfo.ID;
                }
                mService.Use(service => service
                .GetLawContentByClose(model, orgvm, lawSys)
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

        /// <summary>
        /// 執行程序
        /// </summary>
        /// <param name="LawId"></param>
        /// <param name="LawNoteNo"></param>
        /// <returns></returns>
        public string GetLawDoProgressByLawIdNotoNo(string LawId, string LawNoteNo)
        {
            string result = string.Empty;
            List<LawDoProgress> lawDoProgresses = _Service.GetLawDoProgressByLawIdNotoNo(LawId, LawNoteNo);

            for (int j = 0; j < lawDoProgresses.Count; j++)
            {
                result = j == 0 ? "●" + lawDoProgresses[j].LawDoprogress + "<br />" : result + "●" + lawDoProgresses[j].LawDoprogress + "<br />";
            }
            return result;
        }

        /// <summary>
        /// 依照條件抓取報表
        /// </summary>
        /// <param name="productionY">業績年</param>
        [HttpPost]
        public JsonResult GetSearchReport(LawSearchDetail model)
        {
            //if (!string.IsNullOrEmpty(model.ProductionYM))
            //{
            //    model.ProductionYM = model.ProductionYM;
            //}
            //if (!string.IsNullOrEmpty(model.ProductionY))
            //{
            //    model.ProductionY = model.ProductionY;
            //}
            //IWorkbook workBook = new XSSFWorkbook();
            //ISheet sheet = workBook.CreateSheet("法追系統-查詢明細");
            //ICellStyle cellStyle = workBook.CreateCellStyle();

            string fileName = string.Empty;
            byte[] data = null;
            var service = ServiceHelper.Create<ILAWService>();
            switch (model.LawSearchType)
            {
                case "1":
                    fileName = "法追系統-查詢明細" + ".xls";
                    data = service.GetSearchReportList(model);
                    break;
                case "2":
                    fileName = "法追系統-未結案件明細" + ".xls";
                    data = service.GetNotCloseReportList(model);
                    break;
                case "3":
                    fileName = "法追系統-已結案件明細" + ".xls";
                    data = service.GetCloseReportList(model);
                    break;
            }

            string handle = Guid.NewGuid().ToString();
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

        /// <summary>
        /// 下載輸出
        /// </summary>
        /// <param name="fileGuid">guid</param>
        /// <param name="fileName">檔名</param>
        /// <returns></returns>
        [HttpGet]
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