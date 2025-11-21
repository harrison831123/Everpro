using EB.Common;
using EB.Platform.Service;
using EB.SL.MerSal.Models;
using EB.SL.MerSal.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EB.SL.MerSal.Web.Areas.MerSal.Controllers
{
    [Program("MERSALTX001")]
    public class MerSalTX001Controller : BaseController
    {
        private IMerSalService _service;
        private IVlifeService _vlifeService;
        private static string _programID = "MerSalTX001";

        public MerSalTX001Controller()
        {
            _service = ServiceHelper.Create<IMerSalService>();
            _vlifeService = ServiceHelper.Create<IVlifeService>();
        }
        // GET: MerSal/MerSalTX001
        [HasPermission("EB.SL.MerSal.MerSalTX001")]

        public ActionResult Index()
        {
            return View();
        }


        /// <summary>
        /// 取得發佣原始檔狀態
        /// </summary>
        /// <param name="ProductionYM"></param>
        /// <param name="Sequence"></param>
        /// <param name="CompanyCode"></param>
        /// <returns></returns>
        public string CheckSeqNo(string ProductionYM, string Sequence, string CompanyCode)
        {

            string SeqNo = _service.GetSeqNo(StringExtension.WYearMonthToCYearMonth(ProductionYM), Sequence, CompanyCode);

            if (!String.IsNullOrEmpty(SeqNo))
            {
                return SeqNo;
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 取得佣酬檢核狀態為「3-資料入佣酬調整
        /// </summary>
        /// <param name="ProductionYM"></param>
        /// <param name="Sequence"></param>
        /// <param name="CompanyCode"></param>
        /// <returns></returns>
        public string GetPSMSR(string ProductionYM, string Sequence, string CompanyCode)
        {
            string result = _service.Getprocess_status_MSRC(StringExtension.WYearMonthToCYearMonth(ProductionYM), Sequence, CompanyCode) ?? "";
            return result;
        }

        #region Batch

        public static string BatchFileFolder = System.Web.Configuration.WebConfigurationManager.AppSettings["BatchFileFolder"] ?? @"D:\Microsoft\ShareFileFolder\BatchFiles\";
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalTX001")]
        //[EB.Web.PdLogFilter("MerSal", EB.Web.PITraceType.Download)]
        public JsonResult RealtimeBatch(string ProductionYm, string FileSeq, string Sequence, string programName, string CompanyCode)
        {
            bool isJobOK = false;
            string backValue = "";
            string CreateUserCode = User.AccountInfo.ID;
            string[] sArray = ProductionYm.Split('/');

            //取得cm_batch_control資料檢核是否還在執行(防呆)
            string GetCmBatchControlR = _service.Getcm_batch_controlR(ProductionYm, Sequence, CompanyCode) ?? "";

            if (GetCmBatchControlR != "0")
            {
                backValue = "CmBatchControlR";
            }
            else
            {
                //佣酬檢核_業績年月+序號_保險公司代碼_YYYYMMDDHHmmss.xlsx
                string filename = "佣酬檢核_" + sArray[0] + sArray[1] + Sequence + "_" + CompanyCode + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                //單筆
                string jsonAssembly = "{'ASSEMBLY_NAME':'EB.SL.MerSal.Service','CLASS_NAME':'EB.SL.MerSal.Service.MerSalService','INVOKE_METHOD':'BatchQueryReport'}";
                //要呼叫的參數
                string jsonParamaters = Newtonsoft.Json.JsonConvert.SerializeObject("{'Sequence':'" + Sequence + "', 'ProductionYm':'" + StringExtension.WYearMonthToCYearMonth(ProductionYm)
                    + "', 'FileSeq': '" + FileSeq + "', 'CompanyCode': '" + CompanyCode + "', 'CreateUserCode': '" + CreateUserCode + "'}");

                //報表相關參數
                string jsonReportInfo = "{'reportFilename':'" + filename + "', 'reportType':'Report'}";
                programName = programName + sArray[0] + sArray[1] + Sequence + "-" + CompanyCode;
                new WebChannel<IBatchService>()
                            .Use(proxy => { isJobOK = proxy.RealtimeBatchJobService(_programID, programName, User.Identity.Name, User.Identity.Name, jsonAssembly, jsonParamaters, jsonReportInfo); });
                if (isJobOK)
                {
                    backValue = "JobOK";
                }
                else
                {
                    backValue = "JobNotOk";
                }
            }

            return new JsonResult()
            {
                Data = new
                {
                    DataValue = backValue
                    ,
                    DataStatus = isJobOK
                }
            };

            //return Json(isJobOK);
        }

        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalTX001")]
        public JsonResult BatchQuery(BatchQuerySearchModel searchModel, string CompanyCodeT)
        {
            searchModel.ProgramId = "MerSalTX001";
            IEnumerable<BatchQueryGridModel> dataList = null;
            //呼叫函式取得資料
            new WebChannel<IBatchQueryService>()
            .Use(proxy => { var entityList = proxy.BatchQuery(searchModel);
                //取得資料轉為ViewModel
                dataList = entityList.Select(m =>
                    new BatchQueryGridModel
                    {
                        fbatch_status = m.fbatch_status,
                        fbatch_status_ch = string.IsNullOrEmpty(m.fbatch_status) ? "" : ((BatchStatus)Enum.Parse(typeof(BatchStatus), m.fbatch_status, false)).GetName(),
                        ibatch_no = m.ibatch_no,
                        ffunction_no = m.ffunction_no,
                        imember = Member.Get(Account.Get(m.imember).MemberID).Name,
                        icreate = m.icreate,
                        dstart = !m.dstart.HasValue ? "" : showStartDate(m.dstart.Value),
                        dend = m.dend.HasValue ? m.dend.Value.GetString("yyyy/MM/dd HH:mm:ss") : "",
                        nreport_name = m.nreport_name,
                        nfile_name = m.nfile_name,
                        dupdate = m.dupdate.HasValue ? m.dupdate.Value.GetString("yyyy/MM/dd HH:mm:ss") : "",
                    //qfile_no = m.qfile_no.HasValue ? ((int)m.qfile_no.Value).ToString() : "",
                    ffile_size = m.ffile_size.ToString().Replace("$$", "<br/>"),
                        nreport_filename = m.nreport_filename,
                        is_deleteable = (m.icreate == User.MemberInfo.ID || m.imember == User.MemberInfo.ID || User.HasPermission("EB.Platform.BatchQuery.Maintain")) ? "Y" : "N"

                    }).Where(c => c.nreport_name.Substring(6, 11) == CompanyCodeT.Replace("/", "")).ToList();//依檔名的年月+序號+-+保公代碼撈取顯示
                                                                                                             //}).Where(c => c.nreport_name.Split('-')[1] == CompanyCodeT).ToList();
            });
            

            //將ViewModel存入Cache
            //var cacheKey = new WebChannel<IBatchQueryService, BatchQueryGridModel>().DataToCache(dataList.ToList());
            //設定Gird資料的cacheKey
            //SetGridKey("Batch_cacheKey", cacheKey);

            return Json(dataList);
        }

        private string showStartDate(DateTime dstart)
        {
            return (dstart - DateTime.Now).TotalSeconds > 0 ? "<font color=blue>" + dstart.GetString("yyyy/MM/dd HH:mm:ss") + "</font>" : dstart.GetString("yyyy/MM/dd HH:mm:ss");
        }

        [HasPermission("EB.SL.MerSal.MerSalTX001")]
        public ActionResult FileDownload(string nfile_name, string nreport_filename)
        {
            string fullFileName = Path.Combine(BatchFileFolder, nfile_name);
            if (!System.IO.File.Exists(fullFileName))
            {
                Throw.LogError("批次查詢：此檔案不存在於系統中,file：" + fullFileName);
                return Content("<script>alert('此檔案不存在於系統中');</script>");
            }
            //將Client檔名自動加上流水號
            return File(fullFileName, "application/octet-estream", Url.Encode(DataHelper.AddFileUniqueDownloadName(nreport_filename)));//中文檔名編碼
        }
        #endregion

        #region 檢核執行紀錄查詢

        /// <summary>
        /// 檢核執行紀錄
        /// </summary>
        /// <param name="MerSalViewModel"></param>
        [HttpPost]
        [HasPermission("EB.SL.MerSal.MerSalTX001")]
        public JsonResult Query(MerSalViewModel model)
        {
            //取資料
            List<MerSalViewModel> list = new List<MerSalViewModel>();
            WebChannel<IMerSalService> _channelService = new WebChannel<IMerSalService>();

            _channelService.Use(service => list = service.GetMerSalRun(StringExtension.WYearMonthToCYearMonth(model.ProductionYM), model.Sequence.ToString(), model.CompanyCode));
            for (int i = 0; i < list.Count; i++)
            {
                list[i].CreateUserCode = !String.IsNullOrEmpty(list[i].CreateUserCode) ? Member.Get(Account.Get(list[i].CreateUserCode).MemberID).Name : "";
            }
            //var gridKey = _channelService.DataToCache(list.AsEnumerable());
            //SetGridKey("QueryGrid", gridKey);
            return Json(list);
        }

        #endregion

    }
}