using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.Mvc;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using Newtonsoft.Json;
using System.Data;
using System.Net.NetworkInformation;
using EP.SD.SalesSupport.CUSCRM.Service;
using EP.Web;
using System.Linq;

namespace EP.SD.SalesSupport.CUSCRM.Web.Controllers
{
    /// <summary>
    /// 維護介面
    /// </summary>
    [Program("CUSCRMTX003")]
    public class CUSCRMTX003Controller : BaseController
    {
        private ICUSCRMTX003Service _crmService;
        public CUSCRMTX003Controller()
        {
            _crmService = ServiceHelper.Create<ICUSCRMTX003Service>();
        }

        /// <summary>
        /// 預設查詢未結案的所有受理單
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        public ActionResult Index()
        {
            QueryMaintainCondition condition = new QueryMaintainCondition();
            List<MaintainInfo> infos = _crmService.GetMtnData(condition);

            return View(infos);
        }

        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        [HttpPost]
        [PdLogFilter("EIP客服-維護-查詢", PITraceType.Query)]
        public ActionResult Index(QueryMaintainCondition condition)
        {
            List<MaintainInfo> infos = _crmService.GetMtnData(condition);

            return View(infos);
        }

        /// <summary>
        /// 取歷史維護紀錄資料
        /// </summary>
        /// <param name="no">受理編號</param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        [HttpPost]
        public JsonResult GetMaintainData(string no) 
        {
            MtnHistoryInfo mtnHistoryInfo = _crmService.GetDoHistory(no);

            return Json(mtnHistoryInfo);
        }

        /// <summary>
        /// 執行稽催
        /// </summary>
        /// <param name="no">受理編號</param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        [HttpPost]
        [PdLogFilter("EIP客服-維護-稽催", PITraceType.Insert)]
        public string DoAudit(string no)
        {
            CRMEAudit cRMEAudit = new CRMEAudit();
            cRMEAudit.No = no;
            cRMEAudit.Type = 1;
            cRMEAudit.Date = DateTime.Now;
            cRMEAudit.Content = "";
            cRMEAudit.Creator = User.MemberInfo.ID;
            cRMEAudit.CreateTime = DateTime.Now;
            if (_crmService.DoAudit(cRMEAudit, PlatformHelper.GetClientIPv4(), User.MemberInfo.Name, User.MemberInfo.ID))
            {
                AppendMessage("新增稽催成功");
                return "success";
            }
            else 
            {
                AppendMessage("新增稽催失敗");
                return "faild";
            }
        }

        /// <summary>
        /// 執行催辦
        /// </summary>
        /// <param name="no"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        [HttpPost]
        [PdLogFilter("EIP客服-維護-催辦", PITraceType.Insert)]
        public string DoSData(string no, string content)
        {
            CRMEDoS cRMEDoS = new CRMEDoS();
            cRMEDoS.No = no;
            cRMEDoS.Content = content;
            cRMEDoS.Creator = User.MemberInfo.ID;
            cRMEDoS.CreateTime = DateTime.Now;
            if (_crmService.DoS(cRMEDoS, PlatformHelper.GetClientIPv4(), User.MemberInfo.Name, User.MemberInfo.ID))
            {
                AppendMessage("新增催辦成功");
                return "success";
            }
            else
            {
                AppendMessage("新增催辦失敗");
                return "faild";
            }
        }

        /// <summary>
        /// 產生處理單
        /// </summary>
        /// <param name="no">受理編號</param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        [HttpPost]
        [PdLogFilter("EIP客服-維護-處理單", PITraceType.Download)]
        public ActionResult RenderProcessForm(string no)
        {
            ProcessFormData model =_crmService.ProduceProcessFormData(no);
            ProcessFormViewModel viewModel = new ProcessFormViewModel();
            viewModel.AgentHistory = model.AgentHistory;
            viewModel.CaseHandler = model.CaseHandler;
            viewModel.CloseRecord = model.CloseRecord;
            viewModel.ContactRecord = model.ContactRecord;
            viewModel.ContentTXT = model.ContentTXT;
            viewModel.CreatTime = model.CreatTime;
            viewModel.DepositLetter = model.DepositLetter;
            viewModel.depositLetterDateTime = model.depositLetterDateTime;
            viewModel.No = model.No;
            viewModel.Owner = model.Owner;
            viewModel.PolicyHistory = model.PolicyHistory;
            viewModel.ProcessRecord = model.ProcessRecord;
            viewModel.Source = model.Source;

            return View("ProcessSheet", viewModel);
        }

        /// <summary>
        /// 取得維護紀錄歷史資料
        /// </summary>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        [PdLogFilter("EIP客服-維護", PITraceType.Query)]
        [HttpPost]
        public bool MtnRecord()
        {
            bool result = false;
            string no = "";
            string content = "";
            DateTime? busContactCompanyDate = null;
            DateTime? replyCompanyDate = null;
            DateTime? replyYallowBillDate = null;
            DateTime? ccReplyDate = null;
            int? companyDiscipType = null;
            string type = Request.Form["type"].ToString();
            if (type.ToUpper() == "CS")
            {
                no = Request.Form["csSpanNo"];
                content = Request.Form["csContent"];
            }
            else if (type.ToUpper() == "CC") {
                no = Request.Form["ccSpanNo"];
                content = Request.Form["ccContent"];

                //回文日
                if (!String.IsNullOrEmpty(Request.Form["ccReplyDate"])) {
                    ccReplyDate = DateTime.Parse(Request.Form["ccReplyDate"].ToString());
                }
            }
            else if (type.ToUpper() == "SS")
            {
                no = Request.Form["ssSpanNo"];
                content = Request.Form["ssContent"];

                //業連保險公司日期
                if (!String.IsNullOrEmpty(Request.Form["busContactCompanyDate"]))
                {
                    busContactCompanyDate = DateTime.Parse(Request.Form["busContactCompanyDate"].ToString());
                }
                //保險公司回覆日
                if (!String.IsNullOrEmpty(Request.Form["replyCompanyDate"]))
                {
                    replyCompanyDate = DateTime.Parse(Request.Form["replyCompanyDate"].ToString());
                }
                //回覆單位黃聯日
                if (!String.IsNullOrEmpty(Request.Form["replyYallowBillDate"]))
                {
                    replyYallowBillDate = DateTime.Parse(Request.Form["replyYallowBillDate"].ToString());
                }
                //保公決議選項
                if (!String.IsNullOrEmpty(Request.Form["companyDiscipType"]))
                {
                    companyDiscipType = Int32.Parse(Request.Form["companyDiscipType"].ToString());
                }
            }
            else if (type.ToUpper() == "SC")
            {
                no = Request.Form["scSpanNo"];
                content = Request.Form["scContent"];
            }
            string resultDiscipType = Request.Form["resultDiscipType"];
            string resultDiscipType2 = Request.Form["resultDiscipType2"];
            for (int i = 0; i < Request.Files.Count; i++)
            {
                HttpPostedFileBase file = Request.Files[i];
                if (file.ContentLength > 0) 
                {
                    string extension = Path.GetExtension(file.FileName);
                    string mdfFileName = Guid.NewGuid().ToString("N") + extension;
                    string savePath = Path.Combine(PlatformHelper.GetVarConfig("CUSCRMDir"), no);
                    if (!Directory.Exists(savePath))
                    {
                        Directory.CreateDirectory(savePath);
                    }
                    file.SaveAs(Path.Combine(savePath, mdfFileName));

                    CRMEFile cRMEFile = new CRMEFile();
                    cRMEFile.SourceType = 1;
                    cRMEFile.FolderNo = no;
                    cRMEFile.FileNo = "CUSCRMTX003";
                    cRMEFile.FileName = file.FileName;
                    cRMEFile.FileMD5Name = mdfFileName;
                    cRMEFile.Creator = User.MemberInfo.ID;
                    cRMEFile.CreateTime = DateTime.Now;
                    result = _crmService.AddUploadFile(cRMEFile);
                }
            }
            CRMEDo cRMEDo = new CRMEDo();
            cRMEDo.No = no;
            cRMEDo.Content = content;
            cRMEDo.Creator = User.MemberInfo.ID;
            cRMEDo.BUSContactCompanyDate = busContactCompanyDate;
            cRMEDo.ReplyCompanyDate = replyCompanyDate;
            cRMEDo.ReplyYallowBillDate = replyYallowBillDate;
            cRMEDo.CCReplyDate = ccReplyDate;
            cRMEDo.CreateTime = DateTime.Now;
            cRMEDo.CompanyResult = companyDiscipType;
            result = _crmService.InserCRMEDo(cRMEDo);

            return result;
        }

        /// <summary>
        /// 結案
        /// </summary>
        /// <param name="no">受理編號</param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        [HttpPost]
        [PdLogFilter("EIP客服-維護-結案", PITraceType.Update)]
        public bool MtnClose()
        {
            string type = Request.Form["type"].ToString();
            int minNum = 0;

            CRMECloseLog cRMECloseLog = new CRMECloseLog();
            cRMECloseLog.No = Request.Form[(type+"SpanNo")];

            //上傳檔案
            for (int i = 0; i < Request.Files.Count; i++)
            {
                HttpPostedFileBase file = Request.Files[i];
                if (file.ContentLength > 0)
                {
                    string extension = Path.GetExtension(file.FileName);
                    string mdfFileName = Guid.NewGuid().ToString("N") + extension;
                    string savePath = Path.Combine(PlatformHelper.GetVarConfig("CUSCRMDir"), cRMECloseLog.No);
                    if (!Directory.Exists(savePath))
                    {
                        Directory.CreateDirectory(savePath);
                    }
                    file.SaveAs(Path.Combine(savePath, mdfFileName));

                    CRMEFile cRMEFile = new CRMEFile();
                    cRMEFile.SourceType = 1;
                    cRMEFile.FolderNo = cRMECloseLog.No;
                    cRMEFile.FileNo = "CUSCRMTX003";
                    cRMEFile.FileName = file.FileName;
                    cRMEFile.FileMD5Name = mdfFileName;
                    cRMEFile.Creator = User.MemberInfo.ID;
                    cRMEFile.CreateTime = DateTime.Now;
                    _crmService.AddUploadFile(cRMEFile);
                }
            }

            //寫入結案
            if (!String.IsNullOrEmpty(Request.Form["resultDiscipType"]))
            {
                cRMECloseLog.ResultCode = Int32.TryParse(Request.Form["resultDiscipType"].ToString(), out minNum) ? Int32.Parse(Request.Form["resultDiscipType"].ToString()) : minNum;
            }
            if (!String.IsNullOrEmpty(Request.Form["resultDiscipType2"]))
            {
                cRMECloseLog.ResultCode2 = Int32.TryParse(Request.Form["resultDiscipType2"].ToString(), out minNum) ? Int32.Parse(Request.Form["resultDiscipType2"].ToString()) : minNum;
            }
            cRMECloseLog.Creator = User.MemberInfo.ID;
            cRMECloseLog.CreateTime = DateTime.Now;
            var result = _crmService.InsertCloseLog(cRMECloseLog);
            
            return result;
        }

        /// <summary>
        /// 依據ID刪除檔案
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        [HttpPost]
        [PdLogFilter("EIP客服-維護-刪除檔案", PITraceType.Delete)]
        public bool DelFile(string values)
        {
            List<int> list = new List<int>();
            int minInt = 0;
            foreach (string v in values.Split(','))
            {
                if (Int32.TryParse(v, out minInt))
                {
                    list.Add(Int32.Parse(v));
                }
            }
            return _crmService.DeleteCRMEFileDbAndFileById(list);
        }

        /// <summary>
        /// 下載檔案
        /// </summary>
        /// <param name="id">檔案ID</param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        [PdLogFilter("EIP客服-維護-下載檔案", PITraceType.Download)]
        public ActionResult FileDownload(int id)
        {
            CRMEFile cRMEFile = _crmService.GetCRMEFileById(id);
            var filePath = Path.Combine(PlatformHelper.GetVarConfig("CUSCRMDir"), cRMEFile.FolderNo);
            var fullFileName = Path.Combine(filePath, cRMEFile.FileMD5Name);
            var downloadName = cRMEFile.FileName;
            if (System.IO.File.Exists(fullFileName))
            {
                byte[] fileBytes = System.IO.File.ReadAllBytes(fullFileName);
                return File(fileBytes, "application/octet-estream", downloadName);
            }
            else
            {
                AppendMessage("檔案不存在", true);
                return View("Index");
            }
        }

        /// <summary>
        /// 檢核該業務員是否已離職
        /// </summary>
        /// <param name="dosNo">受理編號</param>
        /// <returns>true: 已離職; false: 在職</returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        [HttpPost]
        public bool CheckMemberIsLeave(string dosNo)
        {
            var _notifyService = ServiceHelper.Create<INotifyService>();
            var policy = _notifyService.GetCRMEInsurancePolicy(dosNo).FirstOrDefault();

            // 取得人員資料
            MemberCondition condition = new MemberCondition();
            condition.MemberID = policy.SUAgentCode == null ? policy.AgentCode : policy.SUAgentCode;

            var organizationProvider = ServiceHelper.Create<IOrganizationProvider>();
            var members = organizationProvider.QueryMembers(condition).Where(m => (!m.LeaveOfficeDate.HasValue || m.LeaveOfficeDate > DateTime.Now));  //離職日大於系統日，表示在職

            if (members.Count() > 0)
            {
                return false;
            }
            else {
                return true;
            }
        }

        /// <summary>
        /// 檢核單位是否已裁撤
        /// </summary>
        /// <param name="wcCCode">單位編號</param>
        /// <returns>true: 已裁撤; false: 尚未裁撤</returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMTX003")]
        [HttpPost]
        public bool CheckWCCenterCodeIsLife(string dosNo)
        {
            // 用行專與實駐對應表檢核單位是否還在
            var _notifyService = ServiceHelper.Create<INotifyService>();
            var policy = _notifyService.GetCRMEInsurancePolicy(dosNo).FirstOrDefault();
            var wcCCode = policy.SUAgentCode == null ? policy.WCCode : policy.SUWCCode;
            var data = _notifyService.GetCRMENotifyEmployee(wcCCode).ToList();

            if (data.Count() > 0)  //單位還在
            {
                return false;
            }
            else  //單位裁撤
            {
                return true;
            }
        }
    }
}
