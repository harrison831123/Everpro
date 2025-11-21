using EP.H2OModels;
using EP.Platform.Service;
using EP.PSL.IB.Service;
using EP.SD.SalesSupport.LAW.Models;
using EP.SD.SalesSupport.LAW.Service;
using EP.SD.SalesSupport.LAW.Web.Areas.LAW.Model;
using EP.SD.SalesSupport.LAW.Web.Areas.LAW.Utilities;
using EP.Web;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Controllers
{
    [Program("LAWTX009")]
    public class LAWTX009Controller : BaseController
    {
        // GET: LAW/LAWTX009
        private ILAWService _Service;
        public LAWTX009Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX009")]
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 維護作業查詢
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public void Query(LawContent model)
        {
            //取資料
            List<LawContent> list = new List<LawContent>();
            List<LawContentDetail> Viewmodel = new List<LawContentDetail>();
            WebChannel<ILAWService> _channelService = new WebChannel<ILAWService>();
            _channelService.Use(service => list = service.GetLawContent(model));

            //列表
            for (int i = 0; i < list.Count; i++)
            {
                LawContentDetail contentDetail = new LawContentDetail();
                contentDetail.LawId = list[i].LawId;
                contentDetail.LawNoteNo = list[i].LawNoteNo;
                contentDetail.LawDueName = list[i].LawDueName;
                contentDetail.LawDueAgentId = list[i].LawDueAgentId;
                contentDetail.LawDueMoney = list[i].LawDueMoney;
                contentDetail.LawProductionYM = list[i].LawYear + "/" + list[i].LawMonth;
                contentDetail.LawDoUnitName = list[i].LawDoUnitName;
                contentDetail.LawStatusType = list[i].LawStatusType;
                contentDetail.LawCloseDate = list[i].LawCloseDate.ToString("yyyy/MM/dd");
                switch (list[i].LawStatusType)
                {
                    case 1: //維護
                        if (list[i].LawContentCancelType == "1")
                        {
                            contentDetail.LawNotCloseTypeName = "結案";
                        }
                        else
                        {
                            if (!String.IsNullOrEmpty(list[i].LawNotCloseTypeName))
                            {
                                contentDetail.LawNotCloseTypeName = list[i].LawNotCloseTypeName;
                            }
                            else
                            {
                                contentDetail.LawNotCloseTypeName = "未結案";
                            }
                        }
                        break;
                    case 2: //已結
                        if (list[i].LawContentCancelType == "1" || list[i].LawCloseDate.ToString("yyyy") != "0001")
                        {
                            contentDetail.LawNotCloseTypeName = "結案";
                        }
                        else
                        {
                            if (!String.IsNullOrEmpty(list[i].LawNotCloseTypeName))
                            {
                                contentDetail.LawNotCloseTypeName = "未結案";
                            }
                            else
                            {
                                contentDetail.LawNotCloseTypeName = "未結案";
                            }
                        }
                        break;
                    case 3: //無實益
                        if (list[i].LawContentCancelType == "1")
                        {
                            contentDetail.LawNotCloseTypeName = "結案";
                        }
                        else
                        {
                            if (!String.IsNullOrEmpty(list[i].LawNotCloseTypeName))
                            {
                                contentDetail.LawNotCloseTypeName = list[i].LawNotCloseTypeName;
                            }
                            else
                            {
                                contentDetail.LawNotCloseTypeName = "未結案";
                            }
                        }
                        break;
                }


                Viewmodel.Add(contentDetail);
            }
            var gridKey = _channelService.DataToCache(Viewmodel.AsEnumerable());
            SetGridKey("QueryGrid", gridKey);
        }

        public JsonResult BindGrid(jqGridParam jqParams)
        {
            var cacheKey = GetGridKey("QueryGrid");
            return BaseGridBinding<LawContentDetail>(jqParams,
                () => new WebChannel<ILAWService, LawContentDetail>().Get(cacheKey));
        }

        #region 個人資料
        /// <summary>
        /// 個人資料頁面
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpGet]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX009")]
        public ActionResult UpdatePersonalData(string AgentID, string NoteNo, string LawId)
        {
            LawAgentDetail agentDetails = _Service.GetLawAgentDetail(AgentID, NoteNo);
            List<LawContent> PreviousCase = _Service.GetLawContentByIDNo(AgentID.Substring(0, 10), NoteNo);
            if (PreviousCase.Count != 0)
            {
                for (int i = 0; i < PreviousCase.Count; i++)
                {
                    if (i == 0)
                    {
                        agentDetails.NoteStr = PreviousCase[i].LawNoteNo;
                    }
                    else
                    {
                        agentDetails.NoteStr = agentDetails.NoteStr + "," + PreviousCase[i].LawNoteNo;
                    }
                    if (PreviousCase[i].LawCloseDate.ToString("yyyy") == "0001")
                    {
                        if (!String.IsNullOrEmpty(PreviousCase[i].LawNotCloseTypeName))
                        {
                            agentDetails.StatusStr = "(" + PreviousCase[i].LawNotCloseTypeName + ")";
                        }
                        else
                        {
                            agentDetails.StatusStr = "";
                        }
                        agentDetails.NoteStr = agentDetails.NoteStr + agentDetails.StatusStr;
                    }
                    else
                    {
                        agentDetails.StatusStr = "(" + "結案日期：" + PreviousCase[i].LawCloseDate.ToString("yyyy/MM/dd") + ")";
                        agentDetails.NoteStr = agentDetails.NoteStr + agentDetails.StatusStr;
                    }
                }
            }
            else
            {
                agentDetails.NoteStr = "";
            }
            agentDetails.AgentCode = agentDetails.AgentCode.Substring(0, 10);
            agentDetails.AgentID = AgentID;
            return View(agentDetails);
        }

        /// <summary>
        /// 個人資料更新
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public ActionResult UpdatePersonalData(LawAgentDetail viewmodel)
        {
            viewmodel.Name = viewmodel.Name == null ? "" : viewmodel.Name;
            viewmodel.Cell01 = viewmodel.Cell01 == null ? "" : viewmodel.Cell01;
            viewmodel.Cell02 = viewmodel.Cell02 == null ? "" : viewmodel.Cell02;
            viewmodel.Phone01 = viewmodel.Phone01 == null ? "" : viewmodel.Phone01;
            viewmodel.Phone02 = viewmodel.Phone02 == null ? "" : viewmodel.Phone02;
            viewmodel.Address = viewmodel.Address == null ? "" : viewmodel.Address;
            viewmodel.Address01 = viewmodel.Address01 == null ? "" : viewmodel.Address01;
            viewmodel.Address02 = viewmodel.Address02 == null ? "" : viewmodel.Address02;
            viewmodel.AgentID = viewmodel.AgentID.Substring(0, 10);
            _Service.UpdateLawAgentDetail(viewmodel);
            return RedirectToAction("Index", "LAWTX009");
        }
        #endregion

        #region  進度表
        /// <summary>
        /// 進度表頁面
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpGet]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX009")]
        public ActionResult UpdateSchedule(string AgentID, string NoteNo, string LawId)
        {
            AgentID = AgentID.Substring(0, 10);
            LawContentDetail contentDetail = new LawContentDetail();
            LawContent content = _Service.GetSchedule(AgentID, NoteNo);
            contentDetail.AgentID = AgentID;
            contentDetail.LawContentCreateDate = content.LawContentCreateDate;
            contentDetail.LawContentCreateName = content.LawContentCreateName;
            contentDetail.LawId = content.LawId;
            contentDetail.LawDueReason = content.LawDueReason;
            contentDetail.LawNoteNo = content.LawNoteNo;
            contentDetail.LawFileNo = content.LawFileNo;
            contentDetail.LawPhoneCall1Desc = content.LawPhoneCall1Desc;
            contentDetail.LawPhoneCall1Date = content.LawPhoneCall1Date.ToString("yyyy/MM/dd") == "0001/01/01" ? "" : content.LawPhoneCall1Date.ToString("yyyy/MM/dd");
            contentDetail.OldLawPhoneCall1Desc = content.LawPhoneCall1Desc;
            contentDetail.OldLawPhoneCall1Date = content.LawPhoneCall1Date.ToString("yyyy/MM/dd") == "0001/01/01" ? "" : content.LawPhoneCall1Date.ToString("yyyy/MM/dd");
            contentDetail.LawPhoneCall2Desc = content.LawPhoneCall2Desc;
            contentDetail.LawPhoneCall2Date = content.LawPhoneCall2Date.ToString("yyyy/MM/dd") == "0001/01/01" ? "" : content.LawPhoneCall2Date.ToString("yyyy/MM/dd");
            contentDetail.OldLawPhoneCall2Desc = content.LawPhoneCall2Desc;
            contentDetail.OldLawPhoneCall2Date = content.LawPhoneCall2Date.ToString("yyyy/MM/dd") == "0001/01/01" ? "" : content.LawPhoneCall2Date.ToString("yyyy/MM/dd");
            contentDetail.LawDueAgentId = content.LawDueAgentId;
            contentDetail.LawStatusType = content.LawStatusType;
            contentDetail.LawNotCloseTypeName = content.LawNotCloseTypeName;
            contentDetail.LawCloseType = content.LawCloseType;

            //LawEvidenceDesc evidenceDesc = _Service.GetLawEvidenceDesc(LawId);
            //contentDetail.LawEvidencedesc = evidenceDesc.LawEvidencedesc;
            //contentDetail.LawEvidenceId = evidenceDesc.LawEvidenceId;

            LawDoUnitLog doUnitLog = _Service.GetLawDoUnitLog(LawId);
            if (doUnitLog != null)
            {
                contentDetail.LawDounitLogId = doUnitLog.LawDounitLogId;
                contentDetail.CaseDate = doUnitLog.CaseDate;
                contentDetail.UnitName = doUnitLog.UnitName;
                contentDetail.OldUnitName = doUnitLog.UnitName;
                contentDetail.OldCaseDate = doUnitLog.CaseDate;
            }

            List<LawDoUnitLog> doUnitLogList = _Service.GetLawDoUnitLogByLawId(LawId);

            for (int i = 0; i < doUnitLogList.Count; i++)
            {
                if (i == 0)
                {
                    contentDetail.dounitlist = doUnitLogList[i].UnitName;
                }
                else
                {
                    contentDetail.dounitlist = contentDetail.dounitlist + ">" + doUnitLogList[i].UnitName;
                }
            }

            LawCloseOtherLog closeOtherLog = _Service.GetLawCloseOtherLogByLawId(LawId);
            if (closeOtherLog != null)
            {
                contentDetail.LawCloseOtherDesc = closeOtherLog.LawCloseOtherDesc;
                contentDetail.LawCloseOtherId = closeOtherLog.LawCloseOtherId;
            }

            return View(contentDetail);
        }

        /// <summary>
        /// 更新進度表
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public ActionResult UpdateSchedule(LawContentDetail contentDetail)
        {
            LawContent content = new LawContent();
            content.LawDoUnitName = contentDetail.UnitName;
            content.LawDueReason = contentDetail.LawDueReason;
            content.LawFileNo = contentDetail.LawFileNo;
            content.LawId = contentDetail.LawId;
            _Service.UpdateLawContentReasonfile(content);
            LawDoUnit doUnit = _Service.GetLawDoUnitByUnitName(contentDetail.UnitName);
            content.LawDoUnitId = doUnit.LawDoUnitId;

            if(contentDetail.LawEvidencedesc != null)
            {
                LawEvidenceDesc lawEvidenceDesc = new LawEvidenceDesc();
                lawEvidenceDesc.LawEvidencedesc = contentDetail.LawEvidencedesc;
                lawEvidenceDesc.LawId = contentDetail.LawId;
                lawEvidenceDesc.LawNoteNo = contentDetail.LawNoteNo;
                lawEvidenceDesc.CreateDate = DateTime.Now;
                lawEvidenceDesc.EvidenceDescCreatorID = User.MemberInfo.ID;
                _Service.InsertLawEvidenceDesc(lawEvidenceDesc);
            }

            if(contentDetail.LawLitigationprogress != null)
            {
                LawLitigationProgress lawLitigationProgress = new LawLitigationProgress();
                lawLitigationProgress.LawLitigationprogress = contentDetail.LawLitigationprogress;
                lawLitigationProgress.LawId = contentDetail.LawId;
                lawLitigationProgress.LawNoteNo = contentDetail.LawNoteNo;
                lawLitigationProgress.CreateDate = DateTime.Now;
                lawLitigationProgress.LawLitigationProgressCreatorID = User.MemberInfo.ID;
                _Service.InsertLawLitigationProgress(lawLitigationProgress);
            }

            if (contentDetail.LawDoprogress != null)
            {
                LawDoProgress lawDoProgress = new LawDoProgress();
                lawDoProgress.LawDoprogress = contentDetail.LawDoprogress;
                lawDoProgress.LawNoteNo = contentDetail.LawNoteNo;
                lawDoProgress.LawId = contentDetail.LawId;
                lawDoProgress.CreateDate = DateTime.Now;
                lawDoProgress.LawDoProgressCreatorID = User.MemberInfo.ID;
                _Service.InsertLawDoProgress(lawDoProgress);
            }

            if(contentDetail.LawOtherdesc != null)
            {
                LawOtherDesc lawOtherDesc = new LawOtherDesc();
                lawOtherDesc.LawOtherdesc = contentDetail.LawOtherdesc;
                lawOtherDesc.LawNoteNo = contentDetail.LawNoteNo;
                lawOtherDesc.LawId = contentDetail.LawId;
                lawOtherDesc.LawDueAgentId = contentDetail.LawDueAgentId;
                lawOtherDesc.CreateDate = DateTime.Now;
                lawOtherDesc.OtherDescCreatorID = User.MemberInfo.ID;
                _Service.InsertLawOtherDesc(lawOtherDesc);
            }

            if(contentDetail.LawDescdesc != null)
            {
                LawDescDesc lawDescDesc = new LawDescDesc();
                lawDescDesc.LawDescdesc = contentDetail.LawDescdesc;
                lawDescDesc.LawNoteNo = contentDetail.LawNoteNo;
                lawDescDesc.LawId = contentDetail.LawId;
                lawDescDesc.CreateDate = DateTime.Now;
                lawDescDesc.LawDescDescCreatorID = User.MemberInfo.ID;
                _Service.InsertLawDescDesc(lawDescDesc);
            }

            if (contentDetail.OldUnitName == contentDetail.UnitName)
            {
                if (contentDetail.OldCaseDate != contentDetail.CaseDate)
                {
                    LawDoUnitLog doUnitLog = new LawDoUnitLog();
                    doUnitLog.CaseDate = contentDetail.CaseDate;
                    doUnitLog.LawDounitLogId = contentDetail.LawDounitLogId;
                    _Service.UpdateCaseDate(doUnitLog);
                }
            }
            else
            {
                LawDoUnitLog doUnitLog = new LawDoUnitLog();
                doUnitLog.CaseDate = contentDetail.CaseDate == null ? DateTime.Now.ToString("yyyy/MM/dd") : contentDetail.CaseDate;
                //doUnitLog.LawDounitLogId = contentDetail.LawDounitLogId;
                doUnitLog.LawId = contentDetail.LawId;
                doUnitLog.LawNoteNo = contentDetail.LawNoteNo;
                //doUnitLog.UnitName = contentDetail.UnitName;
                doUnitLog.CreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
                doUnitLog.CreateDate = DateTime.Now;
                doUnitLog.DounitLogCreatorID = User.MemberInfo.ID;
                _Service.InsertLawDoUnitLogUpdateLawContentByLawDoUnitId(doUnitLog, content);
            }

            if ((contentDetail.LawPhoneCall1Date == "" && contentDetail.LawPhoneCall1Desc == "") || (contentDetail.LawPhoneCall1Date == contentDetail.OldLawPhoneCall1Date && contentDetail.LawPhoneCall1Desc == contentDetail.OldLawPhoneCall1Desc))
            {

            }
            else
            {
                LawPhoneCallLog phoneCallLog = new LawPhoneCallLog();
                phoneCallLog.LawId = contentDetail.LawId;
                phoneCallLog.LawNoteNo = contentDetail.LawNoteNo;
                phoneCallLog.LawPhoneCallNo = "1";
                phoneCallLog.LawPhoneCallDate = contentDetail.LawPhoneCall1Date;
                phoneCallLog.LawPhoneCallLimitedDate = (Convert.ToDateTime(contentDetail.LawPhoneCall1Date).AddDays(7)).ToString("yyyy/MM/dd");
                phoneCallLog.LawPhoneCallReadLog = "0";
                phoneCallLog.PhoneCallCreatorID = User.MemberInfo.ID;
                content.LawPhoneCall1Desc = contentDetail.LawPhoneCall1Desc;
                content.LawPhoneCall1Date = contentDetail.LawPhoneCall1Date == null ? DateTime.Now : Convert.ToDateTime(contentDetail.LawPhoneCall1Date);
                _Service.InsertLawPhoneCallLogUpdateLawContentByLawId(phoneCallLog, content);
            }
            if ((contentDetail.LawPhoneCall2Date == "" && contentDetail.LawPhoneCall2Desc == "") || (contentDetail.LawPhoneCall2Date == contentDetail.OldLawPhoneCall2Date && contentDetail.LawPhoneCall2Desc == contentDetail.OldLawPhoneCall2Desc))
            {

            }
            else
            {
                LawPhoneCallLog phoneCallLog = new LawPhoneCallLog();
                phoneCallLog.LawId = contentDetail.LawId;
                phoneCallLog.LawNoteNo = contentDetail.LawNoteNo;
                phoneCallLog.LawPhoneCallNo = "2";
                phoneCallLog.LawPhoneCallDate = contentDetail.LawPhoneCall1Date;
                phoneCallLog.LawPhoneCallLimitedDate = (Convert.ToDateTime(contentDetail.LawPhoneCall1Date).AddDays(7)).ToString("yyyy/MM/dd");
                phoneCallLog.LawPhoneCallReadLog = "0";
                phoneCallLog.PhoneCallCreatorID = User.MemberInfo.ID;
                content.LawPhoneCall2Desc = contentDetail.LawPhoneCall2Desc;
                content.LawPhoneCall2Date = contentDetail.LawPhoneCall2Date == null ? DateTime.Now : Convert.ToDateTime(contentDetail.LawPhoneCall2Date);
                _Service.InsertLawPhoneCallLog2UpdateLawContentByLawId(phoneCallLog, content);
            }
            string CloseTypeName = Request.Form["CloseTypeName"];

            if (contentDetail.closetype == "0")
            {
                content.LawStatusType = 1;
                switch (contentDetail.nclose)
                {
                    case "1":
                        content.LawNotCloseTypeName = "進行中";
                        break;
                    case "2":
                        content.LawNotCloseTypeName = "分期";
                        break;
                    case "3":
                        content.LawNotCloseTypeName = "執行無實益";
                        break;
                }
                _Service.UpdateLawContentNotCloseType(content);
            }
            else
            {
                string[] sArray = CloseTypeName.Split(',');
                for (int i = 0; i < sArray.Length; i++)
                {
                    if (sArray[i] == "7") //其他
                    {
                        LawCloseOtherLog closeOtherLog = new LawCloseOtherLog();
                        closeOtherLog.LawId = contentDetail.LawId;
                        closeOtherLog.LawNoteNo = contentDetail.LawNoteNo;
                        closeOtherLog.LawCloseOtherDesc = contentDetail.LawCloseOtherDesc;
                        closeOtherLog.CreateDate = DateTime.Now;
                        closeOtherLog.CloseOtherLogCreatorID = User.MemberInfo.ID;
                        _Service.InsertLawCloseOtherLog(closeOtherLog);
                    }
                    if (sArray[i] == "8") //取消案件
                    {
                        content.LawContentCancelType = "1";
                        _Service.UpdateLawContentCancelType(content);
                    }
                }
                content.LawStatusType = 2;
                content.LawCloseType = CloseTypeName;
                content.LawCloseDate = contentDetail.LawCloseDate == null ? DateTime.Now : Convert.ToDateTime(contentDetail.LawCloseDate);
                content.LawContentCloserID = User.MemberInfo.ID;
                content.LawCloserName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
                content.LawContentLastchangeDate = DateTime.Now.ToString("yyyy/MM/dd");
                _Service.UpdateLawContentCloseType(content);
            }
            AppendMessage(EP.PlatformResources.更新成功, false);
            return RedirectToAction("UpdateSchedule", "LAWTX009", new { AgentID = contentDetail.AgentID, NoteNo = contentDetail.LawNoteNo, LawId = contentDetail.LawId });
        }

        ///// <summary>
        ///// 存證信函新增
        ///// </summary>
        ///// <param name="LawNote"></param>
        //[HttpPost]
        //public void InsertLawEvidencedesc(LawEvidenceDesc model)
        //{
        //    model.CreateDate = DateTime.Now;
        //    model.EvidenceDescCreatorID = User.MemberInfo.ID;
        //    _Service.InsertLawEvidenceDesc(model);
        //}

        ///// <summary>
        ///// 訴訟程序新增
        ///// </summary>
        ///// <param name="LawNote"></param>
        //[HttpPost]
        //public void InsertLawLitigationProgress(LawLitigationProgress model)
        //{
        //    model.CreateDate = DateTime.Now;
        //    model.LawLitigationProgressCreatorID = User.MemberInfo.ID;
        //    _Service.InsertLawLitigationProgress(model);
        //}

        ///// <summary>
        ///// 行政程序新增
        ///// </summary>
        ///// <param name="LawNote"></param>
        //[HttpPost]
        //public void InsertLawDoProgress(LawDoProgress model)
        //{
        //    model.CreateDate = DateTime.Now;
        //    model.LawDoProgressCreatorID = User.MemberInfo.ID;
        //    _Service.InsertLawDoProgress(model);
        //}

        ///// <summary>
        ///// 其他新增
        ///// </summary>
        ///// <param name="LawNote"></param>
        //[HttpPost]
        //public void InsertLawOtherDesc(LawOtherDesc model)
        //{
        //    model.CreateDate = DateTime.Now;
        //    model.OtherDescCreatorID = User.MemberInfo.ID;
        //    _Service.InsertLawOtherDesc(model);
        //}

        ///// <summary>
        ///// 備註新增
        ///// </summary>
        ///// <param name="LawNote"></param>
        //[HttpPost]
        //public void InsertLawDescDesc(LawDescDesc model)
        //{
        //    model.CreateDate = DateTime.Now;
        //    model.LawDescDescCreatorID = User.MemberInfo.ID;
        //    _Service.InsertLawDescDesc(model);
        //}

        /// <summary>
        /// 存證信函|訴訟程序|執行程序|其他的編輯頁面
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpGet]
        public ActionResult UpdateSharedPage(string ID, string LawId, string EditViewType, string AgentID)
        {
            LawEditModel lawEditModel = new LawEditModel();
            switch (EditViewType) //1:存證信函、2:訴訟程序、3:執行程序、4:其他、5:其他金額說明
            {
                case "1":
                    LawEvidenceDesc evidenceDesc = _Service.GetLawEvidenceDescById(Convert.ToInt32(ID));
                    lawEditModel.LawEvidencedesc = evidenceDesc.LawEvidencedesc;
                    lawEditModel.ID = evidenceDesc.LawEvidenceId;
                    lawEditModel.LawId = evidenceDesc.LawId;
                    lawEditModel.LawNoteNo = evidenceDesc.LawNoteNo;
                    break;
                case "2":
                    LawLitigationProgress litigation = _Service.GetLawLitigationProgressById(Convert.ToInt32(ID));
                    lawEditModel.LawLitigationprogress = litigation.LawLitigationprogress;
                    lawEditModel.ID = litigation.LawLitigationId;
                    lawEditModel.LawId = litigation.LawId;
                    lawEditModel.LawNoteNo = litigation.LawNoteNo;
                    break;
                case "3":
                    LawDoProgress doProgress = _Service.GetLawDoProgressById(Convert.ToInt32(ID));
                    lawEditModel.LawDoprogress = doProgress.LawDoprogress;
                    lawEditModel.ID = doProgress.LawDoId;
                    lawEditModel.LawId = doProgress.LawId;
                    lawEditModel.LawNoteNo = doProgress.LawNoteNo;
                    break;
                case "4":
                    LawOtherDesc otherDesc = _Service.GetLawOtherDescById(Convert.ToInt32(ID));
                    lawEditModel.LawOtherdesc = otherDesc.LawOtherdesc;
                    lawEditModel.LawId = otherDesc.LawId;
                    lawEditModel.LawNoteNo = otherDesc.LawNoteNo;
                    lawEditModel.LawRepaymentMoney = otherDesc.LawRepaymentMoney;
                    lawEditModel.ID = otherDesc.LawOtherId;
                    break;
                case "5":
                    LawOtherDesc otherMoneyDesc = _Service.GetLawOtherDescByLawId(AgentID, Convert.ToInt32(LawId), Convert.ToInt32(ID));
                    //lawEditModel.ID = otherMoneyDesc.LawOtherId;
                    LawRepaymentList repaymentList = _Service.GetLawRepaymentListById(Convert.ToInt32(ID));
                    lawEditModel.LawNoteNo = repaymentList.LawNoteNo;
                    lawEditModel.LawId = repaymentList.LawId;
                    lawEditModel.LawRepaymentMoney = repaymentList.LawRepaymentMoney;
                    lawEditModel.LawDueAgentId = repaymentList.LawDueAgentId;
                    lawEditModel.ID = repaymentList.LawRepaymentId;
                    break;
                case "6":
                    LawDescDesc descDesc = _Service.GetLawLawDescDescById(Convert.ToInt32(ID));
                    lawEditModel.LawNoteNo = descDesc.LawNoteNo;
                    lawEditModel.LawId = descDesc.LawId;
                    lawEditModel.LawDescdesc = descDesc.LawDescdesc;
                    lawEditModel.ID = descDesc.LawDescId;
                    break;
                case "7":
                    LawOtherDesc otherCommDeductionDesc = _Service.GetLawOtherDescByLawId(AgentID, Convert.ToInt32(LawId), Convert.ToInt32(ID));
                    //lawEditModel.ID = otherMoneyDesc.LawOtherId;
                    LawRepaymentList CommDeductionList = _Service.GetLawRepaymentListById(Convert.ToInt32(ID));
                    lawEditModel.LawNoteNo = CommDeductionList.LawNoteNo;
                    lawEditModel.LawId = CommDeductionList.LawId;
                    lawEditModel.LawCommDeduction = CommDeductionList.LawCommDeduction;
                    lawEditModel.LawDueAgentId = CommDeductionList.LawDueAgentId;
                    lawEditModel.ID = CommDeductionList.LawRepaymentId;
                    break;
            }
            lawEditModel.AgentID = AgentID;
            lawEditModel.EditViewType = EditViewType;
            lawEditModel.CreateTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            lawEditModel.EditUser = User.UnitInfo.Name + User.MemberInfo.Name;
            return View(lawEditModel);
        }

        /// <summary>
        /// 存證信函|訴訟程序|執行程序|其他的更新
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        public ActionResult UpdateSharedPage(LawEditModel model)
        {
            LawContent content = new LawContent();
            switch (model.EditViewType) //1:存證信函、2:訴訟程序、3:執行程序、4:其他、5:其他金額說明
            {
                case "1":
                    LawEvidenceDesc evidenceDesc = new LawEvidenceDesc();
                    evidenceDesc.LawEvidenceId = model.ID;
                    evidenceDesc.EvidenceDescCreatorID = User.MemberInfo.ID;
                    evidenceDesc.LawEvidencedesc = model.LawEvidencedesc;
                    _Service.UpdateLawEvidenceDesc(evidenceDesc);
                    break;
                case "2":
                    LawLitigationProgress litigationProgress = new LawLitigationProgress();
                    litigationProgress.LawLitigationId = model.ID;
                    litigationProgress.LawLitigationprogress = model.LawLitigationprogress;
                    litigationProgress.LawLitigationProgressCreatorID = User.MemberInfo.ID;
                    content.LawId = model.LawId;
                    content.LawContentLastchangeDate = DateTime.Now.ToString("yyyy/MM/dd");
                    _Service.UpdateLawLitigationProgress(litigationProgress, content);
                    break;
                case "3":
                    LawDoProgress doProgress = new LawDoProgress();
                    doProgress.LawDoId = model.ID;
                    doProgress.LawDoprogress = model.LawDoprogress;
                    doProgress.LawDoProgressCreatorID = User.MemberInfo.ID;
                    content.LawId = model.LawId;
                    content.LawContentLastchangeDate = DateTime.Now.ToString("yyyy/MM/dd");
                    _Service.UpdateLawDoProgress(doProgress, content);
                    break;
                case "4":
                    LawOtherDesc otherDesc = new LawOtherDesc();
                    otherDesc.LawOtherId = model.ID;
                    otherDesc.LawOtherdesc = model.LawOtherdesc;
                    otherDesc.OtherDescCreatorID = User.MemberInfo.ID;
                    _Service.UpdateLawOtherDesc(otherDesc);
                    break;
                case "5":
                    LawOtherDesc otherDescMoney = new LawOtherDesc();
                    otherDescMoney.LawRepaymentMoney = model.LawRepaymentMoney;
                    otherDescMoney.OtherDescCreatorID = User.MemberInfo.ID;
                    otherDescMoney.LawRepaymentId = model.ID;
                    otherDescMoney.LawDueAgentId = model.LawDueAgentId;
                    otherDescMoney.LawOtherdesc = model.LawOtherdesc;
                    otherDescMoney.LawNoteNo = model.LawNoteNo;
                    otherDescMoney.LawId = model.LawId;
                    otherDescMoney.CreateDate = DateTime.Now;
                    _Service.InsertLawOtherDescHaveMoney(otherDescMoney);
                    break;
                case "6":
                    LawDescDesc descDesc = new LawDescDesc();
                    descDesc.LawDescId = model.ID;
                    descDesc.LawDescdesc = model.LawDescdesc;
                    descDesc.LawDescDescCreatorID = User.MemberInfo.ID;
                    _Service.UpdateLawDescDesc(descDesc);
                    break;
                case "7":
                    LawOtherDesc otherDescCo = new LawOtherDesc();
                    otherDescCo.LawRepaymentMoney = model.LawCommDeduction;
                    otherDescCo.OtherDescCreatorID = User.MemberInfo.ID;
                    otherDescCo.LawRepaymentId = model.ID;
                    otherDescCo.LawDueAgentId = model.LawDueAgentId;
                    otherDescCo.LawOtherdesc = model.LawOtherdesc;
                    otherDescCo.LawNoteNo = model.LawNoteNo;
                    otherDescCo.LawId = model.LawId;
                    otherDescCo.CreateDate = DateTime.Now;
                    _Service.InsertLawOtherDescHaveMoney(otherDescCo);
                    break;
            }
            return RedirectToAction("UpdateSchedule", "LAWTX009", new { AgentID = model.AgentID, NoteNo = model.LawNoteNo, LawId = model.LawId });
        }

        /// <summary>
        /// 刪除一筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        //[HasPermission("EB.SL.PayRoll.PRTX002")]
        public void Delete(string ID, string EditViewType)
        {
            bool result = false;

            switch (EditViewType) //1:存證信函、2:訴訟程序、3:執行程序、4:其他、5:其他金額說明
            {
                case "1":
                    result = _Service.DeleteLawEvidenceDescByID(ID);
                    break;
                case "2":
                    result = _Service.DeleteLawLitigationProgressByID(ID);
                    break;
                case "3":
                    result = _Service.DeleteLawDoProgressByID(ID);
                    break;
                case "4":
                    result = _Service.DeleteLawOtherDescByID(ID);
                    break;
                case "5":
                    result = _Service.DeleteLawOtherDescByID(ID);
                    break;
                case "6":
                    result = _Service.DeleteLawDescDescByID(ID);
                    break;
                case "7":
                    result = _Service.DeleteLawOtherDescByID(ID);
                    break;
            }
            if (result)
                AppendMessage(PlatformResources.刪除成功, false);
            else
                AppendMessage(PlatformResources.刪除失敗, false);
        }
        #endregion

        #region 結欠金額
        #region 結欠明細
        /// <summary>
        ///  結欠頁面 
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult UpdateDueMoneyList(string AgentID, string NoteNo, string LawId)
        {
            LawContentDetail viewmodel = new LawContentDetail();
            viewmodel.AgentID = AgentID;
            viewmodel.LawNoteNo = NoteNo;
            viewmodel.LawId = Convert.ToInt32(LawId);
            return View(viewmodel);
        }

        /// <summary>
        ///  結欠明細 
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpGet]
        public JsonResult UpdateBalanceDetail(string AgentID, string NoteNo)
        {
            List<LawContentDetail> QResultList = new List<LawContentDetail>();
            List<LawContent> contents = _Service.GetDueMoneyDetail(AgentID.Substring(0, 10), NoteNo);
            for (int i = 0; i < contents.Count; i++)
            {
                LawContentDetail lawContentDetail = new LawContentDetail();
                lawContentDetail.LawNoteNo = contents[i].LawNoteNo;
                lawContentDetail.LawProductionYM = contents[i].LawYear + "/" + contents[i].LawMonth + "-" + contents[i].LawPaySequence + "薪";
                lawContentDetail.LawDueMoney = contents[i].LawDueMoney;
                lawContentDetail.LawInterestRatesId = contents[i].LawInterestRatesId;
                lawContentDetail.LawInterestSdate = contents[i].LawInterestSdate == null ? "" : contents[i].LawInterestSdate;
                lawContentDetail.LawInterestEdate = contents[i].LawInterestEdate == null ? "" : contents[i].LawInterestEdate;
                lawContentDetail.LawInterestDays = contents[i].LawInterestDays;
                lawContentDetail.LawInterestMoney = contents[i].LawInterestMoney;
                lawContentDetail.LawTotalMoney = contents[i].LawTotalMoney;
                lawContentDetail.LawId = contents[i].LawId;
                lawContentDetail.LawDueName = contents[i].LawDueName;
                QResultList.Add(lawContentDetail);
            }

            return Json(QResultList, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        ///  結欠明細更新
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UpdateBalanceDetail(LawContentDetail contentDetail)
        {
            LawContent content = new LawContent();
            int inte_money, tot_money;
            int inte_day;
            if (contentDetail.LawInterestRatesId == 10) //不須計算利息
            {
                List<LawInterestRates> lawInterestRates = new List<LawInterestRates>();
                lawInterestRates = _Service.GetLawInterestRatesByID(contentDetail.LawInterestRatesId.ToString());
                inte_day = 0;
                if (lawInterestRates.Count != 0)
                {
                    inte_money = Convert.ToInt32((contentDetail.LawDueMoney * contentDetail.InterestRates / 100) * inte_day / 365);
                }
                else
                {
                    inte_money = Convert.ToInt32(contentDetail.LawDueMoney);
                }
                tot_money = Convert.ToInt32(contentDetail.LawDueMoney + inte_money);
                content = _Service.GetLawContentByLawId(contentDetail.LawId);
                if (content != null && (!String.IsNullOrEmpty(content.LawInterestSdate)
                        || (contentDetail.LawInterestRatesId != contentDetail.OldLawInterestRatesId)
                        || (contentDetail.LawInterestSdate != contentDetail.OldLawInterestSdate)
                        || (contentDetail.LawInterestEdate != contentDetail.OldLawInterestEdate)
                        || (contentDetail.LawDueMoney != contentDetail.OldLawDueMoney)))
                {
                    //if (!String.IsNullOrEmpty(content.LawInterestSdate)
                    //    || (contentDetail.LawInterestRatesId != contentDetail.OldLawInterestRatesId)
                    //    || (contentDetail.LawInterestSdate != contentDetail.OldLawInterestSdate)
                    //    || (contentDetail.LawInterestEdate != contentDetail.OldLawInterestEdate))
                    //{
                    LawInterestLog interestLog = new LawInterestLog();
                    interestLog.LawId = contentDetail.LawId;
                    interestLog.LawNoteNo = contentDetail.LawNoteNo;
                    interestLog.LawDueMoney = contentDetail.LawDueMoney;
                    interestLog.LawInterestSdate = contentDetail.LawInterestSdate;
                    interestLog.LawInterestEdate = contentDetail.LawInterestEdate;
                    interestLog.LawInterestDays = contentDetail.LawInterestDays;
                    interestLog.LawInterestRatesId = contentDetail.LawInterestRatesId;
                    interestLog.LawInterestMoney = inte_money;
                    interestLog.LawTotalMoney = tot_money;
                    interestLog.CreateDate = DateTime.Now;
                    interestLog.InterestLogCreateID = User.MemberInfo.ID;
                    _Service.InsertLawInterestLog(interestLog);

                    LawContent lawContent = new LawContent();
                    lawContent.LawDueMoney = contentDetail.LawDueMoney;
                    lawContent.LawInterestSdate = contentDetail.LawInterestSdate;
                    lawContent.LawInterestEdate = contentDetail.LawInterestEdate;
                    lawContent.LawInterestRatesId = contentDetail.LawInterestRatesId;
                    lawContent.LawInterestMoney = inte_money;
                    lawContent.LawInterestDays = inte_day;
                    lawContent.LawTotalMoney = tot_money;
                    lawContent.LawId = contentDetail.LawId;
                    _Service.UpdateLawContentDueMoney(lawContent);
                    //}
                }
            }
            else
            {
                List<LawInterestRates> lawInterestRates = new List<LawInterestRates>();
                lawInterestRates = _Service.GetLawInterestRatesByID(contentDetail.LawInterestRatesId.ToString());
                inte_day = (Convert.ToDateTime(contentDetail.LawInterestEdate) - Convert.ToDateTime(contentDetail.LawInterestSdate)).Days;
                if (lawInterestRates.Count != 0)
                {
                    inte_money = Convert.ToInt32((contentDetail.LawDueMoney * contentDetail.InterestRates / 100) * inte_day / 365);
                }
                else
                {
                    inte_money = Convert.ToInt32(contentDetail.LawDueMoney);
                }
                tot_money = Convert.ToInt32(contentDetail.LawDueMoney + inte_money);
                content = _Service.GetLawContentByLawId(contentDetail.LawId);
                if (content != null)
                {
                    if (content != null && (!String.IsNullOrEmpty(content.LawInterestSdate)
                            || (contentDetail.LawInterestRatesId != contentDetail.OldLawInterestRatesId)
                            || (contentDetail.LawInterestSdate != contentDetail.OldLawInterestSdate)
                            || (contentDetail.LawInterestEdate != contentDetail.OldLawInterestEdate)
                            || (contentDetail.LawDueMoney != contentDetail.OldLawDueMoney)))
                    {
                        //if (!String.IsNullOrEmpty(content.LawInterestSdate)
                        //    || (contentDetail.LawInterestRatesId != contentDetail.OldLawInterestRatesId)
                        //    || (contentDetail.LawInterestSdate != contentDetail.OldLawInterestSdate)
                        //    || (contentDetail.LawInterestEdate != contentDetail.OldLawInterestEdate))
                        //{
                        LawInterestLog interestLog = new LawInterestLog();
                        interestLog.LawId = contentDetail.LawId;
                        interestLog.LawNoteNo = contentDetail.LawNoteNo;
                        interestLog.LawDueMoney = contentDetail.LawDueMoney;
                        interestLog.LawInterestSdate = contentDetail.LawInterestSdate;
                        interestLog.LawInterestEdate = contentDetail.LawInterestEdate;
                        interestLog.LawInterestDays = contentDetail.LawInterestDays;
                        interestLog.LawInterestRatesId = contentDetail.LawInterestRatesId;
                        interestLog.LawInterestMoney = inte_money;
                        interestLog.LawTotalMoney = tot_money;
                        interestLog.CreateDate = DateTime.Now;
                        interestLog.InterestLogCreateID = User.MemberInfo.ID;
                        _Service.InsertLawInterestLog(interestLog);

                        LawContent lawContent = new LawContent();
                        lawContent.LawDueMoney = contentDetail.LawDueMoney;
                        lawContent.LawInterestSdate = contentDetail.LawInterestSdate;
                        lawContent.LawInterestEdate = contentDetail.LawInterestEdate;
                        lawContent.LawInterestRatesId = contentDetail.LawInterestRatesId;
                        lawContent.LawInterestMoney = inte_money;
                        lawContent.LawInterestDays = inte_day;
                        lawContent.LawTotalMoney = tot_money;
                        lawContent.LawId = contentDetail.LawId;
                        _Service.UpdateLawContentDueMoney(lawContent);
                        //}
                    }
                }
            }
            return RedirectToAction("UpdateDueMoneyList", "LAWTX009", new { AgentID = contentDetail.AgentID, NoteNo = contentDetail.LawNoteNo, LawId = contentDetail.LawId });
        }

        /// <summary>
        ///  結欠明細 
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        public int ChkIRInsert(string LawDueAgentId, string LawID)
        {
            LawContent lawContent = _Service.ChkIRInsert(Convert.ToInt32(LawID), LawDueAgentId);

            return lawContent.LawTotalMoney;
        }
        #endregion

        #region 清償明細||律師服務費
        /// <summary>
        /// 清償明細||律師服務費查詢
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX009")]
        public void RepayMentQuery(LawContentDetail model)
        {
            //取資料
            List<LawRepaymentList> list = new List<LawRepaymentList>();
            List<LawLawyerReward> LRlist = new List<LawLawyerReward>();
            List<LawContentDetail> Viewmodel = new List<LawContentDetail>();
            WebChannel<ILAWService> _channelService = new WebChannel<ILAWService>();
            string repay_idstr = string.Empty;
            string RemainLawRepaymentMoney = string.Empty;
            List<int> ls = new List<int>();
            if (model.PayMentType == "1")
            {
                _channelService.Use(service => list = service.GetLawRepaymentListByLawId(Convert.ToInt32(model.LawId)));
                //列表
                for (int i = 0; i < list.Count; i++)
                {
                    LawContentDetail contentDetail = new LawContentDetail();
                    contentDetail.LawId = list[i].LawId;
                    contentDetail.LawRepaymentId = list[i].LawRepaymentId;
                    contentDetail.LawNoteNo = list[i].LawNoteNo;
                    contentDetail.LawRepaymentDate = list[i].LawRepaymentDate.ToString("yyyy/MM/dd") == "0001/01/01" ? "" : list[i].LawRepaymentDate.ToString("yyyy/MM/dd");
                    contentDetail.LawRepaymentMoney = list[i].LawRepaymentMoney;
                    contentDetail.LawRepaymentCapital = list[i].LawRepaymentCapital;
                    contentDetail.LawRepaymentInterest = list[i].LawRepaymentInterest;
                    contentDetail.LawRepaymentCourt = list[i].LawRepaymentCourt;
                    contentDetail.LawRepaymentOther = list[i].LawRepaymentOther;
                    contentDetail.LawCommDeduction = list[i].LawCommDeduction;

                    ls.Add(list[i].LawRepaymentId);
                    LawRepaymentList lre = _Service.GetTotalLawRepaymentMoney(list[i].LawDueAgentId, ls);
                    contentDetail.TotalLawRepaymentMoney = lre.LawRepaymentMoney;
                    LawContent ct = _Service.GetRemainLawRepaymentMoney(list[i].LawDueAgentId, list[i].LawId);
                    RemainLawRepaymentMoney = ct.LawTotalMoney.ToString();
                    Viewmodel.Add(contentDetail);
                }
                if (RemainLawRepaymentMoney != "0" && !String.IsNullOrEmpty(RemainLawRepaymentMoney))
                {
                    Viewmodel[Viewmodel.Count - 1].RemainLawRepaymentMoney = (Convert.ToInt32(RemainLawRepaymentMoney) - Viewmodel[Viewmodel.Count - 1].TotalLawRepaymentMoney).ToString();
                }
            }
            else
            {
                _channelService.Use(service => LRlist = service.GetLawLawyerReward(Convert.ToInt32(model.LawId), model.AgentID));
                //列表
                for (int i = 0; i < LRlist.Count; i++)
                {
                    LawContentDetail contentDetail = new LawContentDetail();
                    contentDetail.LawLawyerPayId = LRlist[i].LawLawyerPayId;
                    contentDetail.LawId = LRlist[i].LawId;
                    contentDetail.LawNoteNo = LRlist[i].LawNoteNo;
                    contentDetail.LawDueAgentId = LRlist[i].LawDueAgentId;
                    contentDetail.LawRepaymentMoneyORG = LRlist[i].LawRepaymentMoney;
                    contentDetail.LawFees = LRlist[i].LawFees;
                    contentDetail.LawRewardPayYearMonth = LRlist[i].LawRewardPayYear + "/" + LRlist[i].LawRewardPayMonth;
                    List<LawLawyerServiceRates> serviceRates = _Service.GetLawLawyerServiceRatesByID(LRlist[i].LawRates.ToString());
                    for (int j = 0; j < serviceRates.Count; j++)
                    {
                        contentDetail.LawyerServiceRates = serviceRates[j].LawyerServiceRates.ToString();
                    }
                    contentDetail.LawServiceReward = LRlist[i].LawServiceReward;

                    Viewmodel.Add(contentDetail);
                }
            }

            var gridKey = _channelService.DataToCache(Viewmodel.AsEnumerable());
            SetGridKey("RepayMentQueryGrid", gridKey);
        }

        public JsonResult RepayMentBindGrid(jqGridParam jqParams)
        {
            var cacheKey = GetGridKey("RepayMentQueryGrid");
            return BaseGridBinding<LawContentDetail>(jqParams,
                () => new WebChannel<ILAWService, LawContentDetail>().Get(cacheKey));
        }

        /// <summary>
        /// 新增清償明細||律師服務費
        /// </summary>
        /// <returns></returns>
        public ActionResult Create(int LawId, string PayMentType)
        {

            LawContentDetail contentDetail = new LawContentDetail();
            LawContent content = _Service.GetLawContentByLawId(LawId);
            contentDetail.LawId = content.LawId;
            contentDetail.LawNoteNo = content.LawNoteNo;
            contentDetail.AgentID = content.LawDueAgentId;
            contentDetail.LawDueAgentId = content.LawDueAgentId;
            contentDetail.PayMentType = PayMentType;
            return View(contentDetail);
        }

        /// <summary>
        /// 寫入清償明細||律師服務費
        /// </summary>
        /// <param name="model">系統設定view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Create(LawContentDetail model)
        {
            if (model.PayMentType == "1") //清償明細
            {
                LawRepaymentList repaymentList = new LawRepaymentList();
                repaymentList.CreateDate = DateTime.Now;
                repaymentList.RepaymentListCreatorID = User.MemberInfo.ID;
                repaymentList.LawDueAgentId = model.AgentID;
                repaymentList.LawRepaymentDate = Convert.ToDateTime(model.LawRepaymentDate);
                repaymentList.LawRepaymentMoney = model.LawRepaymentMoney;
                repaymentList.LawRepaymentCapital = model.LawRepaymentCapital;
                repaymentList.LawRepaymentInterest = model.LawRepaymentInterest;
                repaymentList.LawRepaymentCourt = model.LawRepaymentCourt;
                repaymentList.LawRepaymentOther = model.LawRepaymentOther;
                repaymentList.LawCommDeduction = model.LawCommDeduction;
                repaymentList.LawId = model.LawId;
                repaymentList.LawNoteNo = model.LawNoteNo;
                _Service.InsertLawRepaymentList(repaymentList);
                return RedirectToAction("UpdateDueMoneyList", "LAWTX009", new { AgentID = model.AgentID, NoteNo = model.LawNoteNo, LawId = model.LawId });
            }
            else
            {
                decimal rates_str = 0, reward_tot = 0;
                LawLawyerReward lawyerReward = new LawLawyerReward();
                lawyerReward.LawId = model.LawId;
                lawyerReward.LawDueAgentId = model.LawDueAgentId;
                lawyerReward.LawNoteNo = model.LawNoteNo;
                lawyerReward.LawRewardPayYear = model.LawRewardPayYear;
                lawyerReward.LawRewardPayMonth = model.LawRewardPayMonth;
                lawyerReward.LawRepaymentMoney = model.LawRepaymentMoney;
                lawyerReward.LawFees = model.LawFees;
                lawyerReward.LawRates = model.LawRates;
                List<LawLawyerServiceRates> serviceRates = _Service.GetLawLawyerServiceRatesByID(model.LawRates.ToString());
                for (int j = 0; j < serviceRates.Count; j++)
                {
                    rates_str = serviceRates[j].LawyerServiceRates / 100;
                }
                reward_tot = (model.LawRepaymentMoney - model.LawFees) * rates_str;
                lawyerReward.LawServiceReward = Convert.ToInt32(reward_tot);
                lawyerReward.CreateDate = DateTime.Now;
                lawyerReward.LawyerRewardCreateID = User.MemberInfo.ID;
                _Service.InsertLawLawyerReward(lawyerReward);
                return RedirectToAction("UpdateDueMoneyList", "LAWTX009", new { AgentID = model.LawDueAgentId, NoteNo = model.LawNoteNo, LawId = model.LawId });
            }
        }

        /// <summary>
        /// 更新律師服務費
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateLawyerReward(int LawLawyerPayId, int LawId, string LawDueAgentId)
        {
            LawLawyerReward content = _Service.GetLawLawyerRewardById(LawLawyerPayId);
            return View(content);
        }

        [HttpPost]
        /// <summary>
        /// 更新律師服務費
        /// </summary>
        /// <returns></returns>
        public ActionResult UpdateLawyerReward(LawLawyerReward model)
        {
            decimal rates_str = 0, reward_tot = 0;
            List<LawLawyerServiceRates> serviceRates = _Service.GetLawLawyerServiceRatesByID(model.LawRates.ToString());
            for (int j = 0; j < serviceRates.Count; j++)
            {
                rates_str = serviceRates[j].LawyerServiceRates / 100;
            }
            reward_tot = (model.LawRepaymentMoney - model.LawFees) * rates_str;
            model.LawServiceReward = Convert.ToInt32(reward_tot);
            _Service.UpdateLawLawyerReward(model);
            return RedirectToAction("UpdateDueMoneyList", "LAWTX009", new { AgentID = model.LawDueAgentId, NoteNo = model.LawNoteNo, LawId = model.LawId });
        }

        /// <summary>
        /// 刪除一筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        [HttpPost]
        public void DeleteRepayment(string ID)
        {
            LawRepaymentList repaymentList = _Service.GetLawRepaymentListById(Convert.ToInt32(ID));
            if (repaymentList != null)
            {
                LawRepaymentListLog repaymentListLog = new LawRepaymentListLog();
                repaymentListLog.CreateDate = DateTime.Now;
                repaymentListLog.DelName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
                repaymentListLog.DelTime = DateTime.Now;
                repaymentListLog.LawCommDeduction = repaymentList.LawCommDeduction;
                repaymentListLog.LawCommDeductionDate = repaymentList.LawCommDeductionDate;
                repaymentListLog.LawDueAgentId = repaymentList.LawDueAgentId;
                repaymentListLog.LawId = repaymentList.LawId;
                repaymentListLog.LawNoteNo = repaymentList.LawNoteNo;
                repaymentListLog.LawRepaymentCapital = repaymentList.LawRepaymentCapital;
                repaymentListLog.LawRepaymentCourt = repaymentList.LawRepaymentCourt;
                repaymentListLog.LawRepaymentDate = repaymentList.LawRepaymentDate;
                repaymentListLog.LawRepaymentId = repaymentList.LawRepaymentId;
                repaymentListLog.LawRepaymentInterest = repaymentList.LawRepaymentInterest;
                repaymentListLog.LawRepaymentMoney = repaymentList.LawRepaymentMoney;
                repaymentListLog.LawRepaymentOther = repaymentList.LawRepaymentOther;
                repaymentListLog.RepaymentListLogCreatorID = User.MemberInfo.ID;
                repaymentListLog.RepaymentListLogDelID = User.MemberInfo.ID;
                _Service.InsertLawRepaymentListLog(repaymentListLog);
            }
            bool result = _Service.DeleteLawRepaymentListByID(ID);

            if (result)
                AppendMessage(PlatformResources.刪除成功, false);
            else
                AppendMessage(PlatformResources.刪除失敗, false);
        }
        #endregion
        #endregion

        #region 存證信函
        ///// <summary>
        ///// 存證信函頁面
        ///// </summary>
        ///// <param name="LawNote"></param>
        //[HttpGet]
        //[HasPermission("EP.SD.SalesSupport.LAW.LAWTX009")]
        //public ActionResult UpdateLawEvidence(string AgentID, string NoteNo, string LawId)
        //{
        //    LawEvidenceDetail evidenceDetail = _Service.GetLawEvidenceById(AgentID.Substring(0, 10), NoteNo);
        //    if (evidenceDetail == null)
        //    {
        //        LawEvidenceDetail lawEvidenceDetail = new LawEvidenceDetail();
        //        lawEvidenceDetail.MyPageStatus = "Create";
        //        lawEvidenceDetail = _Service.GetLawAgentDetailById(AgentID, NoteNo);
        //        lawEvidenceDetail.EvidSender = "永達保險經紀人(股)公司";
        //        lawEvidenceDetail.EvidSenderAdd = "台北市中山北路二段79號5樓";
        //        lawEvidenceDetail.EvidAgentName = lawEvidenceDetail.Name;
        //        lawEvidenceDetail.EvidAgentId = lawEvidenceDetail.AgentCode;
        //        lawEvidenceDetail.EvidAgentAdd = lawEvidenceDetail.Address;
        //        lawEvidenceDetail.EvidAgentAdd1 = lawEvidenceDetail.Address01;
        //        lawEvidenceDetail.EvidAgentAdd2 = lawEvidenceDetail.Address02;
        //        lawEvidenceDetail.EvidAccount = lawEvidenceDetail.SuperAccount;
        //        lawEvidenceDetail.EvidAgentId = AgentID;
        //        lawEvidenceDetail.LawNoteNo = NoteNo;
        //        lawEvidenceDetail.LawId = Convert.ToInt32(LawId);
        //        return View(lawEvidenceDetail);
        //    }
        //    else
        //    {
        //        evidenceDetail.MyPageStatus = "Edit";
        //        evidenceDetail.EvidAgentId = AgentID;
        //        evidenceDetail.LawNoteNo = NoteNo;
        //        evidenceDetail.LawId = Convert.ToInt32(LawId);
        //        return View(evidenceDetail);
        //    }
        //}

        ///// <summary>
        ///// 更新存證信函
        ///// </summary>
        ///// <param name="LawNote"></param>
        //[HttpPost]
        //public ActionResult UpdateLawEvidence(LawEvidenceDetail model)
        //{
        //    Int64[] numbers ={ Convert.ToInt64(model.EvidMoneyNum) };
        //    foreach (var number in numbers)
        //    {
        //        var helper = new Number2Text(number);
        //        string result = helper.GetText();
        //        model.EvidMoney = result;
        //    }
        //    string chkadd = Request.Form["chkadd"];
        //    if(chkadd == "2")
        //    {
        //        model.EvidAgentAdd1 = model.EvidAgentAdd2;
        //    }
        //    if (!String.IsNullOrEmpty(model.reason))
        //    {
        //        model.EvidReason = model.reason;
        //    }
        //    model.EvidAgentId = model.EvidAgentId.Substring(0, 10);
        //    if (model.MyPageStatus == "Edit")
        //    {
        //        LawEvidenceLog evidenceLog = new LawEvidenceLog();
        //        evidenceLog.EvidId = model.EvidId;
        //        evidenceLog.EvidNo = model.EvidNo == null ? "" : model.EvidNo;
        //        evidenceLog.EvidSender = model.EvidSender;
        //        evidenceLog.EvidSenderAdd = model.EvidSenderAdd;
        //        evidenceLog.EvidAgentId = model.EvidAgentId;
        //        evidenceLog.EvidAgentName = model.EvidAgentName;
        //        evidenceLog.EvidAgentAdd = model.EvidAgentAdd;
        //        evidenceLog.EvidAgentAdd1 = model.EvidAgentAdd1;
        //        evidenceLog.EvidReason = model.EvidReason;
        //        evidenceLog.EvidMoney = model.EvidMoney;
        //        evidenceLog.EvidMoneyNum = model.EvidMoneyNum;
        //        evidenceLog.EvidAccount = model.EvidAccount;
        //        evidenceLog.EvidUser = model.EvidUser;
        //        evidenceLog.EvidPhone = model.EvidPhone;
        //        evidenceLog.EvidenceLogCreatorID = User.MemberInfo.ID;
        //        evidenceLog.EvidCreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
        //        evidenceLog.EvidCreateDate = DateTime.Now;
        //        evidenceLog.LawId = model.LawId;
        //        evidenceLog.LawNoteNo = model.LawNoteNo;
        //        _Service.InsertLawEvidenceLog(evidenceLog);

        //        LawEvidence evidence = new LawEvidence();
        //        evidence.EvidId = model.EvidId;
        //        evidence.EvidNo = model.EvidNo == null ? "" : model.EvidNo;
        //        evidence.EvidSender = model.EvidSender;
        //        evidence.EvidSenderAdd = model.EvidSenderAdd;
        //        evidence.EvidAgentId = model.EvidAgentId;
        //        evidence.EvidAgentName = model.EvidAgentName;
        //        evidence.EvidAgentAdd = model.EvidAgentAdd;
        //        evidence.EvidAgentAdd1 = model.EvidAgentAdd1;
        //        evidence.EvidReason = model.EvidReason;
        //        evidence.EvidMoney = model.EvidMoney;
        //        evidence.EvidMoneyNum = model.EvidMoneyNum;
        //        evidence.EvidAccount = model.EvidAccount;
        //        evidence.EvidUser = model.EvidUser;
        //        evidence.EvidPhone = model.EvidPhone;
        //        evidence.EvidenceCreatorID = User.MemberInfo.ID;
        //        evidence.EvidCreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
        //        evidence.EvidCreateDate = DateTime.Now;
        //        evidence.LawId = model.LawId;
        //        evidence.LawNoteNo = model.LawNoteNo;
        //        _Service.UpdateLawEvidence(evidence);
        //    }
        //    else
        //    {
        //        LawEvidence evidence = new LawEvidence();
        //        evidence.EvidNo = model.EvidNo == null ? "" : model.EvidNo;
        //        evidence.EvidSender = model.EvidSender;
        //        evidence.EvidSenderAdd = model.EvidSenderAdd;
        //        evidence.EvidAgentId = model.EvidAgentId;
        //        evidence.EvidAgentName = model.EvidAgentName;
        //        evidence.EvidAgentAdd = model.EvidAgentAdd;
        //        evidence.EvidAgentAdd1 = model.EvidAgentAdd1;
        //        evidence.EvidReason = model.EvidReason;
        //        evidence.EvidMoney = model.EvidMoney;
        //        evidence.EvidMoneyNum = model.EvidMoneyNum;
        //        evidence.EvidAccount = model.EvidAccount;
        //        evidence.EvidUser = model.EvidUser;
        //        evidence.EvidPhone = model.EvidPhone;
        //        evidence.EvidenceCreatorID = User.MemberInfo.ID;
        //        evidence.EvidCreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
        //        evidence.EvidCreateDate = DateTime.Now;
        //        evidence.LawId = model.LawId;
        //        evidence.LawNoteNo = model.LawNoteNo;
        //        _Service.InsertLawEvidence(evidence);
        //    }
        //    return RedirectToAction("Index", "LAWTX009");
        //}
        #endregion

        #region 照會單
        /// <summary>
        /// 照會單頁面
        /// </summary>
        /// <param name="LawNote"></param>
        [HttpGet]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX009")]
        public ActionResult UpdateNoteNoList(string AgentID, string NoteNo, string LawId)
        {
            LawContent content = _Service.GetLawContentByLawId(Convert.ToInt32(LawId));
            LawNoteDetail noteDetail = new LawNoteDetail();
            noteDetail.LawNoteNo = NoteNo;
            noteDetail.LawNoteName = content.LawDueName;
            noteDetail.AgentID = AgentID;
            noteDetail.LawId = Convert.ToInt32(LawId);
            return View(noteDetail);
        }

        /// <summary>
        /// 寫入照會單
        /// </summary>
        /// <param name="model">系統設定view model</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult UpdateNoteNoList(LawNoteDetail model)
        {
            string retMsg = string.Empty, MSID;
            var MsgService = ServiceHelper.Create<IMessageService>();
            List<AccountGroupJsonString> JsonClassLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(model.RecipientToJson);
            List<UserSimpleInfo> idnameLst = PlatformHelper.GetRecipientUsersList(JsonClassLst);

            //新增留言
            Message k = new Message();
            k.MSGSubject = DateTime.Now.ToString("yyyy/MM/dd") + "【" + model.LawNoteNo + "】：法追照會單通知作業";
            k.MSGDESC = model.LawNoteNo + "照會單通知，請至法追業務系統通知作業下載照會單，以利法追作業進行。";
            k.MSGCreater = User.MemberInfo.ID;// (int)Session["orgID"];
            k.MSGCreateName = Microsoft.CUF.CodeName.GetUnitName(User.UnitInfo.ID).Trim() + "-" + Microsoft.CUF.CodeName.GetMemberName(User.MemberInfo.ID);
            k.MSGCreateIP = PlatformHelper.GetClientIPv4();
            k.MSGTime = DateTime.Now;
            k.MSGClass = 3;
            k.MSGNote = model.LawNoteNo;
            //產生留言人員名單
            List<MessageTo> kt_list = null;
            List<MessageFile> kf_list = new List<MessageFile>();
            MessageTo partdata = new MessageTo();
            partdata.MSGOBJDate = k.MSGTime;
            partdata.MSGOBJReaderIP = "";
            partdata.MSGOBJSendIP = k.MSGCreateIP;
            partdata.MSGOBJCreateTime = k.MSGTime;
            partdata.MSGOBJSendID = User.MemberInfo.ID;
            partdata.MSGOBJNote = model.LawNoteNo;
            kt_list = GetMsgIdsToMsgTList(idnameLst, partdata);
            MSID = MsgService.CreateMessage(k, kf_list, kt_list, TabUniqueId, out retMsg);

            //LawNote note = new LawNote();
            //note.LawNoteNo = model.LawNoteNo;
            //note.LawNoteKm = Convert.ToInt32(MSID);
            //note.LawNoteCenter = model.LawNoteCenter;
            //note.LawNoteTo = model.LawNoteTo;
            //note.LawNoteLevel = model.LawNoteLevel;
            //note.LawNoteName = model.LawNoteName;
            //note.LawNoteDep = LAWHelper.ChangeUnitName(User.MemberInfo.ID);
            //note.LawNotePro = model.LawNotePro;
            //note.LawNoteTel = model.LawNoteTel;
            //note.LawNoteType = 1;
            //note.LawNoteCreatername = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
            //note.LawNoteCreatedate =DateTime.Now;
            //note.NoteCreatorID =User.MemberInfo.ID;
            //_Service.InsertLawNote(note);

            return View(model);
        }

        /// <summary>
        /// 照會單作業-已發送紀錄
        /// </summary>
        /// <param name="model">單筆調整view model</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult NoteQuery(string LawNoteNo)
        {
            List<MessageTo> QResultList = new List<MessageTo>();
            var mService = new WebChannel<ILAWService>();

            mService.Use(service => service
            .GetMessageToByLawNoteNo(LawNoteNo)
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new MessageTo();
                    item.MSGOBJCreateTime = d.MSGOBJCreateTime;
                    item.MSGOBJName = d.MSGOBJName;
                    item.MSGOBJReaderIP = d.MSGOBJReaderIP == null ? "" : d.MSGOBJReaderIP;
                    QResultList.Add(item);
                }
            }));

            return Json(QResultList, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 組成留言接收者清單 
        /// </summary>
        /// <param name="idNamelist">接收者資料</param>
        /// <param name="partdata">共用參數</param>
        /// <returns></returns>
        private List<MessageTo> GetMsgIdsToMsgTList(List<UserSimpleInfo> idNamelist, MessageTo partdata)
        {
            List<MessageTo> list = new List<MessageTo>();

            for (int i = 0; i < idNamelist.Count; i++)
            {
                MessageTo kt = new MessageTo();
                //kt.mo_mgid =
                //收件Id
                kt.MSGOBJID = idNamelist[i].MemberId;// Convert.ToInt32(idAry[i]);
                kt.MSGOBJDate = partdata.MSGOBJDate;
                kt.MSGOBJReaderIP = partdata.MSGOBJReaderIP;
                kt.MSGOBJSendIP = partdata.MSGOBJSendIP;
                kt.MSGOBJCreateTime = partdata.MSGOBJCreateTime;
                kt.MSGOBJSendID = partdata.MSGOBJSendID;
                kt.MSGOBJNote = partdata.MSGOBJNote;
                //收件名稱
                //kt.MSGOBJName = idNamelist[i].Name  ;
                kt.MSGOBJName = idNamelist[i].Unit + "-" + idNamelist[i].Name;
                list.Add(kt);

            }
            return list;
        }
        #endregion

        #region 轉國字
        class Number2Text
        {
            private string[] numberTexts = { "零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖" }; //0-9對應的國字

            private string[] typeMappings = { "元", "拾", "佰", "仟", "萬", "拾", "佰", "仟", "億", "拾", "佰", "仟", "兆" }; //每一位數的單位

            private string[] baseTypes = { "元", "萬", "億", "兆" }; //重要的區隔單位

            private Int64 number;

            private Stack<Node> nodes;

            public Number2Text(Int64 number)
            {
                if (number <= 0)
                {
                    throw new ArgumentOutOfRangeException("number", "數值必需為正整數");
                }

                if (number.ToString().Length > 13)
                {
                    throw new ArgumentOutOfRangeException("number", "數值超過 13 位數");
                }

                this.number = number;

                int index = 0;

                this.nodes = new Stack<Node>(number.ToString().Length);

                foreach (var chr in this.number.ToString().Reverse())
                {
                    var newNode = new Node { Unit = this.typeMappings[index], Value = Int16.Parse(chr.ToString()) };

                    if (index == 0)
                    {
                        newNode.nodeType = NodeType.Root;
                    }
                    else
                    {
                        newNode.nodeType = (this.baseTypes.Contains(newNode.Unit)) ? NodeType.Base : NodeType.Number;
                    }

                    this.nodes.Push(newNode);

                    index++;
                }
            }

            /// <summary>
            /// 取得國字
            /// </summary>
            /// <returns></returns>
            public string GetText()
            {

                StringBuilder sb = new StringBuilder();
                int zeroFound = 0; //已連續出現幾個零

                while (this.nodes.Count > 0)
                {
                    var node = this.nodes.Pop();
                    switch (node.nodeType)
                    {
                        case NodeType.Base:
                            parseBase(node, sb, ref zeroFound);
                            break;
                        case NodeType.Root:
                            parseRoot(node, sb, ref zeroFound);
                            break;
                        case NodeType.Number:
                            parseNumber(node, sb, ref zeroFound);
                            break;
                    }
                }

                return sb.ToString();
            }

            private void parseBase(Node node, StringBuilder sb, ref int zeroFound)
            {
                if (node.Value == 0)
                {
                    //若base是零,且前三位同時是零才不顯示文字
                    if (zeroFound < 3)
                    {
                        sb.Append(node.Unit);
                    }
                    zeroFound = 0; //若前一位有發現零,也取消它

                    return;
                }

                parseNormal(node, sb, ref zeroFound);
            }

            private void parseRoot(Node node, StringBuilder sb, ref int zeroFound)
            {
                if (node.Value == 0)
                {
                    sb.Append(node.Unit); //若是root, 就算是零也要生成文字
                    zeroFound = 0; //若前一位有發現零,也取消它
                    return;
                }

                parseNormal(node, sb, ref zeroFound);
            }

            private void parseNumber(Node node, StringBuilder sb, ref int zeroFound)
            {
                if (node.Value == 0)
                {
                    zeroFound++;
                    return;
                }

                parseNormal(node, sb, ref zeroFound);
            }

            private void parseNormal(Node node, StringBuilder sb, ref int zeroFound)
            {
                if (zeroFound > 0) //在這個位數之前有發現零
                {
                    sb.Append("零");
                }

                sb.Append(this.numberTexts[node.Value] + node.Unit);
                zeroFound = 0;
            }

            private enum NodeType
            {
                Root, Number, Base
            }

            private class Node
            {
                public NodeType nodeType { get; set; }

                public string Unit { get; set; }

                public short Value { get; set; }
            }
        }
        #endregion
    }
}