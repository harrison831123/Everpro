using EP.H2OModels;
using EP.VLifeModels;
using EP.SD.SalesSupport.LAW.Models;
using EP.SD.SalesSupport.LAW.Service;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using EP.Web;
using EP.PSL.IB.Service;
using Microsoft.CUF;
using EP.SD.SalesSupport.LAW.Web.Areas.LAW.Utilities;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Controllers
{
    [Program("LAWTX008")]
    public class LAWTX008Controller : BaseController
    {
        // GET: LAW/LAWTX008
        private ILAWService _Service;
        public LAWTX008Controller()
        {
            _Service = ServiceHelper.Create<ILAWService>();
        }
        /// <summary>
        /// 受理作業
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX008")]
        public ActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// 法追案件匯入(正常)
        /// </summary>
        /// <param name="file">上傳的excel</param>
        /// <returns></returns>
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX008")]
        public ActionResult Upload(HttpPostedFileBase file, string tabUniqueId)
        {
            List<string> errorList = new List<string>();  //記錄錯誤
            if (file != null)
            {
                StreamReader sr = new StreamReader(file.InputStream, System.Text.Encoding.Default);
                LawNoteYmCounter noteYmCounter = new LawNoteYmCounter();
                LawNoteYmCounterType counterType = new LawNoteYmCounterType();
                LawAgentContent lawAgentContent = new LawAgentContent();
                LawAgentData lawAgentData = new LawAgentData();
                List<LawReport> lawReportList = new List<LawReport>();
                string data, LawNoteNo = string.Empty, monstr, fileno, superaccount, ymstr = string.Empty, agent_code = string.Empty, blockvm, blockvmcode;

                sr.ReadLine();//標題
                counterType.datacount = 0;

                while (!sr.EndOfStream)
                {
                    counterType.datacount++;
                    data = sr.ReadLine();
                    string[] arrData = data.Split(',');
                    //檢查資料格式
                    if (arrData.Length != 10)
                    {
                        errorList.Add("案件匯入格式不符!請重新檢查文件(欄位中勿使用半形逗點符號!)、欄位數目!");
                        break;
                    }

                    LawReport lawReport = new LawReport();
                    lawReport.Order = arrData[0];
                    lawReport.Year = arrData[1];
                    lawReport.Month = arrData[2];
                    lawReport.PaySequence = arrData[3];
                    lawReport.Fileno = arrData[4];
                    //string DueName = Member.Get(arrData[6]).Name;
                    lawReport.DueName = arrData[5];
                    lawReport.Agentcode = arrData[6];
                    lawReport.DueMoney = arrData[7];
                    lawReport.SuperAccount = arrData[8];
                    lawReport.DueReason = arrData[9];
                    lawReportList.Add(lawReport);

                    monstr = arrData[2];
                    if (monstr.Length == 1)
                    {
                        monstr = "0" + monstr;
                    }
                    switch (arrData[1].ToString().Length)
                    {
                        case 2:
                            ymstr = "00" + arrData[1] + "/" + monstr;
                            break;
                        case 3:
                            ymstr = "0" + arrData[1] + "/" + monstr;
                            break;
                        case 4:
                            ymstr = arrData[1] + "/" + monstr;
                            break;
                    }

                    //檢查和撈取業務員資料
                    OrgibDetail orgib = _Service.CheckOrgib(arrData[6].Trim(), ymstr, arrData[3].Trim());
                    if (orgib == null)
                    {
                        errorList.Add(string.Format("檢測到資料列第{0}次序，查無業務員相關資料!。", counterType.datacount));
                    }
                }

                if (errorList.Count == 0)
                {
                    #region 新增資料
                    for (int i = 0; i < lawReportList.Count; i++)
                    {
                        monstr = lawReportList[i].Month;
                        if (monstr.Length == 1)
                        {
                            monstr = "0" + monstr;
                        }
                        switch (lawReportList[i].Year.ToString().Length)
                        {
                            case 2:
                                ymstr = "00" + lawReportList[i].Year + "/" + monstr;
                                break;
                            case 3:
                                ymstr = "0" + lawReportList[i].Year + "/" + monstr;
                                break;
                            case 4:
                                ymstr = lawReportList[i].Year + "/" + monstr;
                                break;
                        }
                        fileno = lawReportList[i].Fileno;
                        if (fileno.Length == 1)
                        {
                            if (fileno == "無")
                            {
                            }
                            else
                            {
                                fileno = "0" + fileno;
                            }
                        }
                        superaccount = lawReportList[i].SuperAccount;
                        noteYmCounter = _Service.CheckLawNoteYmCounter(lawReportList[i].Year, monstr);

                        #region 新增照會單
                        if (noteYmCounter == null)
                        {                           
                            if (lawReportList[i].Month.Length < 2)
                            {
                                LawNoteNo = lawReportList[i].Year + "照0" + lawReportList[i].Month + "001";
                            }
                            else
                            {
                                LawNoteNo = lawReportList[i].Year + "照" + lawReportList[i].Month + "001";
                            }
                            noteYmCounter = new LawNoteYmCounter();
                            noteYmCounter.LawNoteYear = lawReportList[i].Year;
                            noteYmCounter.LawNoteMonth = monstr;
                            _Service.InsertLawNoteYmCounter(noteYmCounter);
                        }
                        else
                        {
                            if (lawReportList[i].Month.Length < 2)
                            {
                                switch ((noteYmCounter.LawNoteCounter + 1).ToString().Length)
                                {
                                    case 1:
                                        LawNoteNo = lawReportList[i].Year + "照" + "0" + lawReportList[i].Month + "00" + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                    case 2:
                                        LawNoteNo = lawReportList[i].Year + "照" + "0" + lawReportList[i].Month + "0" + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                    case 3:
                                        LawNoteNo = lawReportList[i].Year + "照" + "0" + lawReportList[i].Month + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                }
                            }
                            else
                            {
                                switch ((noteYmCounter.LawNoteCounter + 1).ToString().Length)
                                {
                                    case 1:
                                        LawNoteNo = lawReportList[i].Year + "照" + lawReportList[i].Month + "00" + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                    case 2:
                                        LawNoteNo = lawReportList[i].Year + "照" + lawReportList[i].Month + "0" + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                    case 3:
                                        LawNoteNo = lawReportList[i].Year + "照" + lawReportList[i].Month + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                }
                            }
                            _Service.UpdateLawNoteYmCounter(noteYmCounter);
                        }
                        #endregion

                        #region 新增法追主檔表
                        LawContent content = new LawContent();
                        content.LawNoteNo = LawNoteNo;
                        content.LawYear = lawReportList[i].Year;
                        content.LawMonth = monstr;
                        content.LawPaySequence = Convert.ToInt32(lawReportList[i].PaySequence);
                        content.LawFileNo = fileno;
                        content.LawDueName = lawReportList[i].DueName;
                        content.LawDueAgentId = lawReportList[i].Agentcode;
                        content.LawDueMoney = Convert.ToInt32(lawReportList[i].DueMoney);
                        content.LawSuperAccount = superaccount;
                        content.LawDueReason = lawReportList[i].DueReason;
                        content.LawStatusType = 0;
                        content.LawDoUnitId = 1;
                        content.LawDoUnitName = LAWHelper.ChangeUnitName(User.MemberInfo.ID); //"法務室";
                        content.LawContentCreatorID = User.MemberInfo.ID;
                        content.LawContentCreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
                        content.LawContentCreateDate = DateTime.Now.ToString("yyyy/MM/dd");
                        _Service.InsertLawContent(content);
                        #endregion

                        //檢查八大區塊
                        orgsmb orgsmb = _Service.Checkorgsmb(lawReportList[i].Agentcode.Trim());
                        if (orgsmb == null)
                        {
                            blockvm = "NULL";
                            blockvmcode = "NULL";
                        }
                        else
                        {
                            blockvm = orgsmb.VmName;
                            blockvmcode = orgsmb.VmCode;
                        }

                        //撈取業務員資料
                        OrgibDetail orgib = _Service.CheckOrgib(lawReportList[i].Agentcode.Trim(), ymstr, lawReportList[i].PaySequence.Trim());

                        agent_code = lawReportList[i].Agentcode.Substring(0, 10);

                        adin adin = _Service.Getadin(agent_code);
                        switch (adin.address_type)
                        {
                            case "1": //戶籍
                                lawAgentData.Address = adin.address;
                                break;
                            case "2": //住家
                                lawAgentData.Address01 = adin.address;
                                lawAgentData.Phone01 = adin.phone;
                                break;
                        }

                        #region 組業務員資料檔
                        LawAgentContent agentcontent = new LawAgentContent();
                        agentcontent.LawNoteNo = LawNoteNo;
                        agentcontent.ProductionYm = ymstr;
                        agentcontent.Sequence = lawReportList[i].PaySequence;
                        agentcontent.VmCode = blockvmcode;
                        agentcontent.VmName = blockvm;
                        agentcontent.SmCode = orgib.sm_code;
                        agentcontent.SmName = orgib.sm_name;
                        agentcontent.WcCenter = orgib.wc_center;
                        agentcontent.WcCenterName = orgib.wc_center_name;
                        agentcontent.CenterCode = orgib.center_code;
                        agentcontent.CenterName = orgib.center_name;
                        agentcontent.AdministratId = orgib.administrat_id;
                        agentcontent.AdminName = orgib.admin_name;
                        agentcontent.AdminLevel = orgib.admin_level;
                        agentcontent.AgentCode = orgib.agent_code;
                        agentcontent.Name = orgib.names;
                        agentcontent.AgStatusCode = orgib.ag_status_code;
                        agentcontent.AgLevel = orgib.ag_level;
                        agentcontent.AgLevelName = orgib.level_name_chs;
                        agentcontent.RecordDate = orgib.record_date;
                        agentcontent.RegisterDate = orgib.register_date;
                        agentcontent.SuperAccount = lawReportList[i].SuperAccount;
                        agentcontent.AgStatusDate = orgib.ag_status_date;
                        agentcontent.CreateDate = DateTime.Now.ToString("yyyy/MM/dd");
                        #endregion

                        #region 組業務員個人資料檔
                        LawAgentData agentdata = new LawAgentData();
                        agentdata.AgentCode = agent_code;
                        agentdata.Name = orgib.names;
                        agentdata.Birth = orgib.birth;
                        agentdata.Cell01 = orgib.cellur_phone_no;
                        agentdata.Phone01 = lawAgentData.Phone01;
                        agentdata.Address = lawAgentData.Address;
                        agentdata.Address01 = lawAgentData.Address01;
                        agentdata.SuperAccount = lawReportList[i].SuperAccount;
                        agentdata.CreateDate = DateTime.Now.ToString("yyyy/MM/dd");
                        #endregion

                        //新增業務員資料檔案
                        _Service.InsertLawAgentContent(agentcontent);

                        //檢查個人資料是否重複
                        LawAgentData lawAgentRepeat = _Service.CheckLawAgentDataRepeat(agent_code);
                        //新增或更新業務員個人資料
                        if (lawAgentRepeat != null)
                        {
                            if (lawAgentRepeat.Name != orgib.names)
                            {
                                _Service.UpdateLawAgentContent(orgib.agent_code, orgib.names, lawAgentRepeat.Name);
                                _Service.UpdateLawAgentData(orgib.agent_code, orgib.names, lawAgentRepeat.Name);
                            }

                        }
                        else
                        {
                            _Service.InsertLawAgentData(agentdata);
                        }

                        #region 新增留言                          
                        int m = 0;
                        string ag;
                        string dir1 = string.Empty, dir2 = string.Empty, dir3 = string.Empty, dirn1 = string.Empty, dirn2 = string.Empty, dirn3 = string.Empty, dirc1 = string.Empty, dirc2 = string.Empty, dirc3 = string.Empty, dirl1 = string.Empty, dirl2 = string.Empty, dirl3 = string.Empty;
                        string retMsg = string.Empty, MSID;
                        List<string> directorIDs = new List<string>();
                        var MsgService = ServiceHelper.Create<IMessageService>();
                        ag = orgib.agent_code;
                        //取三代主管資料
                        while (ag != "Z99999999901")
                        {
                            OrginDetail orginDetail = _Service.Getorgin(ag);
                            if (orginDetail != null)
                            {
                                ag = orginDetail.director_id;
                                if (orginDetail.dir_status_code == "1")
                                {
                                    m++;
                                    switch (m)
                                    {
                                        case 1:
                                            dir1 = orginDetail.director_id; //第1代
                                            dirn1 = orginDetail.director_name; //第1代姓名
                                            dirc1 = orginDetail.center_name; //第1代實駐
                                            dirl1 = orginDetail.level_name_chs; //第1代職級
                                            directorIDs.Add(dir1);
                                            break;
                                        case 2:
                                            dir2 = orginDetail.director_id; //第2代
                                            dirn2 = orginDetail.director_name; //第2代姓名
                                            dirc2 = orginDetail.center_name; //第2代實駐    
                                            dirl2 = orginDetail.level_name_chs; //第2代職級
                                            directorIDs.Add(dir2);
                                            break;
                                        case 3:
                                            dir3 = orginDetail.director_id; //第3代
                                            dirn3 = orginDetail.director_name; //第3代姓名
                                            dirc3 = orginDetail.center_name; //第3代實駐
                                            dirl3 = orginDetail.level_name_chs; //第3代職級
                                            directorIDs.Add(dir3);
                                            break;
                                    }
                                }
                            }
                        }
                        //新增留言
                        Message k = new Message();
                        k.MSGSubject = DateTime.Now.ToString("yyyy/MM/dd") + "【" + LawNoteNo + "】：法追照會單通知作業";
                        k.MSGDESC = LawNoteNo + "照會單通知，請至法追業務系統通知作業下載照會單，以利法追作業進行。";
                        k.MSGCreater = User.MemberInfo.ID;// (int)Session["orgID"];
                        k.MSGCreateName = Microsoft.CUF.CodeName.GetUnitName(User.UnitInfo.ID).Trim() + "-" + Microsoft.CUF.CodeName.GetMemberName(User.MemberInfo.ID);
                        k.MSGCreateIP = PlatformHelper.GetClientIPv4();
                        k.MSGTime = DateTime.Now;
                        k.MSGClass = 3;
                        k.MSGNote = LawNoteNo;
                        //產生留言人員名單
                        List<MessageTo> kt_list = null;
                        List<MessageFile> kf_list = new List<MessageFile>();
                        MessageTo partdata = new MessageTo();
                        partdata.MSGOBJDate = k.MSGTime;
                        partdata.MSGOBJReaderIP = "";
                        partdata.MSGOBJSendIP = k.MSGCreateIP;
                        partdata.MSGOBJCreateTime = k.MSGTime;
                        partdata.MSGOBJSendID = User.MemberInfo.ID;
                        partdata.MSGOBJNote = LawNoteNo;
                        kt_list = GetMsgIdsToMsgTList(directorIDs, partdata);
                        MSID = MsgService.CreateMessage(k, kf_list, kt_list, TabUniqueId, out retMsg);

                        //新增照會單資料
                        for (int j = 0; j < directorIDs.Count; j++)
                        {
                            #region 組照會單資料檔
                            LawNote lawNote = new LawNote();
                            lawNote.LawNoteNo = LawNoteNo;
                            lawNote.LawNoteKm = Convert.ToInt32(MSID);
                            lawNote.LawNoteName = orgib.names;
                            lawNote.LawNoteDep = LAWHelper.ChangeUnitName(User.MemberInfo.ID); //User.UnitInfo.Name;
                            switch (j)
                            {
                                case 0:
                                    lawNote.LawNoteTo = dirn1;
                                    lawNote.LawNoteCenter = dirc1;
                                    lawNote.LawNoteLevel = dirl1;
                                    break;
                                case 1:
                                    lawNote.LawNoteTo = dirn2;
                                    lawNote.LawNoteCenter = dirc2;
                                    lawNote.LawNoteLevel = dirl2;
                                    break;
                                case 2:
                                    lawNote.LawNoteTo = dirn3;
                                    lawNote.LawNoteCenter = dirc3;
                                    lawNote.LawNoteLevel = dirl3;
                                    break;
                            }
                            LawDoUser douser = _Service.GetLawDouserTopOne();
                            lawNote.LawNotePro = douser.DouserName;
                            lawNote.LawNoteTel = douser.DouserPhoneExt;
                            lawNote.LawNoteType = 0;
                            lawNote.NoteCreatorID = User.MemberInfo.ID;
                            lawNote.LawNoteCreatername = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
                            lawNote.LawNoteCreatedate = DateTime.Now;

                            #endregion

                            _Service.InsertLawNote(lawNote);
                        }
                        #endregion
                    }
                    #endregion


                    AppendMessage("上傳成功!", false);
                    return Json(counterType.datacount);
                }
                else
                {
                    AppendMessage("上傳失敗!", false);
                    return Json(errorList);
                }

            }
            else
            {
                AppendMessage("檔案不可為空!", false);
                return Json(errorList);
            }
            return Json(errorList);
        }

        /// <summary>
        /// 法追案件匯入(舊案)
        /// </summary>
        /// <param name="file">上傳的excel</param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.LAW.LAWTX008")]
        public ActionResult UploadOld(HttpPostedFileBase file, string tabUniqueId)
        {
            List<string> errorList = new List<string>();  //記錄錯誤
            if (file != null)
            {
                StreamReader sr = new StreamReader(file.InputStream, System.Text.Encoding.Default);
                LawNoteYmCounter noteYmCounter = new LawNoteYmCounter();
                LawNoteYmCounterType counterType = new LawNoteYmCounterType();
                LawAgentContent lawAgentContent = new LawAgentContent();
                LawAgentData lawAgentData = new LawAgentData();
                List<LawReport> lawReportList = new List<LawReport>();
                string data, LawNoteNo = string.Empty, monstr, fileno, superaccount, ymstr = string.Empty, agent_code = string.Empty, blockvm, blockvmcode;
                sr.ReadLine();//標題
                counterType.datacount = 0;

                while (!sr.EndOfStream)
                {
                    counterType.datacount++;
                    data = sr.ReadLine();
                    string[] arrData = data.Split(',');
                    //檢查資料格式
                    if (arrData.Length != 12)
                    {
                        errorList.Add("案件匯入格式不符!請重新檢查文件(欄位中勿使用半形逗點符號!)或欄位數目!");
                        break;
                    }

                    LawReport lawReport = new LawReport();
                    lawReport.Order = arrData[0];
                    lawReport.Year = arrData[1];
                    lawReport.Month = arrData[2];
                    lawReport.PaySequence = arrData[3];
                    lawReport.Fileno = arrData[4];
                    lawReport.DueName = arrData[5];
                    lawReport.Agentcode = arrData[6];
                    lawReport.DueMoney = arrData[7];
                    lawReport.SuperAccount = arrData[8];
                    lawReport.DueReason = arrData[9];
                    lawReport.Litigationprogress = arrData[10];
                    lawReport.Doprogress = arrData[11];

                    lawReportList.Add(lawReport);

                    monstr = arrData[2];
                    if (monstr.Length == 1)
                    {
                        monstr = "0" + monstr;
                    }
                    switch (arrData[1].ToString().Length)
                    {
                        case 2:
                            ymstr = "00" + arrData[1] + "/" + monstr;
                            break;
                        case 3:
                            ymstr = "0" + arrData[1] + "/" + monstr;
                            break;
                        case 4:
                            ymstr = arrData[1] + "/" + monstr;
                            break;
                    }

                    //檢查和撈取業務員資料
                    OrgibDetail orgib = _Service.CheckOrgib(arrData[6].Trim(), ymstr, arrData[3].Trim());
                    if (orgib == null)
                    {
                        errorList.Add(string.Format("檢測到資料列第{0}次序，查無業務員相關資料!。", counterType.datacount));
                    }
                }

                if (errorList.Count == 0)
                {
                    #region 新增資料
                    for (int i = 0; i < lawReportList.Count; i++)
                    {
                        monstr = lawReportList[i].Month;
                        if (monstr.Length == 1)
                        {
                            monstr = "0" + monstr;
                        }
                        switch (lawReportList[i].Year.ToString().Length)
                        {
                            case 2:
                                ymstr = "00" + lawReportList[i].Year + "/" + monstr;
                                break;
                            case 3:
                                ymstr = "0" + lawReportList[i].Year + "/" + monstr;
                                break;
                            case 4:
                                ymstr = lawReportList[i].Year + "/" + monstr;
                                break;
                        }
                        fileno = lawReportList[i].Fileno;
                        if (fileno.Length == 1)
                        {
                            if (fileno == "無")
                            {
                            }
                            else
                            {
                                fileno = "0" + fileno;
                            }
                        }
                        superaccount = lawReportList[i].SuperAccount;
                        noteYmCounter = _Service.CheckLawNoteYmCounter(lawReportList[i].Year, monstr);

                        #region 新增照會單
                        if (noteYmCounter == null)
                        {
                            if (lawReportList[i].Month.Length < 2)
                            {
                                LawNoteNo = lawReportList[i].Year + "照0" + lawReportList[i].Month + "001";
                            }
                            else
                            {
                                LawNoteNo = lawReportList[i].Year + "照" + lawReportList[i].Month + "001";
                            }
                            noteYmCounter = new LawNoteYmCounter();
                            noteYmCounter.LawNoteYear = lawReportList[i].Year;
                            noteYmCounter.LawNoteMonth = monstr;
                            _Service.InsertLawNoteYmCounter(noteYmCounter);
                        }
                        else
                        {
                            if (lawReportList[i].Month.Length < 2)
                            {
                                switch ((noteYmCounter.LawNoteCounter + 1).ToString().Length)
                                {
                                    case 1:
                                        LawNoteNo = lawReportList[i].Year + "照" + "0" + lawReportList[i].Month + "00" + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                    case 2:
                                        LawNoteNo = lawReportList[i].Year + "照" + "0" + lawReportList[i].Month + "0" + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                    case 3:
                                        LawNoteNo = lawReportList[i].Year + "照" + "0" + lawReportList[i].Month + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                }
                            }
                            else
                            {
                                switch ((noteYmCounter.LawNoteCounter + 1).ToString().Length)
                                {
                                    case 1:
                                        LawNoteNo = lawReportList[i].Year + "照" + lawReportList[i].Month + "00" + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                    case 2:
                                        LawNoteNo = lawReportList[i].Year + "照" + lawReportList[i].Month + "0" + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                    case 3:
                                        LawNoteNo = lawReportList[i].Year + "照" + lawReportList[i].Month + (noteYmCounter.LawNoteCounter + 1);
                                        break;
                                }
                            }
                            _Service.UpdateLawNoteYmCounter(noteYmCounter);
                        }
                        #endregion

                        #region 新增法追主檔表
                        LawContent content = new LawContent();
                        content.LawNoteNo = LawNoteNo;
                        content.LawYear = lawReportList[i].Year;
                        content.LawMonth = monstr;
                        content.LawPaySequence = Convert.ToInt32(lawReportList[i].PaySequence);
                        content.LawFileNo = fileno;
                        content.LawDueName = lawReportList[i].DueName;
                        content.LawDueAgentId = lawReportList[i].Agentcode;
                        content.LawDueMoney = Convert.ToInt32(lawReportList[i].DueMoney);
                        content.LawSuperAccount = superaccount;
                        content.LawDueReason = lawReportList[i].DueReason;
                        content.LawStatusType = 0;
                        content.LawDoUnitId = 1;
                        content.LawDoUnitName = LAWHelper.ChangeUnitName(User.MemberInfo.ID); //"法務室";
                        content.LawContentCreatorID = User.MemberInfo.ID;
                        content.LawContentCreateName = LAWHelper.ChangeUnitName(User.MemberInfo.ID) + " " + User.MemberInfo.Name;
                        content.LawContentCreateDate = DateTime.Now.ToString("yyyy/MM/dd");
                        _Service.InsertLawContent(content);
                        #endregion

                        //檢查八大區塊
                        orgsmb orgsmb = _Service.Checkorgsmb(lawReportList[i].Agentcode.Trim());
                        if (orgsmb == null)
                        {
                            blockvm = "NULL";
                            blockvmcode = "NULL";
                        }
                        else
                        {
                            blockvm = orgsmb.VmName;
                            blockvmcode = orgsmb.VmCode;
                        }

                        //撈取業務員資料
                        OrgibDetail orgib = _Service.CheckOrgib(lawReportList[i].Agentcode.Trim(), ymstr, lawReportList[i].PaySequence.Trim());

                        agent_code = lawReportList[i].Agentcode.Substring(0, 10);

                        adin adin = _Service.Getadin(agent_code);
                        switch (adin.address_type)
                        {
                            case "1": //戶籍
                                lawAgentData.Address = adin.address;
                                break;
                            case "2": //住家
                                lawAgentData.Address01 = adin.address;
                                lawAgentData.Phone01 = adin.phone;
                                break;
                        }

                        #region 組業務員資料檔
                        LawAgentContent agentcontent = new LawAgentContent();
                        agentcontent.LawNoteNo = LawNoteNo;
                        agentcontent.ProductionYm = ymstr;
                        agentcontent.Sequence = lawReportList[i].PaySequence;
                        agentcontent.VmCode = blockvmcode;
                        agentcontent.VmName = blockvm;
                        agentcontent.SmCode = orgib.sm_code;
                        agentcontent.SmName = orgib.sm_name;
                        agentcontent.WcCenter = orgib.wc_center;
                        agentcontent.WcCenterName = orgib.wc_center_name;
                        agentcontent.CenterCode = orgib.center_code;
                        agentcontent.CenterName = orgib.center_name;
                        agentcontent.AdministratId = orgib.administrat_id;
                        agentcontent.AdminName = orgib.admin_name;
                        agentcontent.AdminLevel = orgib.admin_level;
                        agentcontent.AgentCode = orgib.agent_code;
                        agentcontent.Name = orgib.names;
                        agentcontent.AgStatusCode = orgib.ag_status_code;
                        agentcontent.AgLevel = orgib.ag_level;
                        agentcontent.AgLevelName = orgib.level_name_chs;
                        agentcontent.RecordDate = orgib.record_date;
                        agentcontent.RegisterDate = orgib.register_date;
                        agentcontent.SuperAccount = lawReportList[i].SuperAccount;
                        agentcontent.AgStatusDate = orgib.ag_status_date;
                        agentcontent.CreateDate = DateTime.Now.ToString("yyyy/MM/dd");
                        #endregion

                        #region 組業務員個人資料檔
                        LawAgentData agentdata = new LawAgentData();
                        agentdata.AgentCode = agent_code;
                        agentdata.Name = orgib.names;
                        agentdata.Birth = orgib.birth;
                        agentdata.Cell01 = orgib.cellur_phone_no;
                        agentdata.Phone01 = lawAgentData.Phone01;
                        agentdata.Address = lawAgentData.Address;
                        agentdata.Address01 = lawAgentData.Address01;
                        agentdata.SuperAccount = lawReportList[i].SuperAccount;
                        agentdata.CreateDate = DateTime.Now.ToString("yyyy/MM/dd");
                        #endregion

                        //新增業務員資料檔案
                        _Service.InsertLawAgentContent(agentcontent);

                        //檢查個人資料是否重複
                        LawAgentData lawAgentRepeat = _Service.CheckLawAgentDataRepeat(agent_code);
                        //新增或更新業務員個人資料
                        if (lawAgentRepeat != null)
                        {
                            if (lawAgentRepeat.Name != orgib.names)
                            {
                                _Service.UpdateLawAgentContent(orgib.agent_code, orgib.names, lawAgentRepeat.Name);
                                _Service.UpdateLawAgentData(orgib.agent_code, orgib.names, lawAgentRepeat.Name);
                            }

                        }
                        else
                        {
                            _Service.InsertLawAgentData(agentdata);
                        }
                        LawContent lawContent = _Service.GetLawContentByLawNoteNo(LawNoteNo);
                        #region 新增訴訟程序檔 
                        if (!String.IsNullOrEmpty(lawReportList[i].Litigationprogress))
                        {
                            LawLitigationProgress litigationProgress = new LawLitigationProgress();
                            litigationProgress.LawId = lawContent.LawId;
                            litigationProgress.LawLitigationprogress = lawReportList[i].Litigationprogress;
                            litigationProgress.LawLitigationProgressCreatorID = User.MemberInfo.ID;
                            litigationProgress.LawNoteNo = LawNoteNo;
                            litigationProgress.CreateDate = DateTime.Now;

                            _Service.InsertLawLitigationProgress(litigationProgress);
                        }
                        #endregion

                        #region 新增執行程序檔     
                        if (!String.IsNullOrEmpty(lawReportList[i].Doprogress))
                        {
                            LawDoProgress doProgress = new LawDoProgress();
                            doProgress.LawId = lawContent.LawId;
                            doProgress.LawNoteNo = LawNoteNo;
                            doProgress.LawDoprogress = lawReportList[i].Doprogress;
                            doProgress.LawDoProgressCreatorID = User.MemberInfo.ID;
                            doProgress.CreateDate = DateTime.Now;
                            
                            _Service.InsertLawDoProgress(doProgress);
                        }
                        #endregion
                    }
                    #endregion


                    AppendMessage("上傳成功!", false);
                    return Json(counterType.datacount);
                }
                else
                {
                    AppendMessage("上傳失敗!", false);
                    return Json(errorList);
                }

            }
            else
            {
                AppendMessage("檔案不可為空!", false);
                return Json(errorList);
            }
        }

        /// <summary>
        /// 組成留言接收者清單 
        /// </summary>
        /// <param name="idNamelist">接收者資料</param>
        /// <param name="partdata">共用參數</param>
        /// <returns></returns>
        private List<MessageTo> GetMsgIdsToMsgTList(List<string> idNamelist, MessageTo partdata)
        {
            List<MessageTo> list = new List<MessageTo>();

            for (int i = 0; i < idNamelist.Count; i++)
            {
                MessageTo kt = new MessageTo();
                kt.MSGOBJID = idNamelist[i];// Convert.ToInt32(idAry[i]);
                kt.MSGOBJDate = partdata.MSGOBJDate;
                kt.MSGOBJReaderIP = partdata.MSGOBJReaderIP;
                kt.MSGOBJSendIP = partdata.MSGOBJSendIP;
                kt.MSGOBJCreateTime = partdata.MSGOBJCreateTime;
                kt.MSGOBJSendID = partdata.MSGOBJSendID;
                kt.MSGOBJName = Unit.Get(Member.Get(idNamelist[i]).UnitGUID).Name + "-" + Member.Get(idNamelist[i]).Name;
                kt.MSGOBJNote = partdata.MSGOBJNote;
                list.Add(kt);
            }
            return list;
        }
    }
}