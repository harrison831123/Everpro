using EP.H2OModels;
using EP.Platform.Service;
using EP.PSL.IB.Service;
using EP.PSL.WorkResources.MeetingMng.Service;
using EP.PSL.WorkResources.MeetingMng.Web.Areas.MeetingMng.Model;
using EP.Web;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace EP.PSL.WorkResources.MeetingMng.Web.Areas.MeetingMng.Controllers
{
    public class MMTX001Controller : BaseController
    {
        // GET: MeetingMng/MMTX001
        private IMeetingMngService _Service;
        public MMTX001Controller()
        {
            _Service = ServiceHelper.Create<IMeetingMngService>();
        }
        #region 新增會議
        /// <summary>
        /// 新增會議 (GET)
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.PSL.WorkResources.MeetingMng.MMTX001")]
        public ActionResult Create()
        {
            MyPageStatus = PageStatus.Create;
            MeetingDetailModel m = new MeetingDetailModel();
            m.MTConvenerName = User.UnitInfo.Name + "-" + User.MemberInfo.Name;
            m.MTStartDate = DateTime.Today;
            m.MTEndDate = DateTime.Today;

            return View(m);
        }
        /// <summary>
        /// 新增會議處理
        /// </summary>
        /// <param name="viewmodel"></param>
        [HttpPost]
        //[ValidateInput(false)]
        [HasPermission("EP.PSL.WorkResources.MeetingMng.MMTX001")]
        public ActionResult Create(MeetingDetailModel viewmodel)
        {
            if (viewmodel == null)
                Throw.BusinessError(EP.PlatformResources.請傳入指定參數資料);
            string retMsg = string.Empty;

            //EP.PSL.IB.Web.IBHelper.GetMeetingDeviceList();

            //初始宣告:當發生Error時頁面才可以保留原始資訊
            this.DefaultView.ViewName = "Create";
            //this.TempData["model"] = viewmodel;
            MyPageStatus = PageStatus.Create;

            //if (!ModelState.IsValid)
            //{
            //Throw.BusinessError("檢核失敗");
            //}    

            var service = ServiceHelper.Create<IMeetingMngService>();


            //處理場地與設備MDID
            int MDID = Convert.ToInt32(viewmodel.MDName);
            //var MDIDList = string.IsNullOrWhiteSpace(MDID) ? new List<int>() :
            //    MDID.Split(',')
            //        .Select(it => int.Parse(it))
            //        .ToList();

            #region 判斷時間重複改寫至GetMDIDToChk
            //List<MeetingDeviceRelation> MDRlistchk = GetDeviceToList(MDIDList, viewmodel.MTID);
            MeetingDetail timechk = new MeetingDetail();
            viewmodel.MTStartDate = new DateTime(viewmodel.MTStartDate.Year, viewmodel.MTStartDate.Month, viewmodel.MTStartDate.Day,Convert.ToInt32(viewmodel.shour), Convert.ToInt32(viewmodel.smin), 0);
            viewmodel.MTEndDate = new DateTime(viewmodel.MTEndDate.Year, viewmodel.MTEndDate.Month, viewmodel.MTEndDate.Day, Convert.ToInt32(viewmodel.ehour), Convert.ToInt32(viewmodel.emin), 0);
            timechk = service.ChkMeetingTime(MDID, viewmodel.MTStartDate, viewmodel.MTEndDate);
            if (timechk != null)
            {
                Throw.BusinessError("與 " + timechk.MTName + " 會議衝突，會議召開時間: " + viewmodel.MTStartDate + "~" + viewmodel.MTEndDate + "場地(設備): " + timechk.MDName + " 已被選用!!無法召開會議!");

            }
            #endregion

            Meeting modelMT = new Meeting();
            modelMT.MTDesc = System.Web.HttpUtility.HtmlDecode(viewmodel.MTDesc);
            modelMT.MTStartDate = Convert.ToDateTime(viewmodel.MTStartDate.ToString("yyyy/MM/dd") + " " + viewmodel.shour + ":" + viewmodel.smin);
            modelMT.MTEndDate = Convert.ToDateTime(viewmodel.MTEndDate.ToString("yyyy/MM/dd") + " " + viewmodel.ehour + ":" + viewmodel.emin);
            modelMT.MTCreatedate = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            modelMT.MTConvener = User.MemberInfo.ID;
            //顯示部門 User.UnitInfo.GetParent().Name;
            modelMT.MTConvenerName = User.UnitInfo.Name + "-" + User.MemberInfo.Name;
            modelMT.MTActive = 1;
            modelMT.MTName = viewmodel.MTName;
            //判斷地點是由下拉式選單新增或Textbox新增
            //if (viewmodel.MTPlace != null)
            //{
            //    modelMT.MTPlace = viewmodel.MTPlace;
            //}
            if (viewmodel.MTNewPlace != null)
            {
                modelMT.MTPlace = viewmodel.MTNewPlace;
            }
            else
            {
                modelMT.MTPlace = "";
            }
            //設定成員
            List<AccountGroupJsonString> JsonClassLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(viewmodel.RecipientToJson);
            List<UserSimpleInfo> idnameLst = PlatformHelper.GetRecipientUsersList(JsonClassLst);
            List<UserSimpleInfo> idnameNewLst = new List<UserSimpleInfo>();
            //設定主席
            List<AccountGroupJsonString> JsonChairmanLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(viewmodel.MTChairmanName);
            List<UserSimpleInfo> Chairman = PlatformHelper.GetRecipientUsersList(JsonChairmanLst);
            UserSimpleInfo UChairman = new UserSimpleInfo();
            for (int i = 0; i < Chairman.Count; i++)
            {
                if (Chairman.Count >= 2)
                {
                    TempData["message"] = "主席只能選一個";
                    //TempData["model"] = viewmodel;
                    return RedirectToAction("Create");
                }
                modelMT.MTChairman = Chairman[i].MemberId;// Convert.ToInt32(idAry[i]);  
                modelMT.MTChairmanName = Chairman[i].Unit + "-" + Chairman[i].Name;
                UChairman.MemberId = Chairman[i].MemberId;
                UChairman.Unit = Chairman[i].Unit;
                UChairman.Name = Chairman[i].Name;
                idnameLst.Add(UChairman);
            }
            //設定紀錄者
            List<AccountGroupJsonString> JsonRecorderLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(viewmodel.MTRecorderName);
            List<UserSimpleInfo> Recorder = PlatformHelper.GetRecipientUsersList(JsonRecorderLst);
            UserSimpleInfo URecorder = new UserSimpleInfo();
            for (int i = 0; i < Recorder.Count; i++)
            {
                if (Recorder.Count >= 2)
                {
                    TempData["message"] = "紀錄者只能選一個";
                    //TempData["model"] = viewmodel;
                    return RedirectToAction("Create");
                }
                modelMT.MTRecorder = Recorder[i].MemberId;// Convert.ToInt32(idAry[i]);  
                modelMT.MTRecorderName = Recorder[i].Unit + "-" + Recorder[i].Name;
                URecorder.MemberId = Recorder[i].MemberId;
                URecorder.Unit = Recorder[i].Unit;
                URecorder.Name = Recorder[i].Name;
                idnameLst.Add(URecorder);
            }
            //存入會議資訊
            int MTID = service.CreateMeeting(modelMT);

            List<MeetingMemberRelation> MMR_list = null;
            UserSimpleInfo UConvener = new UserSimpleInfo();
            //判斷召集人是否為成員
            if (viewmodel.Convenerchk == true)
            {
                UConvener.MemberId = User.MemberInfo.ID;
                UConvener.Unit = User.UnitInfo.Name;
                UConvener.Name = User.MemberInfo.Name;
                idnameLst.Add(UConvener);
            }
            //去除重複人選   
            List<string> DistinctList = idnameLst.Select(x => x.MemberId).Distinct().ToList();
            for (int i = 0; i < DistinctList.Count; i++)
            {
                UserSimpleInfo user = new UserSimpleInfo();
                if (DistinctList[i] == idnameLst[i].MemberId)
                {
                    user.AccountId = idnameLst[i].AccountId;
                    user.MemberId = idnameLst[i].MemberId;
                    user.Name = idnameLst[i].Name;
                    user.Unit = idnameLst[i].Unit;
                    idnameNewLst.Add(user);
                }
            }

            MMR_list = GetIdsToList(idnameNewLst, MTID);
            //存入會議成員
            service.CreateMeetingMemberRelation(MMR_list);
            MeetingDeviceRelation MDmodel = new MeetingDeviceRelation();
            MDmodel.MTID = MTID;
            MDmodel.MDID = MDID;
            //MDR_list = GetDeviceToList(MDIDList, MTID);

            //存入會議設備
            service.CreateMeetingDeviceRelation(MDmodel);

            // 產生附加檔案資料
            List<MeetingFile> MF_list = new List<MeetingFile>();
            if (viewmodel.UploadFilesName != null && viewmodel.UploadFilesName.Length > 0)
            {
                string fnames = viewmodel.UploadFilesName;

                MF_list = (from item in fnames.Split('*')
                           let parts = item.Split('|')
                           select new MeetingFile { MFFileName = parts[0], MFMd5Name = parts[1] }).ToList();
                for (int i = 0; i < MF_list.Count; i++)
                {
                    MF_list[i].MTID = MTID;
                    MF_list[i].MFCreater = modelMT.MTConvener;
                    MF_list[i].MFCreateDate = modelMT.MTCreatedate;
                    MF_list[i].MFFilePath = DateTime.Now.ToString("yyyyMM") + @"\" + MTID;
                    MF_list[i].MFType = 1;
                }
            }
            service.CreateMeetingFile(modelMT, MF_list, viewmodel.TabUniqueId, out retMsg);

            //設定會議管理&設備中心預定的會議室和設備
            List<MeetingOrder> MO_List = new List<MeetingOrder>();

            MeetingOrder modelMO = new MeetingOrder();
            modelMO.imember = User.MemberInfo.ID;
            modelMO.iunit = User.UnitInfo.ID;
            modelMO.MOCheck = 1;
            modelMO.MOCreateDate = DateTime.Now;
            modelMO.MOTitle = "會議:" + viewmodel.MTName;
            modelMO.MODesc = modelMT.MTDesc + "(此資料由會議管理模組新增)";
            modelMO.MOStartDate = Convert.ToDateTime(viewmodel.MTStartDate.ToString("yyyy/MM/dd") + " " + viewmodel.shour + ":" + viewmodel.smin);
            modelMO.MOLendState = 1;
            modelMO.MOEndDate = Convert.ToDateTime(viewmodel.MTEndDate.ToString("yyyy/MM/dd") + " " + viewmodel.ehour + ":" + viewmodel.emin);

            modelMO.MDID = MDID;
            modelMO.MTID = MTID;
            MO_List.Add(modelMO);

            MeetingDevice MD = new MeetingDevice();
            if (MDID != 0)
            {
                MD.MDName = service.GetMDIDToShow(MDID);
            }

            service.CreateMeetingOrder(MO_List);

            //判斷是否使用留言通知
            var MsgService = ServiceHelper.Create<IMessageService>();
            if (viewmodel.MTMessageChk == true)
            {
                Message k = new Message();
                k.MSGSubject = "會議通知: " + modelMT.MTName;
                k.MSGDESC = "會議詳細內容連結:<a href='"+ PlatformHelper.GetVarConfig("MeetingDetailUrl") + MTID+"'>" + PlatformHelper.GetVarConfig("MeetingDetailUrl") +MTID+"</a>";
              

                k.MSGCreater = User.MemberInfo.ID;// (int)Session["orgID"];
                k.MSGCreateName = Microsoft.CUF.CodeName.GetUnitName(User.UnitInfo.ID).Trim() + "-" + Microsoft.CUF.CodeName.GetMemberName(User.MemberInfo.ID);
                k.MSGCreateIP = PlatformHelper.GetClientIPv4();
                k.MSGTime = DateTime.Now;
                k.MSGClass = 5;
                //產生留言人員名單
                List<MessageTo> kt_list = null;
                // 產生留言附加檔案資料
                List<MessageFile> kf_list = new List<MessageFile>();
                if (viewmodel.UploadFilesName != null && viewmodel.UploadFilesName.Length > 0)
                {
                    string fnames = viewmodel.UploadFilesName;

                    kf_list = (from item in fnames.Split('*')
                               let parts = item.Split('|')
                               select new MessageFile { MSGFileName = parts[0], MSGMD5Name = parts[1] }).ToList();
                }

                MessageTo partdata = new MessageTo();
                partdata.MSGOBJDate = k.MSGTime;
                partdata.MSGOBJReaderIP = "";
                partdata.MSGOBJSendIP = k.MSGCreateIP;
                partdata.MSGOBJCreateTime = k.MSGTime;
                partdata.MSGOBJSendID = User.MemberInfo.ID;// (int)Session["orgID"];

                kt_list = GetMsgIdsToMsgTList(idnameNewLst, partdata);
                MsgService.CreateMessage(k, kf_list, kt_list, viewmodel.TabUniqueId, out retMsg);
            }
            //TempData["message"] = "新增成功";
            AppendMessage(PlatformResources.新增成功, false);
            return RedirectToAction("index", "MMQU001");
        }
        #endregion

        #region 修改會議
        /// <summary>
        /// 修改頁面
        /// </summary>
        /// <param name="MTID">會議流水序號</param>
        /// <returns></returns>
        [HasPermission("EP.PSL.WorkResources.MeetingMng.MMTX001")]
        public ActionResult Edit(int MTID)
        {
            MyPageStatus = PageStatus.Edit;
            var mService = new WebChannel<IMeetingMngService>();
            MeetingDetailModel model = new MeetingDetailModel();
            Meeting m = new Meeting();
            List<MeetingFile> mf_lst = new List<MeetingFile>();
            string imember = User.MemberInfo.ID;

            mService.Use(service => m = service.GetMeetingDetailById(MTID, imember));
            model.MTID = m.MTID;
            model.MTName = m.MTName;
            model.MTEndDate = m.MTEndDate;
            model.MTDesc = System.Web.HttpUtility.HtmlDecode(m.MTDesc);
            model.MTStartDate = m.MTStartDate;
            model.shour = m.MTStartDate.ToString("HH");
            model.smin = m.MTStartDate.ToString("mm");
            model.ehour = m.MTEndDate.ToString("HH");
            model.emin = m.MTEndDate.ToString("mm");
            model.MTConvenerName = m.MTConvenerName;

            //處理主席資料轉JSON
            List<UserSimpleInfo> USChairman_list = new List<UserSimpleInfo>();
            mService.Use(service => service.GetEditMeetingChairmanById(MTID)
          .ForEach(d =>
          {
              if (d != null)
              {
                  var item = new UserSimpleInfo();
                  item.Name = d.nmember;
                  item.Unit = d.nunit;
                  item.AccountId = d.iaccount;
                  USChairman_list.Add(item);
              }
          }));
            List<AccountGroupJsonString> AGC = GetidNameToAGJSList(USChairman_list);
            DataContractJsonSerializer jsonC = new DataContractJsonSerializer(AGC.GetType());
            using (MemoryStream stream = new MemoryStream())
            {
                jsonC.WriteObject(stream, AGC);
                model.MTChairmanName = Encoding.UTF8.GetString(stream.ToArray());
            }
            //處理紀錄者資料轉JSON
            List<UserSimpleInfo> USRecorder_list = new List<UserSimpleInfo>();
            mService.Use(service => service.GetEditMeetingRecorderById(MTID)
         .ForEach(d =>
         {
             if (d != null)
             {
                 var item = new UserSimpleInfo();
                 item.Name = d.nmember;
                 item.Unit = d.nunit;
                 item.AccountId = d.iaccount;
                 USRecorder_list.Add(item);
             }
         }));
            List<AccountGroupJsonString> AGR = GetidNameToAGJSList(USRecorder_list);
            DataContractJsonSerializer jsonR = new DataContractJsonSerializer(AGR.GetType());
            using (MemoryStream stream = new MemoryStream())
            {
                jsonR.WriteObject(stream, AGR);
                model.MTRecorderName = Encoding.UTF8.GetString(stream.ToArray());
            }
            //處理成員資料轉JSON
            List<UserSimpleInfo> USParticipants_list = new List<UserSimpleInfo>();
            mService.Use(service => service.GetEditMeetingParticipantsById(MTID)
           .ForEach(d =>
           {
               if (d != null)
               {
                   var item = new UserSimpleInfo();
                   item.Name = d.nmember;
                   item.Unit = d.nunit;
                   item.AccountId = d.iaccount;
                   USParticipants_list.Add(item);
               }
           }));
            List<AccountGroupJsonString> AGP = GetidNameToAGJSList(USParticipants_list);
            DataContractJsonSerializer jsonP = new DataContractJsonSerializer(AGP.GetType());
            using (MemoryStream stream = new MemoryStream())
            {
                jsonP.WriteObject(stream, AGP);
                model.RecipientToJson = Encoding.UTF8.GetString(stream.ToArray());
            }

            //bool chk = _Service.ChkMeetingDevice(MTID);
            //ViewBag.chk = chk;
            MeetingDevice MD = _Service.GetEditMeetingDevice(MTID);
            if (MD != null)
            {
                model.MDID = MD.MDID;
                model.MDName = MD.MDName;
            }
            else
            {
                model.MTPlace = m.MTPlace;
                model.chkplace = true;
            }


            return View("Create", model);
        }

        /// <summary>
        /// 修改處理
        /// </summary>
        /// <param name="viewmodel"></param>
        /// <returns></returns>
        [HttpPost]
        [HasPermission("EP.PSL.WorkResources.MeetingMng.MMTX001")]
        public ActionResult Edit(MeetingDetailModel viewmodel)
        {
            MyPageStatus = PageStatus.Edit;
            string retMsg = string.Empty;
            Meeting modelMT = new Meeting();
            modelMT.MTID = viewmodel.MTID;
            //處理場地與設備MDID
            int MDID = Convert.ToInt32(viewmodel.MDName);
            //更改為自訂會議的話要刪掉設備中心相關的連動
            if (viewmodel.MTNewPlace != null)
            {
                modelMT.MTPlace = viewmodel.MTNewPlace;
                var result = _Service.DeleteMeetingOrder(modelMT.MTID);
                MDID = 0;
            }
            else
            {
                modelMT.MTPlace = "";
                MeetingDeviceRelation mdr = new MeetingDeviceRelation();
                List<MeetingDeviceRelation> MDR_list = new List<MeetingDeviceRelation>();
                mdr.MDID = MDID;
                mdr.MTID = viewmodel.MTID;
                MDR_list.Add(mdr);
                //修改會議設備
                _Service.EditMeetingDeviceRelation(MDR_list);

                //修改會議管理&設備中心預定的會議室和設備
                List<MeetingOrder> MO_List = new List<MeetingOrder>();

                MeetingOrder modelMO = new MeetingOrder();
                modelMO.imember = User.MemberInfo.ID;
                modelMO.iunit = User.UnitInfo.ID;
                modelMO.MOCheck = 1;
                modelMO.MOCreateDate = DateTime.Now;
                modelMO.MOTitle = "會議:" + viewmodel.MTName;
                modelMO.MODesc = modelMT.MTDesc + "(此資料由會議管理模組新增)";
                modelMO.MOStartDate = Convert.ToDateTime(viewmodel.MTStartDate.ToString("yyyy/MM/dd") + " " + viewmodel.shour + ":" + viewmodel.smin);
                modelMO.MOLendState = 1;
                modelMO.MOEndDate = Convert.ToDateTime(viewmodel.MTEndDate.ToString("yyyy/MM/dd") + " " + viewmodel.ehour + ":" + viewmodel.emin);
                modelMO.MDID = MDID;
                modelMO.MTID = viewmodel.MTID;
                MO_List.Add(modelMO);

                _Service.EditMeetingOrder(MO_List);
            }
            modelMT.MTName = viewmodel.MTName;
            //判斷地點是由下拉式選單新增或Textbox新增
            //if (viewmodel.MTPlace != null)
            //{
            //    modelMT.MTPlace = viewmodel.MTPlace;
            //}

            modelMT.MTDesc = System.Web.HttpUtility.HtmlDecode(viewmodel.MTDesc);
            modelMT.MTStartDate = Convert.ToDateTime(viewmodel.MTStartDate.ToString("yyyy/MM/dd") + " " + viewmodel.shour + ":" + viewmodel.smin);
            modelMT.MTEndDate = Convert.ToDateTime(viewmodel.MTEndDate.ToString("yyyy/MM/dd") + " " + viewmodel.ehour + ":" + viewmodel.emin);
            modelMT.MTCreatedate = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            modelMT.MTConvener = User.MemberInfo.ID;
            modelMT.MTActive = 1;
            //顯示部門 User.UnitInfo.GetParent().Name;
            modelMT.MTConvenerName = User.UnitInfo.Name + "-" + User.MemberInfo.Name;

            UserSimpleInfo UConvener = new UserSimpleInfo();

            //設定成員
            List<AccountGroupJsonString> JsonClassLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(viewmodel.RecipientToJson);
            List<UserSimpleInfo> idnameLst = PlatformHelper.GetRecipientUsersList(JsonClassLst);

            //設定主席
            List<AccountGroupJsonString> JsonChairmanLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(viewmodel.MTChairmanName);
            List<UserSimpleInfo> Chairman = PlatformHelper.GetRecipientUsersList(JsonChairmanLst);
            UserSimpleInfo UChairman = new UserSimpleInfo();
            for (int i = 0; i < Chairman.Count; i++)
            {
                if (Chairman.Count >= 2)
                {
                    TempData["message"] = "主席只能選一個";
                    //TempData["model"] = viewmodel;
                    return RedirectToAction("Edit", new { MTID = viewmodel.MTID });
                }
                modelMT.MTChairman = Chairman[i].MemberId;// Convert.ToInt32(idAry[i]);  
                modelMT.MTChairmanName = Chairman[i].Unit + "-" + Chairman[i].Name;
                UChairman.MemberId = Chairman[i].MemberId;
                UChairman.Unit = Chairman[i].Unit;
                UChairman.Name = Chairman[i].Name;
                idnameLst.Add(UChairman);
            }
            //設定紀錄者
            List<AccountGroupJsonString> JsonRecorderLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(viewmodel.MTRecorderName);
            List<UserSimpleInfo> Recorder = PlatformHelper.GetRecipientUsersList(JsonRecorderLst);
            UserSimpleInfo URecorder = new UserSimpleInfo();
            for (int i = 0; i < Recorder.Count; i++)
            {
                if (Recorder.Count >= 2)
                {
                    TempData["message"] = "紀錄者只能選一個";
                    //TempData["model"] = viewmodel;
                    return RedirectToAction("Edit", new { MTID = viewmodel.MTID });
                }
                modelMT.MTRecorder = Recorder[i].MemberId;// Convert.ToInt32(idAry[i]);  
                modelMT.MTRecorderName = Recorder[i].Unit + "-" + Recorder[i].Name;
                URecorder.MemberId = Recorder[i].MemberId;
                URecorder.Unit = Recorder[i].Unit;
                URecorder.Name = Recorder[i].Name;
                idnameLst.Add(URecorder);
            }

            //修改會議資訊
            _Service.EditMeeting(modelMT);

            //判斷召集人是否為成員
            if (viewmodel.Convenerchk == true)
            {
                UConvener.MemberId = User.MemberInfo.ID;
                UConvener.Unit = User.UnitInfo.Name;
                UConvener.Name = User.MemberInfo.Name;
                idnameLst.Add(UConvener);
            }

            List<MeetingMemberRelation> MMR_list = null;
            List<UserSimpleInfo> idnameNewLst = new List<UserSimpleInfo>();
            //去除重複人選   
            List<string> DistinctList = idnameLst.Select(x => x.MemberId).Distinct().ToList();
            for (int i = 0; i < DistinctList.Count; i++)
            {
                UserSimpleInfo user = new UserSimpleInfo();
                if (DistinctList[i] == idnameLst[i].MemberId)
                {
                    user.AccountId = idnameLst[i].AccountId;
                    user.MemberId = idnameLst[i].MemberId;
                    user.Name = idnameLst[i].Name;
                    user.Unit = idnameLst[i].Unit;
                    idnameNewLst.Add(user);
                }
            }

            MMR_list = GetIdsToList(idnameNewLst, viewmodel.MTID);
            //存入會議成員
            _Service.EditMeetingMemberRelation(MMR_list);

            //// 產生附加檔案資料
            List<MeetingFile> MF_list = new List<MeetingFile>();
            if (viewmodel.UploadFilesName != null && viewmodel.UploadFilesName.Length > 0)
            {
                string fnames = viewmodel.UploadFilesName;

                MF_list = (from item in fnames.Split('*')
                           let parts = item.Split('|')
                           select new MeetingFile { MFFileName = parts[0], MFMd5Name = parts[1] }).ToList();
                for (int i = 0; i < MF_list.Count; i++)
                {
                    MF_list[i].MTID = viewmodel.MTID;
                    MF_list[i].MFCreater = modelMT.MTConvener;
                    MF_list[i].MFCreateDate = modelMT.MTCreatedate;
                    MF_list[i].MFFilePath = DateTime.Now.ToString("yyyyMM") + @"\" + viewmodel.MTID;
                    MF_list[i].MFType = 1;
                }
                _Service.EditMeetingFile(modelMT, MF_list, viewmodel.TabUniqueId, out retMsg);
            }
            MeetingDevice MD = new MeetingDevice();
            if (MDID != 0)
            {
                MD.MDName = _Service.GetMDIDToShow(MDID);
            }
            //判斷是否使用留言通知
            var MsgService = ServiceHelper.Create<IMessageService>();
            if (viewmodel.MTMessageChk == true)
            {
                Message k = new Message();
                k.MSGSubject = "會議更動通知: " + modelMT.MTName;
                k.MSGDESC = "由於 會議: " + modelMT.MTName + " 更動，請點選以下連結查看" + " <br /> ";
                k.MSGDESC+="會議詳細內容連結:<a href='" + PlatformHelper.GetVarConfig("MeetingDetailUrl") + viewmodel.MTID + "'>" + PlatformHelper.GetVarConfig("MeetingDetailUrl") + viewmodel.MTID + "</a>";
                k.MSGCreater = User.MemberInfo.ID;// (int)Session["orgID"];
                k.MSGCreateName = Microsoft.CUF.CodeName.GetUnitName(User.UnitInfo.ID).Trim() + "-" + Microsoft.CUF.CodeName.GetMemberName(User.MemberInfo.ID);
                k.MSGCreateIP = PlatformHelper.GetClientIPv4();
                k.MSGTime = DateTime.Now;
                k.MSGClass = 5;
                //產生留言人員名單
                List<MessageTo> kt_list = null;
                //產生留言附加檔案資料
                List<MessageFile> kf_list = new List<MessageFile>();
                if (viewmodel.UploadFilesName != null && viewmodel.UploadFilesName.Length > 0)
                {
                    string fnames = viewmodel.UploadFilesName;

                    kf_list = (from item in fnames.Split('*')
                               let parts = item.Split('|')
                               select new MessageFile { MSGFileName = parts[0], MSGMD5Name = parts[1] }).ToList();
                }

                MessageTo partdata = new MessageTo();
                partdata.MSGOBJDate = k.MSGTime;
                partdata.MSGOBJReaderIP = "";
                partdata.MSGOBJSendIP = k.MSGCreateIP;
                partdata.MSGOBJCreateTime = k.MSGTime;
                partdata.MSGOBJSendID = User.MemberInfo.ID;// (int)Session["orgID"];

                kt_list = GetMsgIdsToMsgTList(idnameNewLst, partdata);
                MsgService.CreateMessage(k, kf_list, kt_list, viewmodel.TabUniqueId, out retMsg);
            }

            return RedirectToAction("Edit", new { MTID = viewmodel.MTID });
        }

        /// <summary>
        /// 取附件列表以供編輯
        /// </summary>
        /// <param name="MTID">會議編號</param>
        /// <returns></returns>
        public JsonResult GetFilesList(int MTID)
        {
            var urlBase = "/MeetingMng/MMTX001/ShowFile?file=";
            var deleteURL = "/JFileUpload/DeleteFile?file=";
            var filesHelper = new FilesHelper(deleteURL, "GET", null, urlBase, null, null);
            var fileList = new List<ViewDataUploadFilesResult>();
            List<MeetingFile> MeetingFileList = _Service.GetMeetingFileByMTId(MTID);

            if (MeetingFileList != null)
            {
                var actasdir = Path.Combine(GetMeetingMngDir());
                MeetingFileList.ForEach(item =>
                {
                    var filePath = Path.Combine(actasdir, item.MFMd5Name);
                    FileInfo file = new FileInfo(filePath);
                    int SizeInt = unchecked((int)file.Length);
                    var list = filesHelper.UploadResult(item.MFFileName, SizeInt, file.FullName, file.Name);
                    list.url = "/MeetingMng/MMTX001/DownFile?filename=" + item.MFFileName + "&filemd5=" + item.MFMd5Name;
                    fileList.Add(list);
                });
            }
            JsonFiles files = new JsonFiles(fileList);
            return Json(files, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 編輯會議時刪除檔案
        /// </summary>
        /// <param name="MTID">會議編號</param>
        /// <returns></returns>
        public void deletefile(string file)
        {
            MeetingFile mf = new MeetingFile();

            string[] sArray = file.Split('=');

            mf.MFMd5Name = sArray[1];
            _Service.DeleteEditMeetingFile(mf);
        }

        #endregion

        #region 刪除會議
        /// <summary>刪除會議處理</summary>
        /// <param name="MTID"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Delete(int MTID)
        {
            //執行刪除          
            var result = _Service.DeleteMeeting(MTID);
            //執行成功時
            if (result)
            {
                AppendMessage(PlatformResources.刪除成功, false);
            }
            else
            {
                AppendMessage(PlatformResources.刪除失敗, false);
            }
            //return RedirectToAction("index", "MMQU001");
            return View("Close");
        }
        #endregion

        #region 人員處理
        /// <summary>
        /// 組成會議成員清單 
        /// </summary>
        /// <param name="idNamelist">成員資料</param>
        /// <param name="MTID">會議編號</param>
        /// <returns></returns>
        private List<MeetingMemberRelation> GetIdsToList(List<UserSimpleInfo> idNamelist, int MTID)
        {
            List<MeetingMemberRelation> list = new List<MeetingMemberRelation>();

            for (int i = 0; i < idNamelist.Count; i++)
            {
                MeetingMemberRelation MMR = new MeetingMemberRelation();

                MMR.imember = idNamelist[i].MemberId;// Convert.ToInt32(idAry[i]);  
                MMR.MTParticipantsName = idNamelist[i].Unit + "-" + idNamelist[i].Name;
                MMR.MTReply = 0;
                MMR.MTID = MTID;
                list.Add(MMR);
            }
            return list;
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
                //收件名稱
                //kt.MSGOBJName = idNamelist[i].Name  ;
                kt.MSGOBJName = idNamelist[i].Unit + "-" + idNamelist[i].Name;
                list.Add(kt);

            }
            return list;
        }

        /// <summary>
        /// 成員清單轉為AccountGroupJsonString
        /// </summary>
        /// <param name="idNamelist">成員資料</param>
        /// <returns></returns>
        private List<AccountGroupJsonString> GetidNameToAGJSList(List<UserSimpleInfo> idNamelist)
        {
            List<AccountGroupJsonString> list = new List<AccountGroupJsonString>();

            for (int i = 0; i < idNamelist.Count; i++)
            {
                AccountGroupJsonString AG = new AccountGroupJsonString();
                AG.Text = idNamelist[i].Name;
                AG.Unit = idNamelist[i].Unit;
                AG.Key = idNamelist[i].AccountId;
                AG.Type = "ByAccount";
                list.Add(AG);
            }
            return list;
        }
        #endregion

        #region 場地與設備處理
        /// <summary>
        /// 場地與設備顯示處理
        /// </summary>
        /// <param name="MDIDList">場地與設備清單</param>
        /// <param name="MTID">會議編號</param>
        /// <returns></returns>
        public string GetMDIDToShow(int MTID)
        {
            MeetingDevice MD = _Service.GetEditMeetingDevice(MTID);
            if (MD != null)
            {
                return MD.MDID.ToString();
            }
            return "0";
        }


        /// <summary>
        /// 場地與設備顯示處理
        /// </summary>
        /// <param name="MDIDList">場地與設備清單</param>
        /// <param name="MTID">會議編號</param>
        /// <returns></returns>
        //public string GetMDIDToShow(string MDID)
        //{
        //    var service = ServiceHelper.Create<IMeetingMngService>();
        //    List<int> MDIDList = JsonConvert.DeserializeObject<List<int>>(MDID);
        //    string result = service.GetMDIDToShow(MDIDList);

        //    return result;
        //}

        /// <summary>
        /// 場地與設備處理
        /// </summary>
        /// <param name="MDIDList">場地與設備清單</param>
        /// <param name="MTID">會議編號</param>
        /// <returns></returns>
        //private List<MeetingDeviceRelation> GetDeviceToList(List<int> MDIDList, int MTID)
        //{
        //    List<MeetingDeviceRelation> list = new List<MeetingDeviceRelation>();

        //    for (int i = 0; i < MDIDList.Count; i++)
        //    {
        //        MeetingDeviceRelation MDR = new MeetingDeviceRelation();

        //        MDR.MDID = MDIDList[i];
        //        MDR.MTID = MTID;
        //        list.Add(MDR);
        //    }
        //    return list;
        //}
        #endregion

        #region 會議地點重覆時間判斷
        /// <summary>
        ///  新增會議地點重覆時間判斷
        /// </summary>
        /// <param name="MDID">場地與設備編號</param>
        /// <param name="sdate">開始日期</param>
        /// <param name="edate">結束日期</param>
        /// <returns></returns>
        public ActionResult GetMDIDToChk(string MDID, string sdate, string edate)
        {
            var service = ServiceHelper.Create<IMeetingMngService>();
            //List<int> MDIDList = JsonConvert.DeserializeObject<List<int>>(MDID);
            int iMDID = Convert.ToInt32(MDID);
            MeetingDetail timechk = new MeetingDetail();

            DateTime s = Convert.ToDateTime(sdate);
            //bool result = true;
            if (iMDID != 0)
            {
                timechk = service.ChkMeetingTime(iMDID, Convert.ToDateTime(sdate), Convert.ToDateTime(edate));
                //if (timechk != null)
                //{
                //    result = false;
                //}
            }
            if (timechk == null || iMDID == 0)
            {
                int a = 0;
                return Json(a, JsonRequestBehavior.AllowGet);
            }
            int havetime = 1;
            return Json(havetime, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        ///  修改會議地點重覆時間判斷
        /// </summary>
        /// <param name="MDID">場地與設備編號</param>
        /// <param name="sdate">開始日期</param>
        /// <param name="edate">結束日期</param>
        /// <returns></returns>
        public ActionResult GetEditMDIDToChk(string MDID, DateTime sdate, DateTime edate, int MTID)
        {
            Meeting modelMT = new Meeting();
            var service = ServiceHelper.Create<IMeetingMngService>();
            modelMT = service.GetMeeting(MTID);
            int iMDID = Convert.ToInt32(MDID);
            int oldMDID = 0;
            List<MeetingDeviceRelation> modelMDRlist = service.GetEditMeetingDeviceRelation(MTID);
            //view傳來的MDID
            //List<int> MDIDList = JsonConvert.DeserializeObject<List<int>>(MDID);
            //資料庫傳來的MDID
            //var MDIDListchk = new List<int>();
            for (int i = 0; i < modelMDRlist.Count; i++)
            {
                oldMDID = modelMDRlist[i].MDID;
            }

            MeetingDetail timechk = new MeetingDetail();
            //bool result = true;

            //假如沒有改時間
            if ((modelMT.MTStartDate == sdate && modelMT.MTEndDate == edate) && iMDID == oldMDID)
            {
                //只要新選設備與場地的值
                //var expectedMDRList = MDIDList.Except(MDIDListchk).ToList();
                //判斷有沒有更改設備與場地 true沒改 false有改
                //bool MDIDupdatechk = MDIDListchk.OrderBy(x => x).SequenceEqual(MDIDList.OrderBy(x => x));

                //if (MDIDupdatechk == false)
                //{
                //    //if (iMDID != null)
                //    //{
                //    //    //timechk = service.ChkMeetingTime(expectedMDRList, sdate, edate);
                //    //    //if (timechk != null)
                //    //    //{
                //    //    //    result = false;
                //    //    //}
                //    //}
                //}
            }
            else if (iMDID != oldMDID)
            {
                timechk = service.ChkMeetingTime(iMDID, sdate, edate);
            }
            else
            {
                if (iMDID != 0)
                {
                    //timechk = service.ChkEditMeetingTime(iMDID, sdate, edate, MTID);
                    timechk = service.ChkMeetingTime(iMDID, sdate, edate);
                    //if (timechk != null)
                    //{
                    //    result = false;
                    //}
                }
            }
            if (timechk == null || iMDID == 0)
            {
                int a = 0;
                return Json(a, JsonRequestBehavior.AllowGet);
            }
            int havetime = 1;
            return Json(havetime, JsonRequestBehavior.AllowGet);
        }


        #endregion

        #region 顯示會議在Ckeditor上傳的圖片顯示處理 ShowFile
        /// <summary>
        /// 顯示公告在Ckeditor上傳的圖片顯示處理
        /// </summary>
        /// <param name="url">路徑</param>
        /// <param name="file">檔案名稱</param>
        /// <param name="type">檔案類型</param>
        /// <returns></returns>
        //public ActionResult ShowFile(string url, string file, string type)
        //{

        //    string[] paths = new string[] { GetMeetingMngDir(), HttpUtility.HtmlDecode(url) ?? "", HttpUtility.HtmlDecode(file) };
        //    string filePath = Path.Combine(paths);
        //    FileInfo fileInfo = new FileInfo(filePath);
        //    FileStream stream = new FileStream(filePath,
        //                  FileMode.Open, FileAccess.Read);
        //    return File(stream, HttpUtility.HtmlDecode(type) ?? MimeMapping.GetMimeMapping(filePath));
        //}
        #endregion

        #region 取得會議檔案路徑 GetMeetingMngDir
        /// <summary>
        /// 取得會議檔案路徑
        /// </summary>
        /// <returns></returns>
        private static string GetMeetingMngDir()
        {
            var result = PlatformHelper.GetVarConfig("MeetingMngDir");
            return result;
        }
        #endregion

        #region 取得暫存的會議檔案路徑 GetTempDir
        /// <summary>
        /// 取得暫存的會議檔案路徑
        /// </summary>
        /// <returns></returns>
        private static string GetTempDir()
        {
            var result = PlatformHelper.GetVarConfig("TempDir");
            return result;
        }
        #endregion

    }
}