using EP.H2OModels;
using EP.Platform.Service;
using EP.PSL.WorkResources.MeetingMng.Service;
using EP.PSL.WorkResources.MeetingMng.Web.Areas.MeetingMng.Model;
using EP.Web;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;

namespace EP.PSL.WorkResources.MeetingMng.Web.Areas.MeetingMng.Controllers
{
    public class MMTX002Controller : BaseController
    {
        // GET: MeetingMng/MMTX002
        private IMeetingMngService _Service;
        public MMTX002Controller()
        {
            _Service = ServiceHelper.Create<IMeetingMngService>();
        }
        #region 會議詳細資料


        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 詳細資料
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HasPermission("EP.PSL.WorkResources.MeetingMng.MMTX002")]
        public ActionResult Detail(int id)
        {
            //頁面狀態
            MyPageStatus = PageStatus.View;

            var mService = new WebChannel<IMeetingMngService>();
            MeetingDetailModel model = new MeetingDetailModel();
            string MDNamesList = "";
            string MFFileNameList = "";
            //string ParticipantsList = "";
            //string ParticipateList = "";
            //string NoParticipateList = "";
            //string NoReplyList = "";
            Meeting m = new Meeting();
            MeetingMemberRelation MMR = new MeetingMemberRelation();
            List<MeetingFile> mf_lst = new List<MeetingFile>();
            string imember = User.MemberInfo.ID;

            mService.Use(service => m = service.GetMeetingDetailById(id, imember));
            model.MTID = id;
            model.MTName = m.MTName;
            model.MTPlace = m.MTPlace;
            model.MTChairmanName = m.MTChairmanName;
            model.MTConvenerName = m.MTConvenerName;
            model.MTConvener = m.MTConvener;
            model.MTRecorderName = m.MTRecorderName;
            model.MTEndDate = m.MTEndDate;
            model.MTDesc = System.Web.HttpUtility.HtmlDecode(m.MTDesc);
            model.MTStartDate = m.MTStartDate;
            model.MTCreatedate = m.MTCreatedate;
            model.MTActive = m.MTActive;
            model.JBStartDate = DateTime.Today;
            model.JBEndDate = DateTime.Today;
            if (model.MTActive == 1)
            {
                //參加狀態顯示
                mService.Use(service => MMR = service.GetMeetingDetailMTReplyById(id, imember));
                //召集人不參加 MMR就會是null
                if (MMR != null)
                {
                    model.MTReply = MMR.MTReply;
                }
                else
                {
                    //MTReply給 3 召集人不顯示參加狀態
                    model.MTReply = 3;
                }
            }
            //檔案顯示
            mService.Use(service => service.GetMeetingDetailFilesById(id)
            .ForEach(d =>
            {
                if (d != null)
                {
                    MFFileNameList += d.MFFileName + "/";
                }
            }));
            model.MFFileName = MFFileNameList;
            //設備與場地顯示
            MeetingDevice md = _Service.GetEditMeetingDevice(id);
            if (model.MTPlace == "") 
            {
                model.MTPlace = md.MDName;
            }
           

            #region 與會人員處理，已改寫至GetMeetingParticipantsByID
            //與會人員顯示
            //mService.Use(service => service.GetMeetingDetailParticipantsById(id)
            //.ForEach(d =>
            //{
            //    if (d != null)
            //    {
            //        ParticipantsList += d.MTParticipantsName + "/";
            //    }
            //}));

            //model.Participants = ParticipantsList;
            //已回覆參加人員顯示
            //mService.Use(service => service.GetMeetingDetailParticipateById(id)
            //.ForEach(d =>
            //{
            //    if (d != null)
            //    {
            //        ParticipateList += d.MTParticipantsName + "/";
            //    }
            //}));

            //model.Participate = ParticipateList;
            //已回覆不參加人員顯示
            // mService.Use(service => service.GetMeetingDetailNoParticipateById(id)
            //.ForEach(d =>
            //{
            //    if (d != null)
            //    {
            //        NoParticipateList += d.MTParticipantsName + "/";
            //    }
            //}));

            // model.NoParticipate = NoParticipateList;
            //未回覆人員顯示
            // mService.Use(service => service.GetMeetingDetailNoReplyById(id)
            //.ForEach(d =>
            //{
            //    if (d != null)
            //    {
            //        NoReplyList += d.MTParticipantsName + "/";
            //    }
            //}));

            // model.NoReply = NoReplyList;
            #endregion

            mService.Use(service => mf_lst = service.GetMeetingFileByMTId(id));
            //取得附件清單
            model.File = mf_lst.Select(x => new MeetingFileModel
            {
                MFID = x.MFID,
                MFFileName = x.MFFileName,
                MTID = x.MTID,
                MFMd5Name = x.MFMd5Name
            }).ToList();

            return View(model);
        }

        /// <summary>
        /// 人員是否參加會議
        /// </summary>
        /// <param name="MTID"></param>
        /// <param name="MTReply">參加狀態</param>
        /// <returns></returns>
        public void CheckReply(int MTID, int MTReply)
        {
            var service = ServiceHelper.Create<IMeetingMngService>();
            string imember = User.MemberInfo.ID;
            service.UpdateMeetingReplyById(MTID, MTReply, imember);
        }

        /// <summary>
        /// 將會議存放到歷史資料
        /// </summary>
        /// <param name="MTID"></param>
        /// <param name="MTActive">會議狀態</param>
        /// <returns></returns>
        public void UpdateMTActive(int MTID, int MTActive)
        {
            var service = ServiceHelper.Create<IMeetingMngService>();
            service.UpdateMeetingActiveById(MTID, MTActive);
        }

        #endregion

        #region 相關檔案
        /// <summary>
        /// 相關檔案
        /// </summary>
        /// <param name="cond"></param>
        /// <returns></returns>
        public void Query(QueryMeetingFilesCondition cond)
        {
            List<MeetingDetailModel> QResultList = new List<MeetingDetailModel>();
            var mService = new WebChannel<IMeetingMngService>();
            mService.Use(service => service.GetMeetingRelatedFilesById(cond)
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new MeetingDetailModel();
                    item.MFID = d.MFID;
                    item.MTID = d.MTID;
                    item.nunit = d.nunit;
                    item.MFFileName = d.MFFileName;
                    string[] MFMd5Name = d.MFMd5Name.Split('\\');
                    item.MFMd5Name = MFMd5Name[1];
                    item.MFCreater = d.MFCreater;
                    item.nmember = d.nmember;
                    item.MFCreateDate = d.MFCreateDate;
                    item.MFDesc = d.MFDesc;
                    QResultList.Add(item);
                }
            }));
            var gridKey = mService.DataToCache(QResultList.AsEnumerable());
            SetGridKey("MMFileGrid", gridKey);
        }

        /// <summary>
        /// 查詢結果資料綁定jquery處理
        /// </summary>
        /// <param name="jqParams"></param>
        /// <returns></returns>
        public JsonResult BindGrid(jqGridParam jqParams)
        {
            //取得CacheKey
            var key = GetGridKey("MMFileGrid");
            return BaseGridBinding<MeetingDetailModel>(jqParams,
                () => new WebChannel<IMeetingMngService, MeetingDetailModel>().Get(key));
        }

        /// <summary>
        /// 新增相關檔案 (GET)
        /// </summary>
        /// <returns></returns>
        public ActionResult CreateFile()
        {
            MeetingDetailModel m = new MeetingDetailModel();
            m.nunit = User.UnitInfo.Name + "-";
            m.nmember = User.MemberInfo.Name;
            return PartialView("CreateFile", m);
        }
        /// <summary>
        /// 新增相關檔案處理
        /// </summary>
        /// <param name="viewmodel"></param>
        /// <returns></returns>
        [HttpPost]
        //[ValidateInput(false)]
        public ActionResult CreateFile(MeetingDetailModel viewmodel)
        {
            string retMsg = string.Empty;
            if (viewmodel.UploadFilesName == null && viewmodel.UploadFilesName.Length < 0)
                Throw.BusinessError(EP.PlatformResources.請傳入指定參數資料);
            //初始宣告:當發生Error時頁面才可以保留原始資訊
            //this.DefaultView.ViewName = "MessageDetail";
            //this.TempData["model"] = model;
            //MyPageStatus = PageStatus.Create;
            //if (!ModelState.IsValid)
            //{
            //Throw.BusinessError("檢核失敗");
            //}
            Meeting modelMT = new Meeting();
            modelMT.MTID = viewmodel.MTID;
            modelMT.MTConvener = User.MemberInfo.ID;
            MeetingFile model = new MeetingFile();
            string time = DateTime.Now.ToString("yyyy/MM/dd tt hh:mm:ss", CultureInfo.InstalledUICulture);

            //產生附加檔案資料
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
                    MF_list[i].MFDesc = System.Web.HttpUtility.HtmlDecode(viewmodel.MFDesc);
                    MF_list[i].MFCreater = User.MemberInfo.ID;
                    MF_list[i].MFCreateDate = time;
                    MF_list[i].MFFilePath = DateTime.Now.ToString("yyyyMM") + @"\" + viewmodel.MTID;
                    MF_list[i].MFType = 2;
                }
            }
            _Service.CreateMeetingFile(modelMT, MF_list, viewmodel.TabUniqueId, out retMsg);

            return RedirectToAction("Detail", new { id = modelMT.MTID });
        }

        /// <summary>
        /// 取得附件資料傳給JQuery FileUpload
        /// </summary>
        /// <param name="bnid"></param>
        /// <returns></returns>
        public JsonResult GetTempFilesList(string UploadFilesName)
        {
            var urlBase = "/JFileUpload/Show?file=";
            var deleteURL = "/JFileUpload/DeleteFile?file=";
            //
            var filesHelper = new FilesHelper(deleteURL, "GET", null, urlBase, null, null);

            var fileList = new List<ViewDataUploadFilesResult>();

            String fullPath = Path.Combine(PlatformHelper.GetVarConfig("TempDir"), User.MemberInfo.ID, TabUniqueId);
            if (Directory.Exists(fullPath))
            {
                UploadFilesName.Split('*').Where(filename => !string.IsNullOrWhiteSpace(filename)).ForEach(filename =>
                {

                    // string[0]為使者用的檔名 string[1]為在server上已經異動為guid後的檔名
                    var filenameInfo = filename.Split('|');
                    var filePath = Path.Combine(fullPath, filenameInfo[1]);
                    if (System.IO.File.Exists(filePath))
                    {
                        FileInfo file = new FileInfo(filePath);
                        int SizeInt = unchecked((int)file.Length);
                        fileList.Add(filesHelper.UploadResult(filenameInfo[0], SizeInt, file.FullName, file.Name));
                    }
                });

            }
            EP.Web.JsonFiles files = new EP.Web.JsonFiles(fileList);

            return Json(files, JsonRequestBehavior.AllowGet);

        }

        /// <summary>刪除檔案資料處理</summary>
        /// <param name="MFID"></param>
        /// <param name="MTID"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult DeleteFile(int MFID, int MTID)
        {
            //執行刪除          
            var result = _Service.DeleteMeetingFile(MFID);
            //執行成功時
            if (result)
            {
                AppendMessage(PlatformResources.刪除成功, false);
            }
            else
            {
                AppendMessage(PlatformResources.刪除失敗, false);
            }
            return View("Close");
        }
        #endregion

        #region 決議事項
        /// <summary>
        /// 決議事項
        /// </summary>
        /// <param name="cond"></param>
        /// <returns></returns>
        public void QueryJob(QueryMeetingJobCondition cond)
        {
            List<JobDetailModel> QResultList = new List<JobDetailModel>();
            var mService = new WebChannel<IMeetingMngService>();
            mService.Use(service => service.GetMeetingJobById(cond)
            .ForEach(d =>
            {
                if (d != null)
                {
                    var item = new JobDetailModel();
                    item.JPID = d.JPID;
                    item.MTID = d.MTID;
                    item.JBID = d.JBID;
                    item.JBStartDate = d.JBStartDate;
                    item.JBEndDate = d.JBEndDate;
                    item.imember = d.imember;
                    item.JBUnitName = d.JBUnitName;
                    item.JPPercentage = d.JPPercentage;
                    item.JBSubject = d.JBSubject;
                    QResultList.Add(item);
                }
            }));
            var gridKey = mService.DataToCache(QResultList.AsEnumerable());
            SetGridKey("MMTX002GridJob", gridKey);
        }

        /// <summary>
        /// 查詢結果資料綁定jquery處理
        /// </summary>
        /// <param name="jqParams"></param>
        /// <returns></returns>
        public JsonResult BindJobGrid(jqGridParam jqParams)
        {
            //取得CacheKey
            var key = GetGridKey("MMTX002GridJob");
            return BaseGridBinding<JobDetailModel>(jqParams,
                () => new WebChannel<IMeetingMngService, JobDetailModel>().Get(key));
        }

        /// <summary>
        /// 新增決議事項 (GET)
        /// </summary>
        //public ActionResult CreateJob(int MTID)
        //{
        //    MeetingDetailModel model = new MeetingDetailModel();
        //    model.JBStartDate = DateTime.Today;
        //    model.JBEndDate = DateTime.Today;
        //    model.MTID = MTID;
        //    return PartialView("CreateJob", model);
        //}

        /// <summary>
        /// 新增決議事項處理
        /// </summary>
        /// <param name="viewmodel"></param>
        [HttpPost]
        //[ValidateInput(false)]
        public ActionResult CreateJob(MeetingDetailModel viewmodel)
        {
            //if (viewmodel == null)
            //    Throw.BusinessError(EP.PlatformResources.請傳入指定參數資料);

            //初始宣告:當發生Error時頁面才可以保留原始資訊
            //this.DefaultView.ViewName = "MessageDetail";
            //this.TempData["model"] = model;
            //MyPageStatus = PageStatus.Create;
            //if (!ModelState.IsValid)
            //{
            //Throw.BusinessError("檢核失敗");
            //}

            string hostName = Dns.GetHostName();
            string getIP = "";

            IPHostEntry ipHostEntry = Dns.GetHostEntry(hostName);
            foreach (IPAddress ipAddress in ipHostEntry.AddressList)
            {
                getIP = Convert.ToString(ipAddress);
            }

            var service = ServiceHelper.Create<IMeetingMngService>();
            Job modelJob = new Job();
            modelJob.JBDesc = System.Web.HttpUtility.HtmlDecode(modelJob.JBDesc);
            modelJob.JBClass = 1;
            modelJob.JBCount = viewmodel.MTID;
            modelJob.JBCreateDate = DateTime.Now;
            modelJob.JBCreateIP = getIP;
            modelJob.JBCreater = User.MemberInfo.ID;
            modelJob.JBDesc = viewmodel.JBDesc;
            modelJob.JBStartDate = viewmodel.JBStartDate;
            modelJob.JBEndDate = viewmodel.JBEndDate;
            modelJob.JBSubject = viewmodel.JBSubject;
            modelJob.JBPrevid = 0;
            int JBID = service.CreateJob(modelJob);

            MeetingJobRelation modelMJR = new MeetingJobRelation();
            modelMJR.imember = User.MemberInfo.ID;
            modelMJR.JBID = JBID;
            modelMJR.MTID = viewmodel.MTID;
            modelMJR.JBType = 2;
            service.CreateMeetingJobRelation(modelMJR);

            List<JobPercentage> JP_list = null;
            List<AccountGroupJsonString> JsonClassLst = JsonConvert.DeserializeObject<List<AccountGroupJsonString>>(viewmodel.RecipientToJson);
            List<UserSimpleInfo> idnameLst = PlatformHelper.GetRecipientUsersList(JsonClassLst);
            JobPercentage modelJP = new JobPercentage();
            modelJP.JPEndDate = viewmodel.JBEndDate;
            modelJP.JBID = JBID;

            JP_list = GetIdsToList(idnameLst, modelJP);

            service.CreateJobPercentage(JP_list);

            return RedirectToAction("Detail", new { id = viewmodel.MTID });
        }

        /// <summary>
        /// 追蹤決議事項詳細資料
        /// </summary>
        /// <param name="MTID">會議流水號</param>
        /// <param name="JBID">追蹤決議事項流水號</param>
        /// <param name="imember">決議事項成員</param>
        /// <returns></returns>
        public ActionResult DetailJob(int JPID)
        {
            var mService = new WebChannel<IMeetingMngService>();
            JobDetail m = new JobDetail();
            JobDetailModel model = new JobDetailModel();
            mService.Use(service => m = service.GetJobDetailById(JPID));
            model.JBDesc = m.JBDesc;
            model.Creater = m.Creater;
            model.MTID = m.MTID;
            model.JBID = m.JBID;
            model.JPID = m.JPID;
            model.JBSubject = m.JBSubject;
            model.JBUnitName = m.JBUnitName;
            model.JBStartDate = m.JBStartDate;
            model.JBEndDate = m.JBEndDate;
            model.JPPercentage = m.JPPercentage;
            model.CreaterNunit = m.CreaterNunit;
            model.Creater = m.Creater;
            model.JBCreateIP = m.JBCreateIP;
            model.JBCreater = m.JBCreater;
            model.JBCreateDate = m.JBCreateDate;


            return View(model);
        }

        /// <summary>刪除決議事項資料處理</summary>
        /// <param name="JBID"></param>
        /// <param name="JPID"></param>
        /// <param name="MTID"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult DeleteJob(int JBID, int JPID, int MTID)
        {
            //執行刪除          
            var result = _Service.DeleteJob(JBID, JPID);
            //執行成功時
            if (result)
            {
                AppendMessage(PlatformResources.刪除成功, false);
            }
            else
            {
                AppendMessage(PlatformResources.刪除失敗, false);
            }
            return RedirectToAction("Detail", new { id = MTID });
        }
        #endregion

        #region 下載檔案
        /// <summary>
        /// 下載檔案
        /// </summary>
        /// <param name="fid"></param>
        /// <returns></returns>
        public ActionResult GetMeetingFile(int fid)
        {
            RemoteFileInfo rf = null;
            var mSevice = ServiceHelper.Create<IMeetingMngService>();
            EP.H2OModels.MeetingFile kmf = mSevice.GetMeetingFileByMfid(fid);
            var streamservice = ServiceHelper.Create<IStreamMediaService>();
            DownloadRequest dr = new DownloadRequest();
            dr.FileName = kmf.MFMd5Name;
            //var mimeType = "application/octet-stream";
            try
            {
                rf = streamservice.DownloadMeetingFile(dr);
            }
            catch
            {
                //return Content(ee.Message );
                return Content("Get File fail");
                //return Content("<script language='javascript' type='text/javascript'>alert('download file fail!');</script>"); 
            }
            return File(rf.FileByteStream, rf.MimeType, kmf.MFFileName);
        }
        #endregion

        #region 取得會議檔案路徑 GetBulletinDir
        /// <summary>
        /// 取得會議檔案路徑
        /// </summary>
        /// <returns></returns>
        private static string GetBulletinDir()
        {
            var result = PlatformHelper.GetVarConfig("MeetingMngDir");
            return result;
        }
        #endregion

        #region 取得暫存路徑 GetTempDir
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

        #region 成員處理
        /// <summary>
        /// 組成決議事項成員清單 
        /// </summary>
        /// <param name="idNamelist">接收者資料</param>
        /// <param name="modelJP">追蹤事項進度表</param>
        /// <returns></returns>
        private List<JobPercentage> GetIdsToList(List<UserSimpleInfo> idNamelist, JobPercentage modelJP)
        {
            List<JobPercentage> list = new List<JobPercentage>();

            for (int i = 0; i < idNamelist.Count; i++)
            {
                JobPercentage JP = new JobPercentage();

                JP.imember = idNamelist[i].MemberId;// Convert.ToInt32(idAry[i]);  
                JP.JBUnitName = idNamelist[i].Unit + "-" + idNamelist[i].Name;
                JP.JPEndDate = modelJP.JPEndDate;
                JP.JBID = modelJP.JBID;
                list.Add(JP);
            }
            return list;
        }
        #endregion

        #region 與會人員明細
        /// <summary>
        /// 與會人員明細
        /// </summary>
        /// <param name="id">會議流水號</param>
        /// <returns></returns>
        [OutputCache(NoStore = true, Location = System.Web.UI.OutputCacheLocation.Client, Duration = 2)]
        public JsonResult GetMeetingParticipantsByID(int id)
        {
            List<string> result = new List<string>();
            MeetingDetailModel MD = new MeetingDetailModel();
            string Participants;
            var mService = new WebChannel<IMeetingMngService>();
            mService.Use(service => service.GetMeetingDetailParticipantsById(id)
          .ForEach(d =>
          {
              if (d != null)
              {
                  if (d.MTReply == 0)
                  {
                      Participants = d.MTParticipantsName + "(未回覆)";
                  }
                  else if (d.MTReply == 1)
                  {
                      Participants = d.MTParticipantsName + "(要參加)";
                  }
                  else
                  {
                      Participants = d.MTParticipantsName + "(不參加)";
                  }
                  result.Add(Participants);
              }
          }));
            return Json(result, JsonRequestBehavior.AllowGet);
        }
        #endregion
    }
}