using EP.Platform.Service;
using EP.SD.SalesSupport.CUSCRM.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Caching;
using EP.Web;
using System.Web.Mvc;
using static EP.SD.SalesSupport.CUSCRM.HistoryMaintainViewModel;

namespace EP.SD.SalesSupport.CUSCRM.Web.Areas.CUSCRM.Controllers
{
    /// <summary>
    /// 客服業務系統歷史查詢(H2O)
    /// </summary>
    [Program("CUSCRMQU001")]
    public class CUSCRMQU001Controller : BaseController
    {
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMQU001")]
        public ActionResult Index()
        {
            var condition = new HistoryQueryCondition();
            var now = DateTime.Now;

            return View(condition);
        }
        /// <summary>
        /// 維護頁面
        /// </summary>
        /// <param name="crm_no"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMQU001")]
        public ActionResult Maintain(string crm_no)
        {
            var _mService = ServiceHelper.Create<IQueryService>();
            var recordresult = _mService.QueryHistoryMaintainRecordList(crm_no).ToList();
            var fileresult = _mService.QueryHistoryMaintainFileList(crm_no).ToList();
            var now = DateTime.Now;
            var model = new HistoryMaintainViewModel();
            model.crm_do_createname = Member.Get(User.MemberInfo.ID).GetUnit().GetParent().Name + " " + User.MemberInfo.Name;
            model.crm_no = crm_no;
            model.crm_do_createdate = GetChineseTimeFormat(now);
            model.crm_do_time = now.ToString("t");
            model.maintainlist = new List<RecordViewModel>();
            if (recordresult.Count > 0)
            {
                
                foreach (var record in recordresult)
                {
                    model.maintainlist.Add(
                        new HistoryMaintainViewModel.RecordViewModel
                        {
                            crm_no=record.crm_no,
                            crm_do = record.crm_do,
                            crm_do_createdate = GetChineseTimeFormat(record.crm_do_createdate),
                            crm_do_time = record.crm_do_createdate.ToString("t"),
                            crm_do_createname = record.crm_do_createname,
                        });
                }
            }
            model.filelist = new List<Platform.Service.ValueText>();
            if (fileresult.Count > 0)
            {
               
                foreach (var file in fileresult)
                {
                    model.filelist.Add(new Platform.Service.ValueText
                    {
                        Value = file.crm_md5name,
                        Text = file.crm_filename,

                    });
                }
            }
            return View(model);
        }
        /// <summary>
        /// 查詢歷史資料報表
        /// </summary>
        /// <param name="condition"></param>
        [PdLogFilter("EIP客服-歷史查詢", PITraceType.Query)]
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMQU001")]
        public void QueryHistoryCSGridDatas(HistoryQueryCondition condition)
        {
            var gridList = Enumerable.Empty<HistoryCSViewModel>().ToList();
            var channel = new WebChannel<IQueryService>();
            channel.Use(service => service
            .QueryHistoryCSGridDatas(condition).ToList()
            .ForEach(d =>
            {
                if (d != null)
                {
                    var model = d;
                    model.MaintainBox = "<INPUT class='btn btn-default' style ='margin: 5px;font-size: 0.9em;' onclick=getedit('" + model.crm_no + "'); type=button value=維護 >";
                    model.UrgeBox = "<INPUT class='btn btn-default' style ='margin: 5px;font-size: 0.9em;' onclick=getadd('" + model.crm_no + "'); type=button value=稽催 >";
                    model.ExportBox = "<INPUT class='btn btn-default' style ='margin: 5px;font-size: 0.9em;' onclick=get_crm_doc('" + model.crm_no + "'); type=button value=匯出 >";
                    gridList.Add(model);
                }
            }));
            var gridKey = channel.DataToCache(gridList);
            SetGridKey("BindHistoryCSGridDatas", gridKey);
        }
        [HttpPost]
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMQU001")]

        public JsonResult BindHistoryCSGridDatas(jqGridParam jqParams)
        {
            //取得CacheKey
            var key = GetGridKey("BindHistoryCSGridDatas");
            return BaseGridBinding<HistoryCSViewModel>(jqParams, () => new WebChannel<IQueryService, HistoryCSViewModel>().Get(key));
        }
        /// <summary>
        /// 新增維護紀錄
        /// </summary>
        /// <returns></returns>
        [PdLogFilter("EIP客服-維護紀錄", PITraceType.Insert)]
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMQU001")]
        [HttpPost]
        public void CreateMaintainRecord([System.Web.Http.FromBody]string crm_no, [System.Web.Http.FromBody] string crm_do)
        {
            if (String.IsNullOrEmpty(crm_do))
                Throw.BusinessError("摘要不得為空");
            var model = new tcrm_do();
            model.crm_do_createdate = DateTime.Now;
            model.crm_do_createid = System.Web.HttpContext.Current.Session["orgID"].ToString();
            model.crm_no = crm_no;
            model.crm_do = crm_do;
            model.crm_do_createname = Member.Get(User.MemberInfo.ID).GetUnit().GetParent().Name + " " + User.MemberInfo.Name;
            var mService = ServiceHelper.Create<IQueryService>();
            mService.CreateCrm_do(model);
            foreach (string key in Request.Files.AllKeys)
            {
                var httpPostedFile = Request.Files[key];
                if (httpPostedFile != null && httpPostedFile.ContentLength != 0)
                {
                    var guidFileName = Guid.NewGuid().ToString("N") + Path.GetExtension(httpPostedFile.FileName);
                    //產存放檔案路徑
                    var dsService = ServiceHelper.Create<IDataSettingService>();
                    var saveDirRoot = dsService.GetConfigValueByName("CRM歷史檔案路徑");
                    //創建資料夾  
                    if (!Directory.Exists(saveDirRoot))
                    {
                        Directory.CreateDirectory(saveDirRoot);
                    }
                    var newPath = Path.Combine(saveDirRoot, Path.GetFileName(guidFileName));
                    httpPostedFile.SaveAs(newPath); //存放檔案到伺服器上
                    var fileModel = new crm_do_file();
                    fileModel.crm_no = crm_no;
                    fileModel.crm_md5name = guidFileName;
                    fileModel.crm_filename = httpPostedFile.FileName;
                    mService.CreateCrm_do_file(fileModel);
                }
            }
        }
        /// <summary>
        /// 結案
        /// </summary>
        /// <param name="crm_no"></param>
        /// <param name="crm_do"></param>
        /// <param name="closestatus"></param>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMQU001")]
        [HttpPost]
        [PdLogFilter("EIP客服-結案Log", PITraceType.Insert)]
        public void CreateCloseRecord([System.Web.Http.FromBody] string crm_no, [System.Web.Http.FromBody] string crm_do, [System.Web.Http.FromBody] string closestatus)
        {
            if (String.IsNullOrEmpty(crm_do))
                Throw.BusinessError("摘要不得為空");
            var model = new crm_close_log();
            model.crm_no = crm_no;
            model.creater_orgid = Convert.ToInt32(System.Web.HttpContext.Current.Session["orgID"]);
            model.old_crm_close = closestatus;
            model.new_crm_close=closestatus;
            model.old_closedate = DateTime.Now;
            model.new_closedate = DateTime.Now;
            model.createdate= DateTime.Now;
            model.new_createdate = DateTime.Now;
            model.desc_log = crm_do;
            var mService = ServiceHelper.Create<IQueryService>();
            mService.CreateCloseRecord(model);
           
        }



        /// <summary>
        /// 下載檔案
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.SalesSupport.CUSCRM.CUSCRMQU001")]
        public ActionResult DownloadCRMFile(string fileName)
        {
            RemoteFileInfo rf = null;
            var streamservice = ServiceHelper.Create<IStreamMediaService>();
            DownloadRequest dr = new DownloadRequest();
            dr.FileName = fileName;
            try
            {
                rf = streamservice.DownloadHistoryCRMFile(dr);
            }
            catch(Exception ex)
            {
                Throw.BusinessError(ex.Message);
                return Content("下載檔案失敗");
            }
            return File(rf.FileByteStream, rf.MimeType, dr.FileName);
        }


        /// <summary>
        /// 調整民國格式
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        private string GetChineseTimeFormat(DateTime time)
        {

            return string.Format("民國{0}年{1}月{2}日", (time.Year - 1911).ToString(), time.ToString("MM"), time.ToString("dd"));
        }

        
    }
}