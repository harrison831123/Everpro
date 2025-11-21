using EP.PSL.WorkResources.MeetingMng.Models;
using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EP.PSL.WorkResources.MeetingMng.Web.Areas.MeetingMng.Model
{
    public class MeetingDetailModel
    {
        /// <summary>會議流水號</summary>
        [Display(Name = "流水號", ResourceType = typeof(MeetingMngResource))]
        public int MTID { get; set; }
        /// <summary>會議地點流水號</summary>
        [Display(Name = "流水號", ResourceType = typeof(MeetingMngResource))]
        public int MDID { get; set; }
        /// <summary>會議名稱</summary>
        [Display(Name = "會議名稱", ResourceType = typeof(MeetingMngResource))]
        public string MTName { get; set; }

        /// <summary>會議說明</summary>
        [Display(Name = "說明", ResourceType = typeof(MeetingMngResource))]
        public string MTDesc { get; set; }

        /// <summary>開始日期</summary>
        [Display(Name = "開始日期", ResourceType = typeof(MeetingMngResource))]
        public DateTime MTStartDate { get; set; }

        /// <summary>結束日期</summary>
        [Display(Name = "結束日期", ResourceType = typeof(MeetingMngResource))]
        public DateTime MTEndDate { get; set; }

        /// <summary>建立日期</summary>
        [Display(Name = "建立日期", ResourceType = typeof(MeetingMngResource))]
        public string MTCreatedate { get; set; }

        /// <summary>判斷是召開會議或歷史資料</summary>
        [Display(Name = "會議狀態", ResourceType = typeof(MeetingMngResource))]
        public int MTActive { get; set; }

        ///// <summary>文件檔案數量</summary>
        //[DisplayName("文件檔案數量")]
        //public int MTFileNum { get; set; }

        /// <summary>會議地點</summary>
        [Display(Name = "會議地點", ResourceType = typeof(MeetingMngResource))]
        public string MTPlace { get; set; }

        /// <summary>新增會議地點</summary>
        [DisplayName("新增會議地點")]
        public string MTNewPlace { get; set; }

        /// <summary>召集人ID</summary>
        [DisplayName("召集人")]
        public string MTConvener { get; set; }

        /// <summary>確認召集人是否是與會人員</summary>
        [DisplayName("確認召集人是否是與會人員")]
        public bool Convenerchk { get; set; }

        /// <summary>主席</summary>
        [DisplayName("主席ID")]
        public string MTChairman { get; set; }

        /// <summary>紀錄者</summary>
        [DisplayName("紀錄者ID")]
        public string MTRecorder { get; set; }

        /// <summary>與會人員</summary>
        [Display(Name = "與會人員", ResourceType = typeof(MeetingMngResource))]
        public string Participants { get; set; }

        /// <summary>附加檔案名稱</summary>
        [Display(Name = "附加檔案名稱", ResourceType = typeof(MeetingMngResource))]
        public string MFFileName { get; set; }

        /// <summary>
        /// 附件檔
        /// </summary>
        //[DisplayName("附件檔")]
        [Display(Name = "附件檔", ResourceType = typeof(MeetingMngResource))]
        public List<MeetingFileModel> File { get; set; }


        /// <summary>場地與設備名稱</summary>
        [Display(Name = "場地與設備", ResourceType = typeof(MeetingMngResource))]
        public string MDName { get; set; }

        /// <summary>已回覆參加人員</summary>
        [Display(Name = "已回覆參加人員", ResourceType = typeof(MeetingMngResource))]
        public string Participate { get; set; }

        /// <summary>未回覆人員</summary>
        [Display(Name = "未回覆人員", ResourceType = typeof(MeetingMngResource))]
        public string NoReply { get; set; }

        /// <summary>已回覆不參加人員</summary>
        [Display(Name = "已回覆不參加人員", ResourceType = typeof(MeetingMngResource))]
        public string NoParticipate { get; set; }

        /// <summary>判斷是否回覆參加會議 0 還沒回覆 1 要參加 2 不參加</summary>
        [Display(Name = "參加狀態", ResourceType = typeof(MeetingMngResource))]
        public int MTReply { get; set; }

        /// <summary>設定成員名單給{Jason}</summary>
        [NonColumn]
        [DisplayName("設定成員名單給{Jason}")]
        public string RecipientToJson { get; set; }

        [NonColumn]
        [DisplayName("留言附檔名")]
        public string UploadFilesName { get; set; }

        /// <summary>
        /// TabUniqueId 用來取得暫存資料夾的名稱
        /// </summary>
        [NonColumn]
        public string TabUniqueId { get; set; }

        /// <summary>會議檔案流水號</summary>
        [DisplayName("流水號")]
        public int MFID { get; set; }

        /// <summary>檔案路徑</summary>
        [DisplayName("檔案路徑")]
        public string MFFilePath { get; set; }

        /// <summary>檔案大小</summary>
        [DisplayName("檔案大小")]
        public int MFFileSize { get; set; }

        /// <summary>Md5檔名</summary>
        [DisplayName("MFMd5Name")]
        public string MFMd5Name { get; set; }

        /// <summary>Md5判斷檔案來源</summary>
        [DisplayName("MFType")]
        public int MFType { get; set; }

        /// <summary>檔案說明</summary>
        [Display(Name = "說明", ResourceType = typeof(MeetingMngResource))]
        public string MFDesc { get; set; }

        /// <summary>檔案上傳者編號</summary>
        [Display(Name = "檔案上傳者", ResourceType = typeof(MeetingMngResource))]
        public string MFCreater { get; set; }

        /// <summary>建立檔案日期</summary>
        [Display(Name = "建立日期", ResourceType = typeof(MeetingMngResource))]
        public string MFCreateDate { get; set; }

        /// <summary></summary>
        [DisplayName("MFOrderBy")]
        public int MFOrderBy { get; set; }

        /// <summary>報告部門</summary>
        [Display(Name = "報告部門", ResourceType = typeof(MeetingMngResource))]
        public string nunit { get; set; }

        /// <summary>員編</summary>
        public string imember { get; set; }

        /// <summary>人名</summary>
        public string nmember { get; set; }

        /// <summary>追蹤事項</summary>       
        [Display(Name = "追蹤事項", ResourceType = typeof(MeetingMngResource))]
        public string JBSubject { get; set; }

        /// <summary>追蹤說明</summary>       
        [Display(Name = "說明", ResourceType = typeof(MeetingMngResource))]
        public string JBDesc { get; set; }

        /// <summary>開始日期</summary>
        [Display(Name = "開始日期", ResourceType = typeof(MeetingMngResource))]
        public DateTime JBStartDate { get; set; }

        /// <summary>結束日期</summary>
        [Display(Name = "結束日期", ResourceType = typeof(MeetingMngResource))]
        public DateTime JBEndDate { get; set; }

        /// <summary>召集人姓名&部門</summary>
        [Display(Name = "召集人", ResourceType = typeof(MeetingMngResource))]
        public string MTConvenerName { get; set; }

        /// <summary>主席姓名&部門</summary>
        [Display(Name = "主席", ResourceType = typeof(MeetingMngResource))]
        public string MTChairmanName { get; set; }

        /// <summary>紀錄者姓名&部門</summary>
        [Display(Name = "紀錄者", ResourceType = typeof(MeetingMngResource))]
        public string MTRecorderName { get; set; }

        /// <summary>傳送訊息</summary>
        [Display(Name = "傳送訊息", ResourceType = typeof(MeetingMngResource))]
        public bool MTMessageChk { get; set; }

        /// <summary>執行人員</summary>
        [Display(Name = "執行人員", ResourceType = typeof(MeetingMngResource))]
        public string Jobimember { get; set; }

        /// <summary>開始時間小時</summary>
        //[Display(Name = "小時", ResourceType = typeof(MeetingMngResource))]
        public string shour { get; set; }

        /// <summary>開始時間分</summary>
        //[Display(Name = "分", ResourceType = typeof(MeetingMngResource))]
        public string smin { get; set; }

        /// <summary>結束時間小時</summary>
        //[Display(Name = "小時", ResourceType = typeof(MeetingMngResource))]
        public string ehour { get; set; }

        /// <summary>結束時間分</summary>
        //[Display(Name = "分", ResourceType = typeof(MeetingMngResource))]
        public string emin { get; set; }

        /// <summary>判斷會議地點自訂選擇</summary>
        //[Display(Name = "判斷會議地點自訂選擇", ResourceType = typeof(MeetingMngResource))]
        public bool chkplace { get; set; }
        



    }
}