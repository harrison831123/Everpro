using EP.Platform.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using EP.H2OModels;

namespace EP.PSL.WorkResources.MeetingMng.Service
{
    [ServiceContract]
    public interface IMeetingMngService
    {
        #region 會議列表
        /// <summary>
        /// 修改召開會議時間狀態
        /// </summary>
        [OperationContract]
        void UpdateMeetingTime();

        /// <summary>
        /// 取得查詢條件的召開會議
        /// </summary>
        [OperationContract]
        List<Meeting> GetMeetingList(QueryMeetingCondition cond);

        /// <summary>
        /// 取得會議的類型名稱列表
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        List<ValueText> GetMeetingKindList();

        /// <summary>
        /// 用會議流水號取得詳細內容
        /// </summary>
        [OperationContract]
        Meeting GetMeetingDetailById(int id, string imember);

        /// <summary>
        /// 用會議流水號取得詳細內容
        /// </summary>
        [OperationContract]
        string GetMDIDToShow(int MDID);

        /// <summary>
        /// 用流水號取得與會人員
        /// </summary>
        [OperationContract]
        List<MeetingMemberRelation> GetMeetingDetailParticipantsById(int id);

        /// <summary>
        /// 用流水號取得編輯頁面場地與設備值
        /// </summary>
        [OperationContract]
        MeetingDevice GetEditMeetingDevice(int id);

        /// <summary>
        /// 用流水號取得編輯頁面場地與設備值
        /// </summary>
        [OperationContract]
        List<MeetingDeviceRelation> GetEditMeetingDeviceRelation(int id);

        /// <summary>
        /// 詳細頁面檔案顯示
        /// </summary>
        /// <param name="id">會議序號</param>
        [OperationContract]
        List<MeetingFile> GetMeetingDetailFilesById(int id);

        /// <summary>
        /// 詳細頁面參加狀態顯示
        /// </summary>
        /// <param name="id">會議序號</param>
        /// <param name="imember">登入成員</param>
        [OperationContract]
        MeetingMemberRelation GetMeetingDetailMTReplyById(int id, string imember);

        /// <summary>
        /// 用流水號取得已回覆參加人員
        /// </summary>
        //[OperationContract]
        //List<MeetingMemberRelation> GetMeetingDetailParticipateById(int id);

        /// <summary>
        /// 用流水號取得不參加人員
        /// </summary>
        //[OperationContract]
        //List<MeetingMemberRelation> GetMeetingDetailNoParticipateById(int id);

        /// <summary>
        /// 用流水號取得未回覆參加人員
        /// </summary>
        //[OperationContract]
        //List<MeetingMemberRelation> GetMeetingDetailNoReplyById(int id);

        /// <summary>
        /// 用流水號更改參加會議參加狀態
        /// </summary>
        [OperationContract]
        void UpdateMeetingReplyById(int mtid, int mtreply, string imember);

        /// <summary>
        /// 用流水號更改參加會議參加狀態
        /// </summary>
        [OperationContract]
        void UpdateMeetingActiveById(int mtid, int mtactive);

        /// <summary>
        /// 用會議檔案流水號取得單一檔案
        /// </summary>
        [OperationContract]
        MeetingFile GetMeetingFileByMfid(int fid);

        /// <summary>
        /// 用會議序號取得會議檔案
        /// </summary>
        [OperationContract]
        List<MeetingFile> GetMeetingFileByMTId(int id);

        /// <summary>
        /// 用會議流水號取得會議資料
        /// </summary>
        [OperationContract]
        Meeting GetMeeting(int id);
        #endregion

        #region 會議CRUD
        /// <summary>
        /// 取得會議地點列表
        /// </summary>
        [OperationContract]
        List<Meeting> GetMeetingPlace();

        /// <summary>
        /// 取得場地和設備列表
        /// </summary>
        [OperationContract]
        List<MeetingDevice> GetMeetingDevice();

        /// <summary>
        /// 用流水號取得編輯頁面主席部門和名字
        /// </summary>
        [OperationContract]
        List<MeetingDetail> GetEditMeetingChairmanById(int id);

        
        /// <summary>
        /// 用流水號取得編輯頁面紀錄者部門和名字
        /// </summary>
        [OperationContract]
        List<MeetingDetail> GetEditMeetingRecorderById(int id);

        /// <summary>
        /// 用流水號取得編輯頁面與會人員部門和名字
        /// </summary>
        [OperationContract]
        List<MeetingDetail> GetEditMeetingParticipantsById(int id);

        /// <summary>
        /// 新增場地和設備關聯
        /// </summary>
        [OperationContract]
        void CreateMeetingDeviceRelation(MeetingDeviceRelation model);

        /// <summary>
        /// 新增會議管理&設備中心預定的會議室和設備
        /// </summary>
        [OperationContract]
        void CreateMeetingOrder(List<MeetingOrder> modeltoList);

        /// <summary>
        /// 新增會議
        /// </summary>
        [OperationContract]
        int CreateMeeting(Meeting model);

        /// <summary>
        /// 新增會議資訊&成員資訊關聯&回覆確認
        /// </summary>
        [OperationContract]
        void CreateMeetingMemberRelation(List<MeetingMemberRelation> modeltoList);

        /// <summary>
        /// 新增會議檔案&相關檔案
        /// </summary>
        [OperationContract]
        void CreateMeetingFile(Meeting model, List<MeetingFile> modelfileList, string tabUniqueId, out string RetMsg);

        /// <summary>
        /// 檢查會議時間有無重複
        /// </summary>
        [OperationContract]
        MeetingDetail ChkMeetingTime(int MDID, DateTime SDate, DateTime EDate);

        /// <summary>
        /// 檢查修改會議時間有無重複
        /// </summary>
        [OperationContract]
        MeetingDetail ChkEditMeetingTime(int MDID, DateTime SDate, DateTime EDate,int MTID);

        /// <summary>
        /// 編輯頁面檢查設備與地點有無
        /// </summary>
        [OperationContract]
        bool ChkMeetingDevice(int MTID);

        /// <summary>
        /// 修改設備與場地關聯
        /// </summary>
        [OperationContract]
        void EditMeetingDeviceRelation(List<MeetingDeviceRelation> modeltoList);

        /// <summary>
        /// 修改會議
        /// </summary>
        [OperationContract]
        void EditMeeting(Meeting model);

        /// <summary>
        /// 修改會議檔案
        /// </summary>
        [OperationContract]
        void EditMeetingFile(Meeting model, List<MeetingFile> modelfileList, string tabUniqueId, out string RetMsg);

        /// <summary>
        /// 修改會議管理&設備中心預定的會議室和設備
        /// </summary>
        [OperationContract]
        void EditMeetingOrder(List<MeetingOrder> modeltoList);

        /// <summary>
        /// 修改會議資訊&人員資訊關聯&回覆確認
        /// </summary>
        [OperationContract]
        void EditMeetingMemberRelation(List<MeetingMemberRelation> modeltoList);

        /// <summary>
        /// 刪除該筆會議檔案
        /// </summary>
        /// <param name="MTID"></param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteMeeting(int MTID);

        /// <summary>
        /// 會議刪除設備中心連動
        /// </summary>
        /// <param name="MTID"></param>
        [OperationContract]
        bool DeleteMeetingOrder(int MTID);

        /// <summary>
        /// 編輯會議時會議檔案刪除功能
        /// </summary>
        /// <param name="mf">會議檔案</param>
        [OperationContract]
        void DeleteEditMeetingFile(MeetingFile mf);
        #endregion

        #region 相關檔案
        /// <summary>
        /// 取得查詢條件的相關檔案
        /// </summary>
        [OperationContract]
        List<MeetingFileTo> GetMeetingRelatedFilesById(QueryMeetingFilesCondition cond);

        /// <summary>
        /// 刪除該筆相關檔案
        /// </summary>
        /// <param name="MFID"></param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteMeetingFile(int MFID);
        #endregion

        #region 決議事項
        /// <summary>
        /// 取得查詢條件的決議事項
        /// </summary>
        [OperationContract]
        List<JobDetail> GetMeetingJobById(QueryMeetingJobCondition cond);

        /// <summary>
        /// 新增決議事項
        /// </summary>
        [OperationContract]
        int CreateJob(Job model);

        /// <summary>
        /// 新增會議資訊&決議事項關聯表
        /// </summary>
        [OperationContract]
        void CreateMeetingJobRelation(MeetingJobRelation model);

        /// <summary>
        /// 新增決議事項進度表
        /// </summary>
        [OperationContract]
        void CreateJobPercentage(List<JobPercentage> modeltoList);

        /// <summary>
        /// 用決議事項流水號取得詳細內容
        /// </summary>
        [OperationContract]
        JobDetail GetJobDetailById(int JPID);

        /// <summary>
        /// 會議追蹤事項查詢條件
        /// </summary>
        [OperationContract]
        List<JobDetail> GetMeetingJobList(QueryMeetingJobCondition cond);

        /// <summary>
        /// 刪除該筆決議事項
        /// </summary>
        /// <param name="JBID"></param>
        /// <param name="JPID"></param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteJob(int JBID, int JPID);
        #endregion
    }
}
