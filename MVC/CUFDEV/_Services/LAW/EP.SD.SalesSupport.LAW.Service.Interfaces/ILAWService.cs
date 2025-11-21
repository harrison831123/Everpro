using EP.H2OModels;
using EP.SD.SalesSupport.LAW.Models;
using EP.VLifeModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.LAW.Service
{
    [ServiceContract]
    public interface ILAWService
    {
        #region 系統設定作業
        #region 管理人員設定
        /// <summary>
        /// 管理人員設定列表
        /// </summary>
        [OperationContract]
        List<LawSys> GetLawSys();

        /// <summary>
        /// 新增管理人員
        /// </summary>
        [OperationContract]
        void InsertLawSys(List<LawSys> model);

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawSysByID(string ID);

        /// <summary>
        /// 取得部門level
        /// </summary>
        /// <param name="unitid">部門編號</param>
        [OperationContract]
        int GetUnitlevel(string unitid);
        #endregion

        #region 經辦人員設定
        /// <summary>
        /// 管理人員設定列表
        /// </summary>
        [OperationContract]
        List<LawDoUser> GetLawDouser();

        /// <summary>
        /// 經辦人員詳細資料
        /// </summary>
        [OperationContract]
        List<LawDoUser> GetLawDouserByID(string ID);

        /// <summary>
        /// 新增管理人員
        /// </summary>
        [OperationContract]
        void InsertLawDouser(List<LawDoUser> model);

        /// <summary>
        /// 更新管理人員
        /// </summary>
        [OperationContract]
        void UpdateLawDouser(LawDoUser model);

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawDouserByID(string ID);

        /// <summary>
        /// 確認經辦順位
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        string CheckDouserSort(string DouserSort);

        /// <summary>
        /// 用ID確認經辦順位
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        string CheckDouserSortByID(string DouserSort, string LawDouserId);
        #endregion

        #region 承辦單位設定
        /// <summary>
        /// 承辦單位設定列表
        /// </summary>
        [OperationContract]
        List<LawDoUnit> GetLawDoUnit();

        /// <summary>
        /// 新增承辦單位
        /// </summary>
        [OperationContract]
        void InsertLawDoUnit(LawDoUnit model);

        /// <summary>
        /// 承辦單位詳細資料
        /// </summary>
        [OperationContract]
        List<LawDoUnit> GetLawDoUnitByID(string ID);

        /// <summary>
        /// 更新承辦單位
        /// </summary>
        [OperationContract]
        void UpdateLawDoUnit(LawDoUnit model);

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawDoUnitByID(string ID);


        /// <summary>
        /// 修改承辦狀態
        /// </summary>
        [OperationContract]
        void ChangeStatusType(string StatusType, string ID);
        #endregion

        #region 相關費率設定
        /// <summary>
        /// 相關費率設定列表
        /// </summary>
        [OperationContract]
        List<LawInterestRates> GetLawInterestRates();

        /// <summary>
        /// 新增相關費率
        /// </summary>
        [OperationContract]
        void InsertLawInterestRates(LawInterestRates model);

        /// <summary>
        /// 相關費率詳細資料
        /// </summary>
        [OperationContract]
        List<LawInterestRates> GetLawInterestRatesByID(string ID);

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawInterestRatesByID(string ID);

        /// <summary>
        /// 修改相關費率狀態
        /// </summary>
        [OperationContract]
        void ChangeInterestRatesType(string InterestRatesType, string ID);

        /// <summary>
        /// 律師費率設定列表
        /// </summary>
        [OperationContract]
        List<LawLawyerServiceRates> GetLawLawyerServiceRates();

        /// <summary>
        /// 新增律師費率
        /// </summary>
        [OperationContract]
        void InsertLawLawyerServiceRates(LawLawyerServiceRates model);

        /// <summary>
        /// 律師費率詳細資料
        /// </summary>
        [OperationContract]
        List<LawLawyerServiceRates> GetLawLawyerServiceRatesByID(string ID);

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawLawyerServiceRatesByID(string ID);

        /// <summary>
        /// 修改律師費率狀態
        /// </summary>
        [OperationContract]
        void ChangeLawyerServiceRatesType(string LawyerServiceRatesType, string ID);
        #endregion

        #region 結案狀態設定
        /// <summary>
        /// 結案狀態設定列表
        /// </summary>
        [OperationContract]
        List<LawCloseType> GetLawCloseType();

        /// <summary>
        /// 新增結案狀態
        /// </summary>
        [OperationContract]
        void InsertLawCloseType(LawCloseType model);

        /// <summary>
        /// 結案狀態詳細資料
        /// </summary>
        [OperationContract]
        List<LawCloseType> GetLawCloseTypeByID(string ID);

        /// <summary>
        /// 更新結案狀態
        /// </summary>
        [OperationContract]
        void UpdateLawCloseType(LawCloseType model);

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawCloseTypeByID(string ID);

        /// <summary>
        /// 修改結案狀態
        /// </summary>
        [OperationContract]
        void ChangeLawCloseTypeStatusType(string StatusType, string ID);
        #endregion

        #region 契撤變原因設定
        /// <summary>
        /// 契撤變原因設定列表
        /// </summary>
        [OperationContract]
        List<LawEvidType> GetLawEvidType();

        /// <summary>
        /// 新增契撤變原因
        /// </summary>
        [OperationContract]
        void InsertLawEvidType(LawEvidType model);

        /// <summary>
        /// 契撤變原因詳細資料
        /// </summary>
        [OperationContract]
        List<LawEvidType> GetLawEvidTypeByID(string ID);

        /// <summary>
        /// 更新契撤變原因
        /// </summary>
        [OperationContract]
        void UpdateLawEvidType(LawEvidType model);

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawEvidTypeByID(string ID);

        /// <summary>
        /// 修改契撤變原因狀態
        /// </summary>
        [OperationContract]
        void ChangeLawEvidTypeStatusType(string StatusType, string ID);

        /// <summary>
        /// 確認契撤變原因有無重複
        /// </summary>
        [OperationContract]
        bool chkEvidTypeNameRepeat(string EvidTypeName);
        #endregion

        #region 報表排序設定
        /// <summary>
        /// 報表排序列表VM
        /// </summary>
        [OperationContract]
        List<LawReportSort> GetLawReportSortVM(string SortYear);

        /// <summary>
        /// 報表排序列表SM
        /// </summary>
        [OperationContract]
        List<LawReportSort> GetLawReportSortSM(string SortYear);

        /// <summary>
        /// 更新VM報表排序
        /// </summary>
        [OperationContract]
        bool UpdateLawReportSortVM(string SortVmName, int? SortVm, string SortYear);

        /// <summary>
        /// 更新SM報表排序
        /// </summary>
        [OperationContract]
        bool UpdateLawReportSortSM(string SortSmName, string SortVmName, int? SortSm, string SortYear);

        /// <summary>
        /// 刪除年度排序
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawReportSortBySortYear(string SortYear);

        /// <summary>
        /// 確認報表排序有沒有設定新年度的資料，沒有則新增
        /// </summary>
        [OperationContract]
        void CheckLawReportSortBySortYear(string SortYear);
        #endregion

        #endregion

        #region 受理作業
        /// <summary>
        /// 檢查照會單號是否存在與流水號大小
        /// </summary>
        [OperationContract]
        LawNoteYmCounter CheckLawNoteYmCounter(string LawNoteYear, string LawNoteMonth);

        /// <summary>
        /// 檢查個人資料是否重複
        /// </summary>
        [OperationContract]
        LawAgentData CheckLawAgentDataRepeat(string agent_code);

        /// <summary>
        /// 檢查業務員相關資料
        /// </summary>
        [OperationContract]
        OrgibDetail CheckOrgib(string agent_code, string production_ym, string sequence);

        /// <summary>
        /// 檢查個人資料是否重複
        /// </summary>
        [OperationContract]
        LawAgentData CheckLawAgentData(string agent_code);

        /// <summary>
        /// 檢查八大區塊
        /// </summary>
        [OperationContract]
        orgsmb Checkorgsmb(string agent_code);

        /// <summary>
        /// 新增照會單號編號記錄檔
        /// </summary>
        [OperationContract]
        void InsertLawNoteYmCounter(LawNoteYmCounter noteYmCounter);

        /// <summary>
        /// 更新照會單號編號記錄檔
        /// </summary>
        [OperationContract]
        void UpdateLawNoteYmCounter(LawNoteYmCounter noteYmCounter);

        /// <summary>
        /// 新增法追主檔
        /// </summary>
        [OperationContract]
        void InsertLawContent(LawContent lawContent);

        /// <summary>
        /// 取得業務員地址(戶籍、住家)與電話
        /// </summary>
        [OperationContract]
        adin Getadin(string agent_code);


        /// <summary>
        /// 取得三代輔導主管
        /// </summary>
        [OperationContract]
        OrginDetail Getorgin(string agent_code);

        /// <summary>
        /// 依照照會單號取得主檔ID
        /// </summary>
        [OperationContract]
        LawContent GetLawContentByLawNoteNo(string LawNoteNo);

        /// <summary>
        /// 新增業務員個人資料檔
        /// </summary>
        [OperationContract]
        void InsertLawAgentData(LawAgentData model);

        /// <summary>
        /// 新增業務員資料檔
        /// </summary>
        [OperationContract]
        void InsertLawAgentContent(LawAgentContent model);

        /// <summary>
        /// 更新業務員個人資料檔
        /// </summary>
        [OperationContract]
        void UpdateLawAgentData(string agentcode, string name, string Rename);

        /// <summary>
        /// 更新業務員資料檔
        /// </summary>
        [OperationContract]
        void UpdateLawAgentContent(string agentcode, string name, string Rename);

        /// <summary>
        /// 新增業務員個人資料檔
        /// </summary>
        [OperationContract]
        void InsertLawNote(LawNote model);


        /// <summary>
        /// 抓取經辦人員
        /// </summary>
        [OperationContract]
        LawDoUser GetLawDouserTopOne();

        /// <summary>
        /// 新增進度表訴訟程序記錄檔
        /// </summary>
        [OperationContract]
        void InsertLawLitigationProgress(LawLitigationProgress model);

        /// <summary>
        /// 新增進度表執行程序紀錄
        /// </summary>
        [OperationContract]
        void InsertLawDoProgress(LawDoProgress model);

        #endregion

        #region 通知作業
        /// <summary>
        /// 通知作業查詢列表
        /// </summary>
        [OperationContract]
        List<LawNote> GetLawNote(string LawNoteNo, int sys, string MemberID, string LawNoteName, string orgID);

        /// <summary>
        /// 系統設定作業目錄,判斷是否為系統管理人員權限
        /// </summary>
        [OperationContract]
        int CheckLawNoteByMemberID(string MemberID);
        #endregion

        #region 維護作業
        /// <summary>
        /// 維護作業查詢列表
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawContent(LawContent model);

        /// <summary>
        /// 取得啟用狀態承辦單位
        /// </summary>
        [OperationContract]
        List<LawDoUnit> GetLawDoUnitByEnable();

        /// <summary>
        /// 取得個人資料頁面
        /// </summary>
        [OperationContract]
        LawAgentDetail GetLawAgentDetail(string AgentID, string NoteNo);

        /// <summary>
        /// 取得前案情形
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawContentByIDNo(string AgentID, string NoteNo);

        /// <summary>
        /// 更新業務員資料檔
        /// </summary>
        [OperationContract]
        void UpdateLawAgentDetail(LawAgentDetail agentDetail);

        /// <summary>
        /// 取得進度表頁面
        /// </summary>
        [OperationContract]
        LawContent GetSchedule(string AgentID, string NoteNo);

        /// <summary>
        /// 取得進度表存證信函備註記
        /// </summary>
        [OperationContract]
        List<LawEvidenceDesc> GetLawEvidenceDesc(string LawId);

        /// <summary>
        /// 取得分案日期
        /// </summary>
        [OperationContract]
        LawDoUnitLog GetLawDoUnitLog(string LawId);

        /// <summary>
        /// 取得啟用狀態承辦單位ID
        /// </summary>
        [OperationContract]
        LawDoUnit GetLawDoUnitByUnitName(string UnitName);

        /// <summary>
        /// 取得承辦單位歷史紀錄
        /// </summary>
        [OperationContract]
        List<LawDoUnitLog> GetLawDoUnitLogByLawId(string LawId);

        /// <summary>
        /// 取得進度表訴訟程序
        /// </summary>
        [OperationContract]
        List<LawLitigationProgress> GetLawLitigationProgressByLawIdNotoNo(string LawId, string LawNoteNo);

        /// <summary>
        /// 取得進度表訴訟程序
        /// </summary>
        [OperationContract]
        List<LawDoProgress> GetLawDoProgressByLawIdNotoNo(string LawId, string LawNoteNo);

        /// <summary>
        /// 取得進度表其他說明
        /// </summary>
        [OperationContract]
        List<LawOtherDescDetail> GetLawOtherDescByLawIdNotoNo(string LawId, string LawNoteNo);

        /// <summary>
        /// 取得進度表清償金額
        /// </summary>
        [OperationContract]
        List<LawRepaymentList> GetLawRepaymentListByLawIdNotoNo(string AgentID, string LawId, string LawNoteNo);

        /// <summary>
        /// 取得進度表備註
        /// </summary>
        [OperationContract]
        List<LawDescDesc> GetLawDescDescByLawIdNotoNo( string LawId, string LawNoteNo);

        /// <summary>
        /// 取得啟用狀態結案狀態設定
        /// </summary>
        [OperationContract]
        List<LawCloseType> GetLawCloseTypeByEnable();

        /// <summary>
        /// 取得進度表案件狀態其他說明
        /// </summary>
        [OperationContract]
        LawCloseOtherLog GetLawCloseOtherLogByLawId(string LawId);

        /// <summary>
        /// 更新分案日期
        /// </summary>
        [OperationContract]
        void UpdateCaseDate(LawDoUnitLog model);

        /// <summary>
        /// 新增進度表存證信函備註記
        /// </summary>
        [OperationContract]
        void InsertLawEvidenceDesc(LawEvidenceDesc model);

        /// <summary>
        /// 新增其他
        /// </summary>
        [OperationContract]
        void InsertLawOtherDesc(LawOtherDesc model);

        /// <summary>
        /// 新增其他金額說明
        /// </summary>
        [OperationContract]
        void InsertLawOtherDescHaveMoney(LawOtherDesc model);

        /// <summary>
        /// 新增備註
        /// </summary>
        [OperationContract]
        void InsertLawDescDesc(LawDescDesc model);

        /// <summary>
        /// 新增承辦單位設定記錄log和更新主檔承辦單位
        /// </summary>
        [OperationContract]
        void InsertLawDoUnitLogUpdateLawContentByLawDoUnitId(LawDoUnitLog model, LawContent content);

        /// <summary>
        /// 新增電催通知記錄log和更新主檔承辦單位
        /// </summary>
        [OperationContract]
        void InsertLawPhoneCallLogUpdateLawContentByLawId(LawPhoneCallLog model, LawContent content);

        /// <summary>
        /// 新增地2次電催通知記錄log和更新主檔承辦單位
        /// </summary>
        [OperationContract]
        void InsertLawPhoneCallLog2UpdateLawContentByLawId(LawPhoneCallLog model, LawContent content);

        /// <summary>
        /// 新增結案狀況其他欄位說明記錄log
        /// </summary>
        [OperationContract]
        void InsertLawCloseOtherLog(LawCloseOtherLog model);

        /// <summary>
        /// 存證信函編輯頁面
        /// </summary>
        [OperationContract]
        LawEvidenceDesc GetLawEvidenceDescById(int LawEvidenceId);

        /// <summary>
        /// 訴訟程序編輯頁面
        /// </summary>
        [OperationContract]
        LawLitigationProgress GetLawLitigationProgressById(int LawLitigationId);

        /// <summary>
        /// 執行程序編輯頁面
        /// </summary>
        [OperationContract]
        LawDoProgress GetLawDoProgressById(int LawDoId);

        /// <summary>
        /// 其他說明編輯頁面
        /// </summary>
        [OperationContract]
        LawOtherDesc GetLawOtherDescById(int LawOtherId);

        /// <summary>
        /// 備註編輯頁面
        /// </summary>
        [OperationContract]
        LawDescDesc GetLawLawDescDescById(int LawDescId);

        /// <summary>
        /// 更新存證信函
        /// </summary>
        [OperationContract]
        void UpdateLawEvidenceDesc(LawEvidenceDesc model);

        /// <summary>
        /// 其他金額編輯頁面
        /// </summary>
        [OperationContract]
        LawOtherDesc GetLawOtherDescByLawId(string LawDueAgentId, int LawId, int LawRepaymentId);

        /// <summary>
        /// 其他金額說明編輯頁面
        /// </summary>
        [OperationContract]
        LawRepaymentList GetLawRepaymentListById(int LawRepaymentId);

        /// <summary>
        /// 更新訴訟程序
        /// </summary>
        [OperationContract]
        void UpdateLawLitigationProgress(LawLitigationProgress model, LawContent content);

        /// <summary>
        /// 更新執行程序
        /// </summary>
        [OperationContract]
        void UpdateLawDoProgress(LawDoProgress model, LawContent content);

        /// <summary>
        /// 更新其他
        /// </summary>
        [OperationContract]
        void UpdateLawOtherDesc(LawOtherDesc model);

        /// <summary>
        /// 更新備註
        /// </summary>
        [OperationContract]
        void UpdateLawDescDesc(LawDescDesc model);

        /// <summary>
        /// 更新結案原因和卷宗
        /// </summary>
        [OperationContract]
        void UpdateLawContentReasonfile(LawContent model);

        /// <summary>
        /// 更新未結案狀態
        /// </summary>
        [OperationContract]
        void UpdateLawContentNotCloseType(LawContent model);

        /// <summary>
        /// 更新為案件取消
        /// </summary>
        [OperationContract]
        void UpdateLawContentCancelType(LawContent model);

        /// <summary>
        /// 更新結案狀態
        /// </summary>
        [OperationContract]
        void UpdateLawContentCloseType(LawContent model);

        /// <summary>
        /// 用ID刪除單筆存證信函資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawEvidenceDescByID(string ID);

        /// <summary>
        /// 用ID刪除單筆訴訟程序資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawLitigationProgressByID(string ID);

        /// <summary>
        /// 用ID刪除單筆執行程序資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawDoProgressByID(string ID);

        /// <summary>
        /// 用ID刪除單筆其他資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawOtherDescByID(string ID);

        /// <summary>
        /// 用ID刪除單筆備註資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawDescDescByID(string ID);

        /// <summary>
        /// 結欠明細
        /// </summary>
        [OperationContract]
        List<LawContent> GetDueMoneyDetail(string AgentID, string NoteNo);

        /// <summary>
        /// 查詢主檔
        /// </summary>
        [OperationContract]
        LawContent GetLawContentByLawId(int LawId);

        /// <summary>
        /// 查詢0利率
        /// </summary>
        [OperationContract]
        LawInterestRates GetLawInterestRatesByZero();

        /// <summary>
        /// 清償明細
        /// </summary>
        [OperationContract]
        List<LawRepaymentList> GetLawRepaymentListByLawId(int LawId);

        /// <summary>
        /// 確認清償明細可否新增
        /// </summary>
        [OperationContract]
        LawContent ChkIRInsert(int LawId, string LawDueAgentId);

        /// <summary>
        /// 累計清償
        /// </summary>
        [OperationContract]
        LawRepaymentList GetTotalLawRepaymentMoney(string LawDueAgentId, IEnumerable<int> LawRepaymentId);

        /// <summary>
        /// 剩餘結欠金額
        /// </summary>
        [OperationContract]
        LawContent GetRemainLawRepaymentMoney(string LawDueAgentId, int LawId);

        /// <summary>
        /// 查詢律師服務費 
        /// </summary>
        [OperationContract]
        List<LawLawyerReward> GetLawLawyerReward(int LawId, string LawDueAgentId);

        /// <summary>
        /// 律師服務費編輯頁面
        /// </summary>
        [OperationContract]
        LawLawyerReward GetLawLawyerRewardById(int LawLawyerPayId);

        /// <summary>
        /// 存證信函
        /// </summary>
        [OperationContract]
        LawEvidenceDetail GetLawEvidenceById(string EvidAgentId, string LawNoteNo);

        /// <summary>
        /// 存證信函個人資料
        /// </summary>
        [OperationContract]
        LawEvidenceDetail GetLawAgentDetailById(string AgentID, string NoteNo);

        /// <summary>
        /// 照會單
        /// </summary>
        [OperationContract]
        List<LawNote> GetLawNoteByLawNoteNo(string LawNoteNo);

        /// <summary>
        /// 照會單作業-已發送紀錄
        /// </summary>
        [OperationContract]
        List<MessageTo> GetMessageToByLawNoteNo(string LawNoteNo);

        /// <summary>
        /// 新增結欠明細Log
        /// </summary>
        [OperationContract]
        void InsertLawInterestLog(LawInterestLog model);

        /// <summary>
        /// 新增清償明細
        /// </summary>
        [OperationContract]
        void InsertLawRepaymentList(LawRepaymentList model);

        /// <summary>
        /// 新增清償明細Log
        /// </summary>
        [OperationContract]
        void InsertLawRepaymentListLog(LawRepaymentListLog model);


        /// <summary>
        /// 新增律師服務費
        /// </summary>
        [OperationContract]
        void InsertLawLawyerReward(LawLawyerReward model);

        /// <summary>
        /// 新增存證信函
        /// </summary>
        [OperationContract]
        void InsertLawEvidence(LawEvidence model);

        /// <summary>
        /// 新增存證信函Log
        /// </summary>
        [OperationContract]
        void InsertLawEvidenceLog(LawEvidenceLog model);

        /// <summary>
        /// 更新存證信函
        /// </summary>
        [OperationContract]
        void UpdateLawEvidence(LawEvidence model);

        /// <summary>
        /// 更新為案件取消
        /// </summary>
        [OperationContract]
        void UpdateLawContentDueMoney(LawContent model);

        /// <summary>
        /// 更新最後編輯日
        /// </summary>
        [OperationContract]
        void UpdateLawContentLawContentLastchangeDate(LawContent model);

        /// <summary>
        /// 更新律師服務費
        /// </summary>
        [OperationContract]
        void UpdateLawLawyerReward(LawLawyerReward model);

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        [OperationContract]
        bool DeleteLawRepaymentListByID(string ID);
        #endregion

        #region 查詢作業
        /// <summary>
        /// 查詢作業
        /// </summary>
        [OperationContract]
        List<LawSearchDetail> GetLawContentBySearch(LawSearchDetail model);

        /// <summary>
        /// 取得啟用狀態結案狀態設定
        /// </summary>
        [OperationContract]
        List<LawCloseType> GetLawCloseTypeByLawCloseType(string LawCloseType);

        /// <summary>
        /// 產出報表
        /// </summary>
        [OperationContract]
        byte[] GetSearchReportList(LawSearchDetail condition);

        /// <summary>
        /// 產出報表
        /// </summary>
        [OperationContract]
        byte[] GetNotCloseReportList(LawSearchDetail model);

        [OperationContract]
        byte[] GetCloseReportList(LawSearchDetail model);

        /// <summary>
        /// 確認是否為管理人員
        /// </summary>
        [OperationContract]
        LawSys GetLawSysBySysMemberID(string SysMemberID);

        /// <summary>
        /// 確認VMSM
        /// </summary>
        [OperationContract]
        OrgVm GetOrfVm(string AccID);

        /// <summary>
        /// 未結案件明細列表
        /// </summary>
        [OperationContract]
        List<LawSearchDetail> GetLawContentByNotClose(LawSearchDetail model, OrgVm orgVm, LawSys lawSys);

        /// <summary>
        /// 結案件明細列表
        /// </summary>
        [OperationContract]
        List<LawSearchDetail> GetLawContentByClose(LawSearchDetail model, OrgVm orgVm, LawSys lawSys);

        /// <summary>
        /// 追回金額
        /// </summary>
        [OperationContract]
        LawRepaymentList GetLawRepaymentListByAgID(string LawDueAgentId, string LawId);

        /// <summary>
        /// 存證信函
        /// </summary>
        [OperationContract]
        LawEvidenceDesc GetLawEvidenceDescByLawId(string LawId);
        #endregion

        #region 報表作業
        /// <summary>
        /// 報表作業
        /// </summary>
        [OperationContract]
        List<LawMasterReportLog> GetLawMasterReportLog();

        /// <summary>
        /// 節欠金額
        /// </summary>
        [OperationContract]
        LawContent GetLawcontentByLawyear(int year);

        /// <summary>
        /// 節欠金額合計
        /// </summary>
        [OperationContract]
        LawContent GetLawcontentByLawyears(string year_str);

        /// <summary>
        /// 節欠總金額
        /// </summary>
        [OperationContract]
        LawMasterReportLog GetLawMasterReportLogByLawyear(int y_str, string year_str);

        /// <summary>
        /// 節欠其他總金額
        /// </summary>
        [OperationContract]
        LawMasterReportLog GetLawMasterReportLogOtherByLawyear(int y_str, int Lawyear);


        /// <summary>
        /// 清償金額合計
        /// </summary>
        [OperationContract]
        LawRepaymentList GetSumLawRepaymentList(string year_str);

        /// <summary>
        /// 清償金額
        /// </summary>
        [OperationContract]
        LawRepaymentList GetLawRepaymentList(int Lawyear);

        /// <summary>
        /// 系統紀錄log
        /// </summary>
        [OperationContract]
        bool InsertLawMasterReportLog(string MeberID);

        [OperationContract]
        Stream GetStatisticReportList(string LawYearType);

        [OperationContract]
        Stream GetSumStatisticReportList(string LawYearType);

        /// <summary>
        /// 當月還款明細
        /// </summary>
        [OperationContract]
        List<LawMonthRepaymentReportDetail> GetLawMonthRepaymentReport(LawMonthRepaymentReportDetail model);

        /// <summary>
        /// 當月還款明細總計
        /// </summary>
        [OperationContract]
        LawRepaymentList GetSumLawRepaymentMoney(LawMonthRepaymentReportDetail model);

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="year">年度</param>
        /// <param name="month">月份</param>
        /// <returns>Stream</returns>
        [OperationContract]
        Stream QueryLawMonthRepaymentReport(string year, string month, string chkm);

        /// <summary>
        /// 團隊明細
        /// </summary>
        [OperationContract]
        List<LawVmSmReport> GetLawVmSmReport(LawVmSmReport model);

        /// <summary>
        /// 團隊明細年度金額加總
        /// </summary>
        [OperationContract]
        LawVmSmDetail GetSumLawVmSmReportMoney(LawVmSmReport model, string year_str);

        /// <summary>
        /// 團隊明細年度金額
        /// </summary>
        [OperationContract]
        LawVmSmDetail GetLawVmSmReportMoney(LawVmSmReport model);

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="year">年度</param>
        /// <returns>Stream</returns>
        [OperationContract]
        Stream QueryTeamReport(string year);

        /// <summary>
        /// 上月總累計未結案件
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawTopYearReport();


        /// <summary>
        /// 本月結案
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawNowYearReport(string mm);

        /// <summary>
        /// 合計未結總數
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawSumNotTypeReport(string mm);

        /// <summary>
        /// 委外案件數
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawOutSideReport(string mm);

        /// <summary>
        /// 自辦案件數
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawSelfReport(string mm);

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="year">年度</param>
        /// <param name="month">月份</param>
        /// <returns>Stream</returns>
        [OperationContract]
        Stream QueryLawYearReport();

        /// <summary>
        /// 受理案件
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawAccpet(string year, string mm);

        /// <summary>
        /// 受理案件總計
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawSumAccpet(string year);

        /// <summary>
        /// 委外件數
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawOutsource(string year, string mm);

        /// <summary>
        /// 委外件數總計
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawSumOutsource(string year);


        /// <summary>
        /// 自辦件數
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawDoself(string year, string mm);

        /// <summary>
        /// 自辦件數總計
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawSumDoself(string year);

        /// <summary>
        /// 未結
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawNotOC(string year, string mm);

        /// <summary>
        /// 未結總計
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawSumNotOC(string year);

        /// <summary>
        /// 結案
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawOC(string year, string mm);

        /// <summary>
        /// 結案總計
        /// </summary>
        [OperationContract]
        List<LawContent> GetLawSumOC(string year);

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="year">年度</param>
        /// <param name="month">月份</param>
        /// <returns>Stream</returns>
        [OperationContract] 
        Stream QueryLawNowYearReport();
        #endregion

        #region 電話催告通知
        /// <summary>
        /// 電話催告通知
        /// </summary>
        [OperationContract]
        List<LawPhoneCallLogDetail> GetLawPhoneCallLog();


        /// <summary>
        /// 維護日期逾30日之案件列表通知
        /// </summary>
        [OperationContract]
        List<LawPhoneCallLogDetail> GetThirtyDayLawContent();

        /// <summary>
        /// 更新催告通知
        /// </summary>
        [OperationContract]
        void UpdateLawPhoneCallLog(string PhoneCallReadID);
        #endregion

        #region 法追明細
        /// <summary>
        /// 未結案件列表
        /// </summary>
        [OperationContract]
        List<LawSearchDetail> GetLawSearchByNotClose(LawSearchDetail model, OrgVm orgVm);

        /// <summary>
        /// 已結案件列表
        /// </summary>
        [OperationContract]
        List<LawSearchDetail> GetLawSearchByClose(LawSearchDetail model, OrgVm orgVm);
        #endregion
    }
}
