using EB.EBrokerModels;
using EB.SL.MerSal.Models;
using EB.VLifeModels;
using EB.WebBrokerModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.MerSal.Service
{
    [ServiceContract]
    public interface IMerSalService
    {
        /// <summary>
        /// 取得工作月
        /// </summary>
        /// <param name="strSel">query type</param>
        /// <param name="exclude88">是否排掉 sequence=88</param>
        /// <param name="selectTop">要取得的筆數</param>
        /// <returns></returns>
        [OperationContract]
        List<agym> GetYMData(string strSel, bool exclude88, int selectTop = 1);

        /// <summary>
        /// 手續費複雜檢核及時批次
        /// </summary>
        /// <param name="jsonParams"></param>
        /// <returns></returns>
        [OperationContract]
        Stream BatchQueryReport(string jsonParams);

        /// <summary>
        /// 取得發佣原始檔狀態
        /// </summary>
        /// <param name="ProductionYm"></param>
        /// <param name="Sequence"></param>
        /// <param name="CompanyCode"></param>
        /// <returns></returns>
        [OperationContract]
        string GetSeqNo(string ProductionYm, string Sequence, string CompanyCode);

        /// <summary>
        /// 取得保公
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        List<TermVal> GetCompanyCode();

        /// <summary>
        /// 佣酬檢核狀態為「4-資料入發佣」
        /// </summary>
        /// <param name="ProductionYm"></param>
        /// <param name="Sequence"></param>
        /// <returns></returns>
        [OperationContract]
        string Getprocess_status_MSR(string ProductionYm, string Sequence);

        /// <summary>
        /// 佣酬檢核狀態為「3-資料入佣酬調整
        /// </summary>
        /// <param name="ProductionYm"></param>
        /// <param name="Sequence"></param>
        /// <param name="CompanyCode"></param>
        /// <returns></returns>
        [OperationContract]
        string Getprocess_status_MSRC(string ProductionYm, string Sequence, string CompanyCode);

        /// <summary>
        /// Batch中有狀態為R且時間不超過3分鐘
        /// </summary>
        /// <param name="ProductionYm"></param>
        /// <param name="Sequence"></param>
        /// <param name="CompanyCode"></param>
        /// <returns></returns>
        [OperationContract]
        string Getcm_batch_controlR(string ProductionYm, string Sequence, string CompanyCode);

        /// <summary>
        /// 檢核執行紀錄
        /// </summary>
        /// <param name="ProductionYm"></param>
        /// <param name="Sequence"></param>
        /// <param name="CompanyCode"></param>
        /// <returns></returns>
        [OperationContract]
        List<MerSalViewModel> GetMerSalRun(string ProductionYm, string Sequence, string CompanyCode);

        /// <summary>
        /// 目前人事關檔年月-次薪
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        string GetAgymIndOne();

        /// <summary>
        /// 目前業績關檔年月-次薪
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        string GetAgbcIndOne();

        /// <summary>
        /// 目前工作月
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        string GetProductionYm();

        /// <summary>
        /// 目前次薪
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        string GetSequence();

        /// <summary>
        /// 入佣資料報表
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        [OperationContract]
        byte[] GetMerSalDReportList(MerSalViewModel condition);


        /// <summary>
        /// 保險公司佣酬檢核表
        /// </summary>
        /// <param name="year">年度</param>
        /// <returns>Stream</returns>
        [OperationContract]
        Stream GetCompanyMerSalDReportList(MerSalViewModel condition);

        /// <summary>
        /// 目前業績工作月年月-次薪
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        string GetProductionYmNow();

        /// <summary>
        /// 入Run薪工作月
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        string GetProductionYmRun();

        /// <summary>
        /// 佣酬類別
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        List<Trmval> GetAmountType();

        /// <summary>
        /// 檢核特殊資料類別
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        List<Trmval> GetChkCodeActType();

        ///// <summary>
        ///// 發放註記
        ///// </summary>
        ///// <returns></returns>
        //[OperationContract]
        //List<Trmval> GetPayType();

        /// <summary>
        /// 佣酬發放調整查詢
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        List<MerSalCheckViewModel> GetMerSalCheck(MerSalCheckViewModel model);

        /// <summary>
        /// 更新佣酬發放調整
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [OperationContract]
        bool UpdateMerSalCheck(MerSalCheckViewModel model, string sA1, string sA2, string sA3);

        /// <summary>
        /// 佣酬發放調整報表
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        [OperationContract]
        byte[] GetMerSalCheckReportList(MerSalCheckViewModel condition,string ShowTitle);

        /// <summary>
        /// 原始檔系統保留暨人工調帳查詢
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [OperationContract]
        List<MerSalCutViewModel> GetMerSalCut(MerSalCutViewModel model);

        /// <summary>
        /// 原始檔系統保留暨人工調帳報表
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        [OperationContract]
        byte[] GetMerSalCutReportList(MerSalCutViewModel condition, string FileName);
        /// <summary>
        /// 目前業績未關檔年月-次薪
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        string GetNotYetAgbcInd();
        /// <summary>
        /// 取得保公
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        List<TermVal> GetALLCompanyCode();
        /// <summary>
        /// 佣酬發放調整查詢
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        List<MerSalCheckSPruleViewModel> GetMerSalCheckSPrule(MerSalCheckSPruleViewModel model,string ClickItem);
        /// <summary>
        /// 更新停用狀態
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [OperationContract]
        bool UpdateMerSalCheckSPruleIsDelete(MerSalCheckSPruleViewModel model);

        /// <summary>
        /// 新增MerSalCheckSPrule
        /// </summary>
        /// <param name="dt">data</param>
        [OperationContract]
        string InsertMerSalCheckSPrule(MerSalCheckSPrule model);

        /// <summary>
        /// 檢核特殊資料設定報表
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        [OperationContract]
        byte[] GetMerSalCheckSPruleReportList(MerSalCheckSPruleViewModel condition);

        /// <summary>
        /// 最大業績已關檔年月
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        string GetYmClose();

        /// <summary>
        /// 最大業績已關檔序號
        /// </summary>
        /// <returns></returns>
        [OperationContract]
        string GetSeqClose();
    }
}
