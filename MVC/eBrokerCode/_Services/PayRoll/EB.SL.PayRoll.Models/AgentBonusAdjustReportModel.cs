///========================================================================================== 
/// 程式名稱：人工調帳
/// 建立人員：Harrison
/// 建立日期：2022/07
/// 修改記錄：（[需求單號]、[修改內容]、日期、人員）
/// 需求單號:20240122004-因現有VLIFE系統(核心系統)使用已長達20多年，架構老舊，已不敷使用，且為提升資訊安全等級，故計劃執行VLIFE系統改版(新核心系統:eBroker系統)。; 修改內容:上線; 修改日期:20240613; 修改人員:Harrison;
/// 需求單號:20240807001-調整人工調帳系統產出之相關畫面及報表修改等功能。; 修改日期:20240807; 修改人員:Harrison;
///==========================================================================================
using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Microsoft.CUF.Framework.Data;

namespace EB.SL.PayRoll.Models
{
    public class AgentBonusAdjustReportModel : IModel
    {
        /// <summary>
        /// 業績年月-序號
        /// </summary>
        public string ProductionYMSequence { get; set; }

        /// <summary>
        /// 調整類別
        /// </summary>
        public string AdjType { get; set; }

        /// <summary>
        /// 業務員姓名
        /// </summary>
        public string AgentName { get; set; }

        /// <summary>
        /// 業務員代碼
        /// </summary>
        public string AgentCode { get; set; }

        /// <summary>
        /// 原因碼
        /// </summary>
        public string ReasonCode { get; set; }

        /// <summary>
        /// 保單號碼
        /// </summary>
        public string PolicyNo2 { get; set; }

        /// <summary>
        /// 金額
        /// </summary>
        public int Amount { get; set; }

        /// <summary>
        /// FYC
        /// </summary>
        public int FYC { get; set; }

        /// <summary>
        /// FYP
        /// </summary>
        public int FYP { get; set; }

        /// <summary>
        /// 金額-千分位
        /// </summary>
        public string AmountTS { get; set; }

        /// <summary>
        /// FYC-千分位
        /// </summary>
        public string FYCTS { get; set; }

        /// <summary>
        /// FYP-千分位
        /// </summary>
        public string FYPTS { get; set; }

        /// <summary>
        /// 計佣保費
        /// </summary>
        public int CommModePrem { get; set; }

        /// <summary>
        /// 繳別繳次
        /// </summary>
        public string ModxSequence { get; set; }

        /// <summary>
        /// 說明內容
        /// </summary>
        public string DescContent { get; set; }

        /// <summary>
        /// 部門別
        /// </summary>
        public string nunit { get; set; }

        /// <summary>
        /// 名字
        /// </summary>
        public string nmember { get; set; }

        /// <summary>
        /// 保險公司代碼
        /// </summary>
        public string CompanyCode { get; set; }

        /// <summary>
        /// 資料建檔時間
        /// </summary>
        public string CreateDatetime { get; set; }

    }
}
