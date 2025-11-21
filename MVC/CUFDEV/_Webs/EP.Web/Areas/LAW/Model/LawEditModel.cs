using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Model
{
    public class LawEditModel
    {

        /// <summary>
        /// 流水號
        /// </summary>
        public int ID { get; set; }

        /// <summary>
        /// 主檔流水號
        /// </summary>
        public int LawId { get; set; }

        /// <summary>
        /// 照會單號
        /// </summary>
        public string LawNoteNo { get; set; }

        /// <summary>
        /// 業務員ID 10馬
        /// </summary>
        public string AgentID { get; set; }

        /// <summary>
        /// 修改頁面狀態
        /// 1:存證信函、2:訴訟程序、 3:執行程序
        /// </summary>
        public string EditViewType { get; set; }

        /// <summary>
        /// 編輯者
        /// </summary>
        [DisplayName("編輯者")]
        public string EditUser { get; set; }

        /// <summary>
        /// 時間
        /// </summary>
        [DisplayName("時間")]
        public string CreateTime { get; set; }

        /// <summary>存證信函備註</summary>
        [DisplayName("存證信函備註")]
        public string LawEvidencedesc { get; set; }

        /// <summary>訴訟程序內容</summary>
        [DisplayName("訴訟程序內容")]
        public string LawLitigationprogress { get; set; }

        /// <summary>執行程序說明</summary>
        [DisplayName("執行程序說明")]
        public string LawDoprogress { get; set; }

        [DisplayName("清償金額")]
        public int LawRepaymentMoney { get; set; }

        /// <summary>其他備註說明</summary>
        [DisplayName("說明")]
        public string LawOtherdesc { get; set; }

        [DisplayName("業務員id")]//12馬
        public string LawDueAgentId { get; set; }

        /// <summary>備註</summary>
        [DisplayName("備註")]
        public string LawDescdesc { get; set; }

        /// <summary>續佣扣抵</summary>
        [DisplayName("續佣扣抵")]
        public int LawCommDeduction { get; set; }
    }
}