using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.CUSCRM.Service
{
    public class BrokerInfo
    {
        /// <summary>
        /// 原招攬人
        /// </summary>
        [Display(Name = "原招攬人")]
        public string AgentName { get; set; }

        /// <summary>
        /// 壽險初登錄日
        /// </summary>
        [Display(Name = "壽險初登錄日")]
        public string ExRecordDate { get; set; }

        /// <summary>
        /// 壽險登錄日
        /// </summary>
        [Display(Name = "壽險登錄日")]
        public string RecordDate { get; set; }

        /// <summary>
        /// 簽約日
        /// </summary>
        [Display(Name = "簽約日")]
        public string RegisterDate { get; set; }

        /// <summary>
        /// 外幣登錄日
        /// </summary>
        [Display(Name = "外幣登錄日")]
        public string FRegDate { get; set; }

        /// <summary>
        /// 任用狀況碼
        /// </summary>
        [Display(Name = "任用狀況碼")]
        public string AgStatusCode { get; set; }

        /// <summary>
        /// 任用狀況日
        /// </summary>
        [Display(Name = "任用狀況日")]
        public string AgStatusDate { get; set; }

        /// <summary>
        /// 原招攬處
        /// </summary>
        [Display(Name = "原招攬處")]
        public string CenterName { get; set; }

        /// <summary>
        /// 接續服務人員
        /// </summary>
        [Display(Name = "接續服務人員")]
        public string SUName { get; set; }

        /// <summary>
        /// 接續服務人員壽險初登錄日
        /// </summary>
        [Display(Name = "壽險初登錄日")]
        public string SUExRecordDate { get; set; }

        /// <summary>
        /// 接續服務人員壽險登錄日
        /// </summary>
        [Display(Name = "壽險登錄日")]
        public string SURecordDate { get; set; }

        /// <summary>
        /// 接續服務人員簽約日
        /// </summary>
        [Display(Name = "簽約日")]
        public string SURegisterDate { get; set; }

    }
}
