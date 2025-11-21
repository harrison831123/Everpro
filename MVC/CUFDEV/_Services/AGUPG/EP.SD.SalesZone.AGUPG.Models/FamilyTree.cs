using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesZone.AGUPG.Models
{
    public class FamilyTree : IModel
    {
        /// <summary>
        /// 代數
        /// </summary>
        [DisplayName("代數")]
        [Column("MgNo")]
        public string MgNo { get; set; }

        /// <summary>
        /// 業務姓名
        /// </summary>
        [DisplayName("業務姓名")]
        [Column("agent_name")]
        public string AgentName { get; set; }

        /// <summary>
        /// 職級名稱
        /// </summary>
        [DisplayName("職級名稱")]
        [Column("AgLevelName")]
        public string AgLevelName { get; set; }

        /// <summary>
        /// 晉升結果
        /// </summary>
        [DisplayName("晉升結果")]
        [Column("Rst_Upgrade")]
        public string RstUpgrade { get; set; }

        /// <summary>
        /// 業務代號
        /// </summary>
        [DisplayName("業務代號")]
        [Column("agent_code")]
        public string AgentCode { get; set; }

        /// <summary>
        /// 狀態日期
        /// </summary>
        [DisplayName("狀態日期")]
        [Column("ag_status_date")]
        public string AgStatusDate { get; set; }

        /// <summary>
        /// 狀態代碼
        /// </summary>
        [DisplayName("狀態代碼")]
        [Column("ag_status_code")]
        public string AgStatusCode { get; set; }

        /// <summary>
        /// 報聘日期
        /// </summary>
        [DisplayName("報聘日期")]
        [Column("register_date")]
        public string RegisterDate { get; set; }

    }
}
