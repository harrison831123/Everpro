using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesZone.AGUPG.Models
{
    public class FamilyBoss : IModel
    {
        /// <summary>
        /// 職級
        /// </summary>
        [DisplayName("職級")]
        [Column("LEVEL_1")]
        public string Level1 { get; set; }

        /// <summary>
        /// 業務代號
        /// </summary>
        [DisplayName("業務代號")]
        [Column("agent_code")]
        public string AgentCode { get; set; }

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

    }
}
