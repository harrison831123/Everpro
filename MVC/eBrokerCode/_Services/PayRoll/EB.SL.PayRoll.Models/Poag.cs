using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.PayRoll.Models
{
	public class Poag : IModel
    {
        /// <summary>
        /// 保單號碼
        /// </summary>
        [Column("policy_no2")]
        [DisplayName("保單號碼")]
        public string PolicyNo2 { get; set; }


        /// <summary>
        /// 業務員代碼
        /// </summary>
        [Column("agent_code")]
        [DisplayName("業務員代碼")]
        public string AgentCode { get; set; }
    }
}
