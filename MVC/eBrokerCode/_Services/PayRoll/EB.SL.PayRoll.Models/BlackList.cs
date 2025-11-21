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
    public class BlackList : IModel
    {
        /// <summary>
        /// 業務員ID
        /// </summary>
        [Column("agent_code")]
        public string AgentCode { get; set; }

        /// <summary>
        /// 原因馬
        /// </summary>
        [Column("reason_code")]
        public string ReasonCode { get; set; }
    }
}
