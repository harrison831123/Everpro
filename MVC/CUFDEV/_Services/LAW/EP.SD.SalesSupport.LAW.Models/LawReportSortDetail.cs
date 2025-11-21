using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.LAW.Models
{
    public class LawReportSortDetail : IModel
    {
        /// <summary>排序年度</summary>
        [Column("year")]
        public string year { get; set; }

        /// <summary>團隊</summary>
        [Column("vm_name")]
        public string vm_name { get; set; }

        /// <summary>體系</summary>
        [Column("sm_name")]
        public string sm_name { get; set; }

        /// <summary>處</summary>
        [Column("center_name")]
        public string center_name { get; set; }

    }
}
