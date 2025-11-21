using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesZone.AGUPG.Models
{
    public class HrUpg25Dto : IModel
    {
        /// <summary>
        /// 業績、組織適用標準（達成其一即可）
        /// </summary>
        //[DisplayName("一、業績、組織適用標準（達成其一即可）")]
        //[Column("PlanSetArea1", IsKey = true, IsIdentity = true)]
        public List<HrUpg25RstGrid1> HrUpg25RstGrid1 { get; set; }

        /// <summary>
        /// 其他必要條件
        /// </summary>
        //[DisplayName("二、其他必要條件")]
        //[Column("PlanSetArea1", IsKey = true, IsIdentity = true)]
        public List<HrUpg25RstGrid2> HrUpg25RstGrid2 { get; set; }

        /// <summary>
        /// 其他說明（參考資訊）
        /// </summary>
        //[DisplayName("三、其他說明（參考資訊）")]
        //[Column("PlanSetArea1", IsKey = true, IsIdentity = true)]
        public List<HrUpg25RstGrid3> HrUpg25RstGrid3 { get; set; }

        /// <summary>
        /// 業務員資訊
        /// </summary>
        //[DisplayName("晉級")]
        //[Column("PlanSetArea1", IsKey = true, IsIdentity = true)]
        public HrUpg25RstTitle HrUpg25RstTitle { get; set; }
    }
}
