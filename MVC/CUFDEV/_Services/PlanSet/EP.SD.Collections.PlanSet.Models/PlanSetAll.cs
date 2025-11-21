using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.Collections.PlanSet.Models
{
    /// <summary>
    /// 需求單號：20241210003 現售商品佣酬資料的Model 2024.12 BY VITA
    /// </summary>
    public class PlanSetAll : IModel
    {
        [DisplayName("初年度業績換算率(業行部)")]
        [Column("PlanSetArea1", IsKey = true, IsIdentity = true)]
        public List<PlanSetArea1> PlanSetArea1 { get; set; }

        [DisplayName("保公獎勵內容(業行部)")]
        [Column("PlanSetArea1", IsKey = true, IsIdentity = true)]
        public List<PlanSetArea2> PlanSetArea2 { get; set; }

        [DisplayName("永達競賽獎勵(業支部)")]
        [Column("PlanSetArea1", IsKey = true, IsIdentity = true)]
        public List<PlanSetArea3> PlanSetArea3 { get; set; }

        [DisplayName("表頭")]
        [Column("PlanTitle", IsKey = true, IsIdentity = true)]
        public PlanTitle PlanTitle { get; set; }
    }
}