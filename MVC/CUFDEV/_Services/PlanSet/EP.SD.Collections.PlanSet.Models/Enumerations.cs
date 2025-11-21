/// 需求單號：20250707001 競賽計C商品清單。 2025/07/03 BY Harrison 
using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.Collections.PlanSet
{
    /// <summary>
    /// 競賽計C-商品清單
    /// </summary>
    public enum ProductList
    {
        [Display(Name = "全部(含維護中)")]
        All = 0,

        [Display(Name = "不計入")]
        N = 1,

        [Display(Name = "計入")]
        Y = 2,

        [Display(Name = "維護中")]
        Maintenance = 3,
    }
}
