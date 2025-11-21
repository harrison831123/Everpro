using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesZone.AGUPG.Models
{
    public class HrUpgGet25Dto
    {
        /// <summary>
        /// 轄下第一代區級
        /// </summary>
        public List<HrUpgGet25Detail1> HrUpgGet25Detail1 { get; set; }

        /// <summary>
        /// 連續四季明細
        /// </summary>
        public List<HrUpgGet25Detail2> HrUpgGet25Detail2 { get; set; }

        /// <summary>
        /// 加計被推介人業績
        /// </summary>
        public List<HrUpgGet25Detail3> HrUpgGet25Detail3 { get; set; }

        /// <summary>
        /// 遞延未核實
        /// </summary>
        public List<HrUpgGet25Detail4> HrUpgGet25Detail4 { get; set; }
    }
}
