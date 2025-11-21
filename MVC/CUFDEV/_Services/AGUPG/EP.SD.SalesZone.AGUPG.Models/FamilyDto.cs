using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesZone.AGUPG.Models
{
    public class FamilyDto
    {
        /// <summary>
        /// 轄下業務資料
        /// </summary>
        public List<FamilyTree> FamilyTree { get; set; }

        /// <summary>
        /// 主管明細資料
        /// </summary>
        public List<FamilyBoss> FamilyBoss { get; set; }

        /// <summary>
        /// 人力資料
        /// </summary>
        public string AgData { get; set; }
    }
}
