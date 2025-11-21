using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesZone.AGUPG.Models
{
    /// <summary>
    /// 報名權限狀態
    /// </summary>
    [DataContract]
    public class Enumerations
    {
        /// <summary>
        /// 業務員身份別
        /// </summary>
        public enum AGUPGUserType
        {
            /// <summary>業務主管</summary>
            Admin,
            /// <summary>籌備業務主管</summary>
            PreAdmin,
            /// <summary>業務員</summary>
            User,
            /// <summary>處代理人</summary>
            //UnitAgent
        }
    }
}
