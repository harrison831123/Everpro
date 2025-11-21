using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;


namespace EP.SD.SalesSupport.LAW.Models
{
    public class OrgVm : IModel
    {
        [DataMember]
        [Column("vm_code")]
        public string vmcode { get; set; }

        [DataMember]
        [Column("vm_name")]
        public string vmname { get; set; }

        [DataMember]
        [Column("vm_leader_id")]
        public string vmleaderid { get; set; }

        [DataMember]
        [Column("virtual_vm")]
        public string virtualvm { get; set; }

        [DataMember]
        [Column("sm_code")]
        public string smcode { get; set; }

        [NonColumn]
        public string smstr { get; set; }

        [NonColumn]
        public int vm_flag { get; set; }

        [NonColumn]
        public int vsm_flag { get; set; }       
    }
}
