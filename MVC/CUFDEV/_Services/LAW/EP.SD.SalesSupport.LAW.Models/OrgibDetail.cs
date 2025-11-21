using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.LAW.Models
{
    public class OrgibDetail : IModel
    {
        [Column("sm_code")]
        public string sm_code { get; set; }

        [Column("sm_name")]
        public string sm_name { get; set; }

        [Column("wc_center")]
        public string wc_center { get; set; }

        [Column("wc_center_name")]
        public string wc_center_name { get; set; }

        [Column("center_code")]
        public string center_code { get; set; }

        [Column("center_name")]
        public string center_name { get; set; }

        [Column("administrat_id")]
        public string administrat_id { get; set; }

        [Column("admin_name")]
        public string admin_name { get; set; }

        [Column("admin_level")]
        public string admin_level { get; set; }

        [Column("agent_code")]
        public string agent_code { get; set; }

        [Column("names")]
        public string names { get; set; }

        [Column("ag_status_code")]
        public string ag_status_code { get; set; }

        [Column("ag_level")]
        public string ag_level { get; set; }

        [Column("level_name_chs")]
        public string level_name_chs { get; set; }

        [Column("birth")]
        public string birth { get; set; }

        [Column("cellur_phone_no")]
        public string cellur_phone_no { get; set; }

        [Column("record_date")]
        public string record_date { get; set; }

        [Column("register_date")]
        public string register_date { get; set; }

        [Column("ag_status_date")]
        public string ag_status_date { get; set; }
    }
}
