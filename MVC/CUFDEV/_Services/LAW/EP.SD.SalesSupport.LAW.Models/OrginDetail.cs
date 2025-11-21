using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.SD.SalesSupport.LAW.Models
{
    public class OrginDetail : IModel
    {
        [Column("center_name")]
        public string center_name { get; set; }

        [Column("ag_status_code")]
        public string ag_status_code { get; set; }

        [Column("name")]
        public string name { get; set; }

        [Column("director_id")]
        public string director_id { get; set; }

        [Column("director_name")]
        public string director_name { get; set; }

        [Column("dir_status_code")]
        public string dir_status_code { get; set; }

        [Column("level_name_chs")]
        public string level_name_chs { get; set; }

        [Column("term_code")]
        public string term_code { get; set; }

        [Column("term_meaning")]
        public string term_meaning { get; set; }

    }
}
