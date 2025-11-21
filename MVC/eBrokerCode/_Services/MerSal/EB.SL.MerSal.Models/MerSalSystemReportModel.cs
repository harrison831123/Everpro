using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.MerSal.Models
{
    public class MerSalSystemReportModel
    {
        public string amount_type { get; set; }
        public string modx_year_type { get; set; }
        public string modx_year_name { get; set; }
        public string Cnt_Cut00 { get; set; }
        public string ModePrem_Cut00 { get; set; }
        public string CommPrem_Cut00 { get; set; }
        public string Cnt_Cut01 { get; set; }
        public string ModePrem_Cut01 { get; set; }
        public string CommPrem_Cut01 { get; set; }
        public string Cnt_Check { get; set; }
        public string ModePrem_Check { get; set; }
        public string CommPrem_Check { get; set; }

        public string Cnt_N { get; set; }
        public string ModePrem_N { get; set; }
        public string CommPrem_N { get; set; }


        public string Cnt_Total { get; set; }
        public string ModePrem_Total { get; set; }
        public string CommPrem_Total { get; set; }
    }
}
