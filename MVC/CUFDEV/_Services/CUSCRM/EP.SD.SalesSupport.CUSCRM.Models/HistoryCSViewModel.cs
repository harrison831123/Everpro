using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;


namespace EP.SD.SalesSupport.CUSCRM
{
    public class HistoryCSViewModel
    {
        public HistoryCSViewModel()
        {

        }

        #region 未顯示資料欄位
        [Display(Name = "保公代號")]
        public string company_code { get; set; }
        [Display(Name = "受理日期")]
        public string crm_no_createdate { get; set; }

        [Display(Name = "實駐")]
        public string now_wc_centername { get; set; }



        [Display(Name = "服務案結案情形")]
        public string crm_close { get; set; }
        #endregion

        #region 顯示欄位
        [Display(Name = "序號")]
        public int sNo { get; set; }

        [Display(Name = "處理單號")]
        public string crm_no { get; set; }
        [Display(Name = "來源")]
        public string crm_source { get; set; }
        [Display(Name = "承辦人")]

        public string crm_douser { get; set; }
        [Display(Name = "保險公司")]

        public string company_name { get; set; }
        [Display(Name = "要保人")]
        public string owner { get; set; }
        [Display(Name = "被保人")]
        public string insured { get; set; }
        [Display(Name = "生效日")]
        public string issdate { get; set; }

        [Display(Name = "保單號碼")]
        public string policy_no2 { get; set; }
        [Display(Name = "類別")]
        public int crm_dotype { get; set; }
        [Display(Name = "服務摘要")]
        public string crm_content { get; set; }
        [Display(Name = "副總體系")]
        public string now_vmname { get; set; }
        [Display(Name = "協理體系")]
        public string now_smname { get; set; }


        [Display(Name = "原經手人")]
        public string orgid_agname { get; set; }
        [Display(Name = "現經手人")]

        public string now_agname { get; set; }
        [Display(Name = "接續經手人")]
        public string sub_agname { get; set; }

        [Display(Name = "實際結案日")]
        public string crm_closedate { get; set; }

        [Display(Name = "備註")]

        public string crm_closedesc { get; set; }


        #region 衍生欄位

        [Display(Name = "生效年度")]

        public string issdate_year
        {
            get
            {
                return (String.IsNullOrWhiteSpace(issdate)) ? "" : int.Parse(issdate.Substring(0, 4)).ToString();
            }
        }
        [Display(Name = "受理年")]
        public string crm_no_createdate_year
        {
            get
            {
                return String.IsNullOrEmpty(crm_no_createdate) ? "" : (int.Parse(crm_no_createdate.Substring(0, 4)) - 1911).ToString();
            }
        }

        [Display(Name = "受理月")]
        public string crm_no_createdate_month
        {
            get
            {
                return String.IsNullOrEmpty(crm_no_createdate) ? "" : crm_no_createdate.Substring(5, 2);
            }
        }
        [Display(Name = "受理日")]
        public string crm_no_createdate_date
        {
            get
            {
                return String.IsNullOrEmpty(crm_no_createdate) ? "" : crm_no_createdate.Substring(8, 2);
            }
        }
        [Display(Name = "受理天數")]
        public string crm_no_createdate_day
        {
            get
            {
                if (String.IsNullOrEmpty(crm_no_createdate))
                    return "";
                TimeSpan span = DateTime.Now.Subtract(DateTime.Parse(crm_no_createdate));
                return ((int)span.TotalDays + 1).ToString();
            }
        }
        [Display(Name = "結案天數")]
        public string crm_closedate_day
        {
            get
            {
                if (String.IsNullOrEmpty(crm_closedate) || String.IsNullOrEmpty(crm_no_createdate))
                    return "";
                TimeSpan span = DateTime.Parse(crm_closedate).Subtract(DateTime.Parse(crm_no_createdate));
                return ((int)span.TotalDays + 1).ToString();
            }
        }
        [Display(Name = "單位")]
        public string now_wc_centername_unit
        {
            get
            {
                if (String.IsNullOrEmpty(now_wc_centername))
                    return "";
                else  //處理沒有含有"(",")"實駐 ex:屏東
                    if (now_wc_centername.Contains("("))
                    return now_wc_centername.Substring(now_wc_centername.IndexOf('(') + 1, now_wc_centername.IndexOf(')') - now_wc_centername.IndexOf('(') - 1);
                else
                    return now_wc_centername;
            }
        }
        [Display(Name = "處別")]
        public string now_wc_centername_div
        {
            get
            {
                if (String.IsNullOrEmpty(now_wc_centername))
                    return "";
                else  //處理沒有含有"(",")"實駐 ex:屏東
                    if (now_wc_centername.Contains("("))
                    return now_wc_centername.Substring(0, now_wc_centername.IndexOf('('));
                else
                    return now_wc_centername;
            }
        }


        [Display(Name = "服務案結案情形")]
        public string crm_close_str
        {
            get
            {
                if (String.IsNullOrEmpty(crm_close))
                    return "";
                switch (crm_close)
                {
                    case "16":
                        return "A";
                    case "22":
                        return "A-IB";
                    case "17":
                        return "B";
                    case "23":
                        return "B-IB";
                    case "19":
                        return "C";
                    case "24":
                        return "C-IB";
                    case "20":
                        return "D";
                    case "21":
                        return "E";
                    case "25":
                        return "E-B";
                    case "83":
                        return "F";
                    case "46":
                        return "婉拒";
                    case "47":
                        return "融通";
                    case "53":
                        return "其他";

                    default:
                        return "";
                }
            }
        }
        [Display(Name = "類別")]
        public string crm_dotype_str
        {
            get
            {
                if (crm_dotype == 0)
                    return "";
                switch (crm_dotype)
                {
                    case 8:
                        return "服務品質";
                    case 9:
                        return "銷售爭議";
                    case 10:
                        return "簽名爭議";
                    case 11:
                        return "理賠爭議";
                    case 12:
                        return "銷售&簽名爭議";
                    case 35:
                        return "契撤";
                    case 36:
                        return "自始契變";
                    case 37:
                        return "年期轉換";
                    case 38:
                        return "更換業代";
                    case 39:
                        return "解約";
                    case 40:
                        return "一般";
                    case 41:
                        return "抱怨";
                    case 42:
                        return "融通契變";
                    case 43:
                        return "試算";
                    case 49:
                        return "其他";
                    case 54:
                        return "保費";
                    case 55:
                        return "理賠";
                    case 56:
                        return "恢復效力";
                    case 57:
                        return "降額繳清展期";
                    case 87:
                        return "契變";
                    default:
                        return "";
                }
            }
        }
        [Display(Name = "來源")]
        public string crm_source_str {
            get {
                if (String.IsNullOrEmpty(crm_source))
                    return "";
                switch (crm_source) {
                    case "13":
                        return "保險公司照會";
                    case "14":
                        return "申訴函";
                    case "15":
                        return "保險局";
                    case "34":
                        return "單位業連";
                    case "60":
                        return "保戶親電";
                    case "61":
                        return "公司網站";
                    case "66":
                        return "保險局";
                    case "67":
                        return "傳真";
                    case "68":
                        return "保戶家屬來電";
                    case "70":
                        return "保戶郵寄";
                    case "71":
                        return "單位來電";
                    case "74":
                        return "其他";
                    case "75":
                        return "單位來函";
                    case "76":
                        return "公司網站";
                    case "81":
                        return "保險公司CSR";
                    case "82":
                        return "保險公司來電";
                    case "84":
                        return "金評中心系統照會";
                    case "85":
                        return "其他";

                    default:
                        return "";
                }
            }
        }
        [Display(Name = "是否到保險局申訴")]
        public string crm_source15
        {
            get {
                if (this.crm_source == "15")
                    return "V";
                return "";
            }
        }

        [Display(Name = "是否到立法院申訴")]
        public string crm_source30
        {
            get
            {
                if (this.crm_source == "30")
                    return "V";
                return "";
            }
        }
        [Display(Name = "是否到北市府申訴")]
        public string crm_source26
        {
            get
            {
                if (this.crm_source == "26")
                    return "V";
                return "";
            }
        }
        [Display(Name = "是否到北縣府申訴")]
        public string crm_source27
        {
            get
            {
                if (this.crm_source == "27")
                    return "V";
                return "";
            }
        }

        #endregion

        #region 按鈕欄位
        [Display(Name = "維護紀錄")]
        public string MaintainBox { get; set; }

        [Display(Name = "稽催紀錄")]
        public string UrgeBox { get; set; }
        [Display(Name = "匯出處理單")]
        public string ExportBox { get; set; }

        #endregion

        #endregion

        //a.crm_status,
        //a.crm_dotype,
        //        convert(char(10), b.crm_no_createdate, 111) as crm_no_createdate,
        //        e.vm_group_name,
        //        e.sm_name,
        //        e.wc_center_name,

        //        e.name as ,
        //        e.vm_group_name as vmname,
        //        e.sm_name as smname,
        //        e.center_name as wc_centername,
        //        f.wc_center_name as now_wc_centername,

    }
}