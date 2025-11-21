using EB.CUFModels;
using EB.EBrokerModels;
using EB.SL.PayRoll.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EB.SL.PayRoll.Web.Areas.PayRoll.Utilities
{
    public static class PayRollHelper
    {
        private static IPayRollService _Service = ServiceHelper.Create<IPayRollService>();

        /// <summary>
        /// 取得新的ProcessNo
        /// </summary>
        /// <returns></returns>
        public static string GetNewProcessNo()
        {
            return _Service.GetProcessNo();
        }

        /// <summary>
        /// 該業務員代碼是否存在於該業績年月的Aginb
        /// </summary>
        /// <param name="productionYM">業績年月</param>
        /// <param name="agentCode">業務員代碼</param>
        /// <returns></returns>
        public static string CheckAginb(string productionYM, string agentCode)
        {
            return _Service.GetAgentNameByAginb(productionYM, agentCode);
        }

        /// <summary>
        /// 輸入原因碼取得對應的型態代碼
        /// </summary>
        /// <param name="reasonCode">原因碼</param>
        /// <returns></returns>
        public static string GetTypeByReasonCode(string reasonCode)
        {
            return _Service.GetReasonCodeToAdjTypeMappingValue(reasonCode.ToUpper());
        }

        /// <summary>
        ///  檢核保險公司代碼是否存在
        /// </summary>
        /// <param name="companyCode"></param>
        /// <returns></returns>
        public static bool CheckCompanyCodeIsExist(string companyCode)
        {
            List<TermVal> sets = _Service.GetCompanySetList();

            if (sets.Where(x => x.TermCode.Equals(companyCode)).Any())
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        ///  西元轉民國 0yyy/MM
        /// </summary>
        /// <returns></returns>
        public static string WYToRoc(string date)
        {
            //string sampleDate = "2012-2-29";
            DateTime dt = DateTime.Parse(date);
            CultureInfo culture = new CultureInfo("zh-TW");
            culture.DateTimeFormat.Calendar = new TaiwanCalendar();
            return dt.ToString("0" + "yyy/MM", culture);
        }

        /// <summary>
        ///  民國轉西元 yyyy/MM
        /// </summary>
        /// <returns></returns>
        public static string RocToWY(string date)
        {
            //string sampleDate = "101/02/29";
            //CultureInfo culture = new CultureInfo("zh-TW");
            //culture.DateTimeFormat.Calendar = new TaiwanCalendar();
            //date = DateTime.Parse(date, culture).ToString();
            string[] sArray = date.Split('/');
            string newDate = (Convert.ToInt32(sArray[0]) + 1911).ToString();
            newDate = newDate + "/" + sArray[1];
            return newDate;
        }

        /// <summary>
        /// 取得部門別清單產生選單項目List
        /// </summary>
        /// <param name="display"></param>
        /// <returns></returns>
        public static IEnumerable<SelectListItem> QuerynunitListItem(string MemberID, Func<sc_unit, string> display = null)
        {
            var result = new List<SelectListItem>();         
            var service = ServiceHelper.Create<IPayRollService>();
            var dataList = service.Getnunit();
            var unitID = ChangeUnitID(MemberID);    
            result.AddRange(dataList.Select(m => new SelectListItem { Value = m.UnitId.ToString(), Text = display(m) }).Where(m => m.Value == unitID));           
            display = display ?? ((sc_unit m) => m.UnitName);
            
            return result;
        }

        /// <summary>
        /// 取得部門別清單產生選單項目List
        /// </summary>
        /// <param name="display"></param>
        /// <returns></returns>
        public static IEnumerable<SelectListItem> QueryMainnunitListItem(string MemberID, Func<sc_unit, string> display = null)
        {
            var result = new List<SelectListItem>();
            var service = ServiceHelper.Create<IPayRollService>();
            var dataList = service.Getnunit();
            var unitID = ChangeUnitID(MemberID);
            result.Add(new SelectListItem { Value = "", Text = "全公司" });
            result.AddRange(dataList.Select(m => new SelectListItem { Value = m.UnitId.ToString(), Text = display(m) }));
            display = display ?? ((sc_unit m) => m.UnitName);

            return result;
        }

        /// <summary>
        /// 科別ID轉部門ID
        /// </summary>
        /// <param name="display"></param>
        /// <returns></returns>
        public static string ChangeUnitID(string MemberID)
        {
            string depart;
            var unit = Member.Get(MemberID).GetUnit();
            var service = ServiceHelper.Create<IPayRollService>();
            //部級
            var unitParent = unit.GetParent();
            int level = service.GetUnitlevel(unit.ID);

            //主管級
            if (level < 4)
            {
                depart = unit.ID;
            }
            else
            {
                depart = unitParent.ID;
            }
            return depart;
        }

        /// <summary>
        /// 科別IDName轉部門Name
        /// </summary>
        /// <param name="display"></param>
        /// <returns></returns>
        public static string ChangeUnitName(string MemberID)
        {
            string depart;
            var unit = Member.Get(MemberID).GetUnit();
            var service = ServiceHelper.Create<IPayRollService>();
            //部級
            var unitParent = unit.GetParent();
            int level = service.GetUnitlevel(unit.ID);

            //主管級
            if (level < 4)
            {
                depart = unit.Name;
            }
            else
            {
                depart = unitParent.Name;
            }
            return depart;
        }

        /// <summary>
        /// 絕對值
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static int ABS(int value)
        {
            return Math.Abs(value);
        }
    }
}