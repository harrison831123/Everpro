using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EB.SL.MerSal.Web.Areas.MerSal.Utilities
{
    public class MerSalHelper
    {
        /// <summary>
        ///  民國轉西元 yyyy/MM
        /// </summary>
        /// <returns></returns>
        public static string RocToWY(string date)
        {
            string[] sArray = date.Split('/');
            string newDate = (Convert.ToInt32(sArray[0]) + 1911).ToString();
            newDate = newDate + "/" + sArray[1];
            return newDate;
        }
    }
}