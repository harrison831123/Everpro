using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using static EP.SD.SalesZone.AGUPG.Models.Enumerations;

namespace EP.SD.SalesZone.AGUPG.Web.Areas.AGUPG.Utilities
{
    public static class AGUPGHelper
    {
        /// <summary>
        /// 判斷是否為主管
        /// </summary>
        /// <param name="user"></param>
        /// <param name="programID"></param>
        /// <returns></returns>
        public static bool IsAdmin(MicrosoftPrincipal user, string programID)
        {
            var type = GetUserType(user, programID);
            return type == AGUPGUserType.Admin;
        }

        /// 取得系統的使用者類別
        /// </summary>
        /// <param name="user">MicrosoftPrincipal</param>
        /// <param name="programID">程式代碼</param>
        /// <returns></returns>
        public static AGUPGUserType GetUserType(MicrosoftPrincipal user, string programID)
        {
            var admin = $"EP.SD.SalesZone.AGUPG.{programID}.Admin";
            //var preadmin = $"EP.SD.SalesZone.AGUPG.{programID}.PreAdmin";
            var agent = $"EP.SD.SalesZone.AGUPG.{programID}.User";

            if (user.HasPermission(admin))
                return AGUPGUserType.Admin;

            //elseif (user.HasPermission(preadmin))
            //    return AGUPGUserType.PRE_Admin;

            return AGUPGUserType.User;
        }
    }
}