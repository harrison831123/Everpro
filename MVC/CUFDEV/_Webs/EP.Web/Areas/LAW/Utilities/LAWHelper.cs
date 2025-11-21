using EP.Platform.Service;
using EP.SD.SalesSupport.LAW.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.SalesSupport.LAW.Web.Areas.LAW.Utilities
{
    public class LAWHelper
    {
        /// <summary>
        /// 科別ID轉部門ID
        /// </summary>
        /// <param name="display"></param>
        /// <returns></returns>
        public static string ChangeUnitID(string MemberID)
        {
            string depart;
            var unit = Member.Get(MemberID).GetUnit();
            var service = ServiceHelper.Create<ILAWService>();
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
            var service = ServiceHelper.Create<ILAWService>();
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

    }
}