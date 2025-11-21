//需求單號：20250707001 競賽計C商品清單。 2025/07/03 BY Harrison 
using EP.Platform.Service;
using EP.SD.Collections.PlanSet.Service;
using Microsoft.CUF.Framework.Service;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.Collections.PlanSet.Web
{
    public class PlanSetHelper
    {
        public static IPlanSetService _service = ServiceHelper.Create<IPlanSetService>();
                
        /// <summary>
        /// 取得保公下拉選單
        /// 預設"全部"
        /// </summary>
        /// <returns></returns>
        public static List<SelectListItem> GetCompanyCode()
        {
            List<ValueText> list = _service.GetCompanyCode();

            List<SelectListItem> items = new List<SelectListItem>();
            items.Add(new SelectListItem() { Text = "全部", Value = "0" });

            foreach (var data in list)
            {
                items.Add(new SelectListItem() { Text = data.Text, Value = data.Value });
            }
            return items;
        }

        /// <summary>
        /// 取得 [Display(Name = "")] 的文字
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string GetDisplayName(Enum value)
        {
            var field = value.GetType().GetField(value.ToString());
            var attribute = field.GetCustomAttributes(typeof(DisplayAttribute), false)
                                 .Cast<DisplayAttribute>()
                                 .FirstOrDefault();
            return attribute?.Name ?? value.ToString();
        }
    }
}