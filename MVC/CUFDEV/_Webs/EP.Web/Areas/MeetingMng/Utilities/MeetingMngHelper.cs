using EP.H2OModels;
using EP.Platform.Service;
using EP.PSL.WorkResources.MeetingMng.Service;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.PSL.WorkResources.MeetingMng.Web
{
    public static class MeetingMngHelper
    {
        /// <summary>
        /// 取得會議地點清單產生選單項目List
        /// </summary>
        /// <param name="display"></param>
        /// <returns></returns>
        public static IEnumerable<SelectListItem> QueryMeetingPlaceListItem(Func<Meeting, string> display = null)
        {
            var result = new List<SelectListItem>();
            result.Add(new SelectListItem { Value = "", Text = EP.IBResources.請選擇 });
            var service = ServiceHelper.Create<IMeetingMngService>();
            var dataList = service.GetMeetingPlace();
            display = display ?? ((Meeting m) => m.MTPlace);
            result.AddRange(dataList.Select(m => new SelectListItem { Text = display(m) }));
            return result;
        }


        /// <summary>
        /// 取得會議地點(設備與場地)清單產生選單項目List
        /// </summary>
        /// <param name="display"></param>
        /// <returns></returns>
        public static IEnumerable<SelectListItem> QueryMeetingDevicePlaceListItem(Func<MeetingDevice, string> display = null)
        {
            var result = new List<SelectListItem>();
            result.Add(new SelectListItem { Value = "0", Text = EP.IBResources.請選擇 });
            var service = ServiceHelper.Create<IMeetingMngService>();
            var dataList = service.GetMeetingDevice();
            display = display ?? ((MeetingDevice m) => m.MDName);
            result.AddRange(dataList.Select(m => new SelectListItem { Value=m.MDID.ToString(),Text = display(m) }));
            return result;
        }



        /// <summary>
        /// 取得會議場地與設備List
        /// </summary>
        public static List<MeetingDevice> GetMeetingDeviceList()
        {
            List<MeetingDevice> result = new List<MeetingDevice>();
            var mService = new WebChannel<IMeetingMngService>();
            mService.Use(service => result = service.GetMeetingDevice().ToList());

            return result;
        }

        /// <summary>
        /// 取得會議修改頁面場地與設備List
        /// </summary>
        //public static List<MeetingDevice> GetEditMeetingDeviceList(int mtid)
        //{
        //    List<MeetingDevice> result = new List<MeetingDevice>();
        //    var mService = new WebChannel<IMeetingMngService>();
        //    mService.Use(service => result = service.GetEditMeetingDevice(mtid).ToList());

        //    return result;
        //}

        /// <summary>
        /// 產生會議類別的下拉選單
        /// </summary>
        public static IEnumerable<SelectListItem> GetMeetingDefaultTypeList(string initval, Func<ValueText, string> display = null)
        {
            var mService = new WebChannel<IMeetingMngService>();
            List<SelectListItem> typeList = new List<SelectListItem>();
            int count = 0;
            mService.Use(service => service.GetMeetingKindList().ForEach(d =>
            {
                if (d != null)
                {
                    //塞預設值
                    if (count == 0)
                    {
                        SelectListItem defaultitem = new SelectListItem();
                        defaultitem.Text = "-" + EP.IBResources.請選擇 + "-";
                        defaultitem.Value = "";
                        typeList.Add(defaultitem);
                    }
                    display = display ?? ((ValueText m) => m.Value + "-" + m.Text);
                    SelectListItem item = new SelectListItem();
                    item.Text = display(d);
                    item.Value = d.Value;
                    if (d.Value == initval)
                    {
                        item.Selected = true;
                    }
                    typeList.Add(item);
                    count += 1;
                }
            }));
            return typeList;
        }
    }
}