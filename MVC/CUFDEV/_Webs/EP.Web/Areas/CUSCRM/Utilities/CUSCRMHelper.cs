using EP.SD.SalesSupport.CUSCRM.Service;
using Microsoft.CUF.Framework;
using Microsoft.CUF.Framework.Service;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EP.SD.SalesSupport.CUSCRM.Web
{
    /// <summary>
    /// 客服業務Helper
    /// </summary>
    public static class CUSCRMHelper
    {
        /// <summary>共用服務</summary>
        public static ICommonService _commonService = ServiceHelper.Create<ICommonService>();

        /// <summary>立案服務</summary>
        public static ICaseService _caseService = ServiceHelper.Create<ICaseService>();

        /// <summary>
        /// 取得類別的下拉清單
        /// </summary>
        /// <returns>類別的清單</returns>
        public static List<SelectListItem> GetSelectListItem<T>() where T : Enum
        {
            return Enum.GetValues(typeof(T)).Cast<T>().Select(se => {
                return new SelectListItem
                {
                    Text = se.GetName(),
                    Value = se.ToString()
                };
            }
            ).ToList();
        }

        /// <summary>
        /// 取得類別的下拉清單
        /// </summary>
        /// <returns>類別的清單</returns>
        public static List<SelectListItem> GetCategoryList()
        {
            return Enum.GetValues(typeof(Category)).Cast<Category>().Select(se => {
                return new SelectListItem
                {
                    Text = se.GetName(),
                    Value = se.GetValue().ToString()
                };
            }
            ).ToList();
        }

        /// <summary>
        /// 依類別取得類型清單
        /// </summary>
        /// <param name="category">類別</param>
        /// <returns>對應的類型清單</returns>
        public static List<SelectListItem> GetCaseTypeList(Category? category)
        {
            var datas = _commonService.GetCaseType(category);
            return datas.Select(m =>
            {
                return new SelectListItem()
                {
                    Value = m.Code.GetValue().ToString(),
                    Text = string.Format("{0}({1})", m.Code.GetValue(), m.Name)
                };
            }).ToList();
        }

        /// <summary>
        /// 依類型取得資料來源
        /// </summary>
        /// <param name="caseType">類型</param>
        /// <returns>資料來源清單</returns>
        public static List<SelectListItem> GetSourceList(DiscipTypeCode caseType)
        {
            var datas = _caseService.GetSourceListByType(caseType);
            return datas.Select(m =>
            {
                return new SelectListItem()
                {
                    Value = m.ID.ToString(),
                    Text =  m.Name
                };
            }).ToList();
        }

        /// <summary>
        /// 依類型取得來電者清單
        /// </summary>
        /// <param name="caseType">類型</param>
        /// <returns>來電者清單</returns>
        public static List<SelectListItem> GetCallerList(DiscipTypeCode caseType)
        {
            var datas = _caseService.GetCallerListByType(caseType);
            return datas.Select(m =>
            {
                return new SelectListItem()
                {
                    Value = m.ID.ToString(),
                    Text = m.Name
                };
            }).ToList();
        }

        /// <summary>
        /// 依類型取得案件類別清單
        /// </summary>
        /// <param name="caseType">類型</param>
        /// <returns>案件類別清單</returns>
        public static List<SelectListItem> GetContentCaseTypeList(DiscipTypeCode caseType)
        {
            var datas = _caseService.GetCaseTypeListByType(caseType);
            return datas.Select(m =>
            {
                return new SelectListItem()
                {
                    Value = m.ID.ToString(),
                    Text = m.Name
                };
            }).ToList();
        }

        /// <summary>
        /// 取得保險公司清單
        /// </summary>
        /// <returns>保險公司清單</returns>
        public static List<SelectListItem> GetCompanyList()
        {
            var datas = _caseService.GetInsCompanyList();
            return datas.Select(m =>
            {
                return new SelectListItem()
                {
                    Value = m.Value,
                    Text =  m.Text
                };
            }).ToList();
        }

        /// <summary>
        /// 依類型取得案件類型清單
        /// </summary>
        /// <param name="caseType">類型</param>
        /// <returns>案件類型清單</returns>
        public static List<SelectListItem> GetContentCategoryTypeList(DiscipTypeCode caseType)
        {
            var datas = _caseService.GetCaseCategoryListByType(caseType);
            return datas.Select(m =>
            {
                return new SelectListItem()
                {
                    Value = m.ID.ToString(),
                    Text = m.Name
                };
            }).ToList();
        }

        /// <summary>
        /// 取得案件狀態類別的下拉清單
        /// </summary>
        /// <returns>類別的清單</returns>
        public static List<SelectListItem> GetStatusCategoryList()
        {
            return Enum.GetValues(typeof(StatusCategory)).Cast<StatusCategory>().Select(se => {
                return new SelectListItem
                {
                    Text = se.GetName(),
                    Value = se.GetValue().ToString()
                };
            }
            ).ToList();
        }
    }
}