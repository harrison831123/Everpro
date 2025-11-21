using EP.SD.SalesSupport.CUSCRM.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using EP.Web;

namespace EP.SD.SalesSupport.CUSCRM.Web
{
    /// <summary>
    /// 立案
    /// </summary>
    [Program("CUSCRMTX001")]
    public class CUSCRMTX001Controller : BaseController
    {
        /// <summary>立案服務</summary>
        private ICaseService caseService;

        /// <summary>
        /// 建構子
        /// </summary>
        public CUSCRMTX001Controller()
        {
            caseService = ServiceHelper.Create<ICaseService>();
        }

        /// <summary>
        /// 立案頁面
        /// </summary>
        /// <returns>立案頁面</returns>
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 保單號碼查詢，並且回傳查詢的結果
        /// </summary>
        /// <param name="policyNo">保單號碼</param>
        /// <returns>result:保單號碼是否只有一筆且為總公司
        ///          datas:符合的保單號碼清單資料
        /// </returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult CheckPolicyNo(string policyNo)
        {
            var result = caseService.CheckPolicyNo(policyNo);

            return Json(new { result = result.Item1, datas = result.Item2 });

        }

        /// <summary>
        /// 依要保人id取得要保人底下的所有保單資料
        /// </summary>
        /// <param name="ownerID">要保人id</param>
        /// <returns>
        ///     ValidPolicys:有效件+休眠件的保單資料清單
        ///     InvalidPolicys:無效件+不列入的保單資料清單
        ///     ownerMobile:要保人電話號碼
        /// </returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult GetPolicyDatas(string ownerID)
        {
            var result = caseService.GetInsPolicyByOwnerID(ownerID);
            var validStatus = new string[] { "有效件", "休眠件" };
            var validDatas = result.Where(m => validStatus.Contains(m.StatusType)).ToList();
            var invalidDatas = result.Where(m => !validStatus.Contains(m.StatusType)).ToList();

            var mobile = result.Where(m => m.StatusType == "有效件" && !string.IsNullOrWhiteSpace(m.OwnerMobile)).OrderByDescending(m => m.IssDate).FirstOrDefault();

            if (mobile == null)
            {
                mobile = result.Where(m => m.StatusType != "有效件" && !string.IsNullOrWhiteSpace(m.OwnerMobile)).OrderByDescending(m => m.IssDate).FirstOrDefault();
            }

            string ownerMobile = mobile?.OwnerMobile;

            return Json(new { ValidPolicys = validDatas, InvalidPolicys = invalidDatas, ownerMobile = ownerMobile });

        }

        /// <summary>
        /// 立案的處理
        /// </summary>
        /// <param name="model">立案的資料</param>
        /// <param name="policies">立案的保單資料</param>
        /// <returns>受理編號</returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        [PdLogFilter("EIP客服-立案", PITraceType.Insert)]
        public ActionResult CreateCase(CRMECaseContent model, string policies)
        {
            model.Creator = User.MemberInfo.ID;
            model.CreateTime = DateTime.Now;
            var policys = Newtonsoft.Json.JsonConvert.DeserializeObject<List<CRMEInsurancePolicy>>(policies);
            var result = caseService.CreateCaseContent(model, policys);
            return Json(new { No = result });
        }

        /// <summary>
        /// 取得已立案未通知的清單資料
        /// </summary>
        /// <returns>已立案未通知的清單資料</returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult GetWaitNofity()
        {
            var result = caseService.GetWaitNofityDatas();
            return Json(result);
        }

        /// <summary>
        /// 檢核保單號碼有未結案的資料
        /// </summary>
        /// <param name="policyNo">保單號碼</param>
        /// <returns>true:有未結案 false:都結案或無立案資料</returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult CheckNotClosedPolicyNo(string policyNo)
        {
            var result = caseService.CheckNotClosedPolicyNo(policyNo);

            return Json(result);
        }

        /// <summary>
        /// 取得業務員資訊
        /// </summary>
        /// <param name="agentCode">業務員代碼</param>
        /// <returns>業務員資訊</returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult GetAgentInfo(string agentCode)
        {
            var result = caseService.GetAgentInfo(agentCode);

            return Json(new {result = result != null, data = result });

        }
    }
}