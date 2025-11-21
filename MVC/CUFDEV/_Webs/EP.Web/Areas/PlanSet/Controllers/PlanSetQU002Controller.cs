using EP.Platform.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using Microsoft.CUF;
using Microsoft.CUF.Framework;
using Microsoft.CUF.Framework.Service;
using EP.SD.Collections.PlanSet.Service;
using EP.VBEPModels;

/// <summary>
/// 佣酬預估試算
/// 202505 by Fion 20250527001_佣酬預估試算
/// 
/// </summary>
namespace EP.SD.Collections.PlanSet.Web.Controllers
{
    [Program("PLANSETQU002")]
    public class PlanSetQU002Controller : BaseController
    {
        private IPlanSetService _service;
        private WebChannel<IPlanSetService> channel;
        
        public PlanSetQU002Controller()
        {
            _service = ServiceHelper.Create<IPlanSetService>();
            channel = new WebChannel<IPlanSetService>();
        }

        /// <summary>
        /// 進入頁面
        /// </summary>
        /// <returns></returns>
        [HasPermission("EP.SD.Collections.PlanSet.PlanSetQU002")]
        public ActionResult Index()
        {
            AgentRewardPolicyCondition condition = new AgentRewardPolicyCondition();
            List<AgentRewardRange> rewardRanges = null ;
            condition.AgentId = User.AccountInfo.ID;
            condition.ViewMsg = "";
            condition.AgentTotalIncome = 0;

            var agentData = _service.GetAgentData(condition);
            if (agentData!= null)
            {
                condition.AgentCode = agentData.GetOrDefault("AgentCode");
                condition.AgLevel = agentData.GetOrDefault("AgLevel");
                condition.AgentName = agentData.GetOrDefault("AgentName");
                condition.AgLevelName = agentData.GetOrDefault("AgLevelName");
                condition.MDRT = _service.GetAgentRewardMDRT(condition);
                rewardRanges = _service.GetAgentRewardRange(condition).ToList();

            }
            ViewBag.CalYMList = GetCalYMListItem();
            //202510 by Fion 20250901003_佣酬預估試算優化
            condition.CalYmS = DateTime.Now.ToString("yyyy/MM");
            condition.CalYmE = DateTime.Now.ToString("yyyy/MM");
            //預估合計
            var incomeTol = _service.QueryAgentRewardTotIncome(condition.AgentCode);
            if (incomeTol != null)
            {
                try
                {
                    condition.AgentTotalIncome = int.Parse(incomeTol.GetOrDefault("income_tol"));
                }
                catch (Exception)
                {
                    condition.AgentTotalIncome = 0;
                }
            }

            Tuple<AgentRewardPolicyCondition, List<AgentRewardRange>> model = null;

            model = Tuple.Create(condition, rewardRanges);
            
            return View(model);
        }

        /// <summary>
        /// 查詢資料
        /// </summary>
        /// <param name="conditions"></param>
        /// <returns></returns>
        [HasPermission("EP.SD.Collections.PlanSet.PlanSetQU002")]
        [HttpPost]
        [EP.Web.PdLogFilter("佣酬試算", EP.Web.PITraceType.Insert)]
        public JsonResult QueryUnPaidRewardPolicyData(AgentRewardPolicyCondition conditions)
        {
            //202510 by Fion 20250901003_佣酬預估試算優化
            conditions.CalYmS = reSetCYm(conditions.CalYmS);
            conditions.CalYmE = reSetCYm(conditions.CalYmE);
            conditions.AgentId = User.AccountInfo.ID;

            //排除已核實的保單
            var result = _service.QueryUnPaidRewardPolicyData(conditions)
                    .Where(m=>m.ExtraDataInfo.GetOrDefault("IsPaid") =="").ToList();
            if(result.Count() == 0)
            {
                return Json(null);
            }
            return Json(result);
        }

        /// <summary>
        /// 受理年月下拉選單處理
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<SelectListItem> GetCalYMListItem()
        {
            var result = new List<SelectListItem>();
            for (int i = 0; i < 6; i++)
            {
                result.Add(new SelectListItem { Value = DateTime.Now.AddMonths(-i).ToString("yyyy/MM"), Text = DateTime.Now.AddMonths(-i).ToString("yyyy/MM") });

            }
            return result;
        }

        /// <summary>
        /// 202510 by Fion 20250901003_佣酬預估試算優化
        /// </summary>
        /// <returns></returns>
        private string reSetCYm(string yM)
        {
            var calYM = DateTime.Parse(yM + "/01");
            return calYM.GetString("CY/MM").PadLeft(7, '0');
        }
    }
}