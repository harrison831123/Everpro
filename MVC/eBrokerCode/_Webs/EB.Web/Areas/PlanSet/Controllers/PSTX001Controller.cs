using EB.EBrokerModels;
using EB.VLifeModels;
using EB.SL.PlanSet.Service;
using Microsoft.CUF.Framework.Service;
using Microsoft.CUF.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.CUF;
using EB.SL.PlanSet.Models;

namespace EB.SL.PlanSet.Web.Areas.PlanSet.Controllers
{
    [Program("PSTX001")]
    public class PSTX001Controller : BaseController
    {
        private IPlanSetService _service;

        public PSTX001Controller()
        {
            _service = ServiceHelper.Create<IPlanSetService>();
        }
        // GET: PlanSet/PSTX001
        public ActionResult Index()
        {
            WebChannel<IPlanSetService> Service = new WebChannel<IPlanSetService>();
            agym agym = new agym();
            Service.Use(s => agym = s.GetYMData("agym", true).FirstOrDefault());

            OpCalendar calendar = new OpCalendar();
            calendar.ProductionYM = agym.ProductionYM;
            return View(calendar);
        }
        /// <summary>
        /// 查詢一筆OpCalendar資料
        /// </summary>
        /// <param name="productionYM">業績年月</param>
        /// <param name="sequence">序號</param>
        [HttpPost]
        [HasPermission("EB.SL.PlanSet.PSTX001")]
        public JsonResult QueryOpCalendar(string productionYM, string sequence)
        {
            OpCalendarViewModel result = new OpCalendarViewModel();
            try
            {
                OpCalendar model = new OpCalendar();
                model.ProductionYM = productionYM == string.Empty ? "" : productionYM;
                model.Sequence = sequence;

                result = _service.QueryOpCalendar(model);
                if (result != null)
                {
                    if (!string.IsNullOrEmpty(result.UpdateUserCode))
                    {
                        result.UpdateUserName = Member.Get(Account.Get(result.UpdateUserCode).MemberID).Name;
                    }
                    if (!string.IsNullOrEmpty(result.CreateUserCode))
                    {
                        result.CreateUserName = Member.Get(Account.Get(result.CreateUserCode).MemberID).Name;
                    }
					if (result.AdjDateTimeStr != null)
					{
						result.AdjDateTimeStrView = result.AdjDateTimeStr.ToString();
                        result.AdjDateTimeStrView = DateTime.Parse(result.AdjDateTimeStrView).ToString("yyyy/MM/dd HH:mm");
                    }
					if (result.AdjDateTimeEnd != null)
					{
						result.AdjDateTimeEndView = result.AdjDateTimeEnd.ToString();
                        result.AdjDateTimeEndView = DateTime.Parse(result.AdjDateTimeEndView).ToString("yyyy/MM/dd HH:mm");
                    }
					if (!string.IsNullOrEmpty(result.CreateDateTime))
                    {
                        result.CreateDateTime = DateTime.Parse(result.CreateDateTime).ToString("yyyy/MM/dd HH:mm:ss");
                    }
                    if (!string.IsNullOrEmpty(result.UpdateDateTime))
                    {
                        result.UpdateDateTime = DateTime.Parse(result.UpdateDateTime).ToString("yyyy/MM/dd HH:mm:ss");
                    }
                }
            }
            catch (Exception ex)
            {
                AppendMessage(ex.Message, false);
            }
            return Json(result);
        }

        /// <summary>
        /// 更新一筆OpCalendar資料
        /// </summary>
        /// <param name="productionYM">業績年月</param>
        /// <param name="hrCloseDate">人事關檔日期</param>
        /// <param name="salRunDate">跑佣日期</param>
        /// <param name="salPayDate">發佣日期</param>
        /// <param name="salReceiptDate">簽收回條截止日</param> 
        /// <param name="adjDateTimeStr">調整起日</param>
        /// <param name="adjDateTimeEnd">調整迄日</param>
        /// <param name="sequence">序號</param>
        /// <param name="remark">備註</param>
        /// <returns></returns>
        [HttpPost]
        [HasPermission("EB.SL.PlanSet.PSTX001")]
        public bool Update(OpCalendar model)
        {
            //         OpCalendar model = new OpCalendar();
            //         model.Iden = iden;
            //         model.ProductionYM = productionYM;
            model.HrCloseDate = model.HrCloseDate ?? "";
            //         model.SalRunDate = salRunDate;
            //         model.SalPayDate = salPayDate;
            //         model.SalReceiptDate = salReceiptDate;
            model.OpenQueryDateAnn = model.OpenQueryDateAnn ?? "";
            model.OpenQueryDate = model.OpenQueryDate ?? "";
            //         model.Sequence = sequence;
            //         //model.AdjDateTimeStr = !String.IsNullOrEmpty(adjDateTimeStr) ? DateTime.Parse(adjDateTimeStr + ":00.000").ToString("yyyy/MM/dd HH:mm:ss.fff") : null;
            //         //model.AdjDateTimeEnd = !String.IsNullOrEmpty(AdjDateTimeEnd) ? DateTime.Parse(adjDateTimeEnd + ":59.000").ToString("yyyy/MM/dd HH:mm:ss.fff") : null;
            //         model.AdjDateTimeStr = adjDateTimeStr;
            //model.AdjDateTimeEnd = adjDateTimeEnd;
            model.Remark = model.Remark ?? "";
            model.UpdateUserCode = User.AccountInfo.ID;

            bool result = _service.UpdateOpCalendar(model);

            if (result)
                AppendMessage("更新成功", false);
            else
                AppendMessage("更新失敗", false);

            return result;
        }

        /// <summary>
        /// 新增一筆OpCalendar資料
        /// </summary>
        /// <param name="productionYM">業績年月</param>
        /// <param name="hrCloseDate">人事關檔日期</param>
        /// <param name="salRunDate">跑佣日期</param>
        /// <param name="salPayDate">發佣日期</param>
        /// <param name="salReceiptDate">簽收回條截止日</param> 
        /// <param name="adjDateTimeStr">調整起日</param>
        /// <param name="adjDateTimeEnd">調整迄日</param>
        /// <param name="sequence">序號</param>
        /// <param name="remark">備註</param>
        [HttpPost]
        [HasPermission("EB.SL.PlanSet.PSTX001")]
        public void Insert(OpCalendar model)
        {
			//OpCalendar model = new OpCalendar();
			//model.ProductionYM = productionYM;
			model.HrCloseDate = model.HrCloseDate ?? "";
			//model.SalRunDate = salRunDate;
			//model.SalPayDate = salPayDate;
			//model.SalReceiptDate = salReceiptDate;
			model.OpenQueryDateAnn = model.OpenQueryDateAnn ?? "";
			model.OpenQueryDate = model.OpenQueryDate ?? "";
			//model.Sequence = sequence;
			////model.AdjDateTimeStr = !String.IsNullOrEmpty(adjDateTimeStr) ? DateTime.Parse(adjDateTimeStr + ":00.000").ToString("yyyy/MM/dd HH:mm:ss.fff") : null;
			////model.AdjDateTimeEnd = !String.IsNullOrEmpty(adjDateTimeEnd) ? DateTime.Parse(adjDateTimeEnd + ":59.000").ToString("yyyy/MM/dd HH:mm:ss.fff") : null;
			//model.AdjDateTimeStr = adjDateTimeStr;
			//model.AdjDateTimeEnd = adjDateTimeEnd;
			model.Remark = model.Remark ?? "";
            model.CreateUserCode = User.AccountInfo.ID;

            string result = _service.InsertOpCalendar(model);

            if (result == "OK")
            {
                AppendMessage("新增成功", false);
            }
            else if (result == "Duplicate")
            {
                //AppendMessage("新增失敗，業績年月重複", false);
                Throw.BusinessError("新增失敗，業績年月重複");
            }
            else
            {
                AppendMessage("新增失敗", false);
            }
        }

        /// <summary>
        /// 刪除一筆OpCalendar資料
        /// </summary>
        /// <param name="iden">自動識別碼</param>
        [HttpPost]
        [HasPermission("EB.SL.PlanSet.PSTX001")]
        public void Delete(string iden)
        {
            bool result = _service.DeleteOpCalendarByIden(iden, User.AccountInfo.ID);

            if (result)
                AppendMessage("刪除成功", false);
            else
                AppendMessage("刪除失敗", false);
        }

        /// <summary>
        /// 查詢OpCalendar Log軌跡
        /// </summary>
        /// <param name="productionYM">業績年月</param>
        /// <param name="sequence">序號</param>
        [HttpPost]
        [HasPermission("EB.SL.PlanSet.PSTX001")]
        public void QueryLog(string productionYM, string sequence)
        {
            OpCalendar model = new OpCalendar();
            model.ProductionYM = productionYM;
            model.Sequence = sequence;

            //取資料
            List<OpCalendarLog> list = new List<OpCalendarLog>();
            WebChannel<IPlanSetService> _channelService = new WebChannel<IPlanSetService>();
            _channelService.Use(service => list = service.QueryAdjDateTimeUpateLog(model));
            for (int i = 0; i < list.Count; i++)
            {
                if (!string.IsNullOrEmpty(list[i].UpdateUserCode))
                {
                    list[i].UpdateUserName = Member.Get(Account.Get(list[i].UpdateUserCode).MemberID).Name;
                }
                if (!string.IsNullOrEmpty(list[i].CreateUserCode))
                {
                    list[i].CreateUserName = Member.Get(Account.Get(list[i].CreateUserCode).MemberID).Name;
                }
                list[i].AdjDateTimeStr = Convert.ToDateTime(list[i].AdjDateTimeStr).ToString("yyyy/MM/dd tt hh:mm:ss");
                list[i].AdjDateTimeEnd = Convert.ToDateTime(list[i].AdjDateTimeEnd).ToString("yyyy/MM/dd tt hh:mm:ss");
            }

            var gridKey = _channelService.DataToCache(list.AsEnumerable());
            SetGridKey("QueryLogGrid", gridKey);
        }

        [HasPermission("EB.SL.PlanSet.PSTX001")]
        public JsonResult BindLogGrid(jqGridParam jqParams)
        {
            var cacheKey = GetGridKey("QueryLogGrid");
            return BaseGridBinding<OpCalendarLog>(jqParams,
                () => new WebChannel<IPlanSetService, OpCalendarLog>().Get(cacheKey));
        }

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="selyear">年度</param>
        /// <param name="selyear">月份</param>
        /// <returns></returns>
        [HttpPost]
        [HasPermission("EB.SL.PlanSet.PSTX001")]
        public ActionResult GetOpCalendarReport(OpCalendar op)
        {
            string fileName = "佣酬相關日期設定報表" + ".xlsx";
			if (!String.IsNullOrEmpty(op.ProductionYM))
			{
                //增加年份撈取條件
                op.ProductionYM = op.ProductionYM.Substring(0, 4);
            }            
            var List = _service.GetOpCalendar();

            if (List.Count != 0)
            {
                var ms = _service.GetOpCalendarReportList(op.ProductionYM);
                var filename = Url.Encode(fileName);
                return File(ms, "application/octet-estream", filename);
            }
            else
            {
                TempData["message"] = "查無資料!";
                return RedirectToAction("Index");
            }
        }
    }
}