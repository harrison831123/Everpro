using EP.Platform.Service;
using EP.SD.Collections.PlanSet.Models;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using EP.VBEPModels;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Reflection;
using System.Drawing;

namespace EP.SD.Collections.PlanSet.Service
{
    /// <summary>
    /// 需求單號：20241210003 現售商品佣獎查詢 2024.12 BY VITA
    /// 需求單號：20250707001 競賽計C商品清單。 2025/07/03 BY Harrison 
    /// 202505 by Fion 20250527001_佣酬預估試算
    /// </summary>
    public class PlanSetService : IPlanSetService
    {
        #region 查詢頁面提供資訊
        /// <summary>
        /// 取得保險公司資料
        /// </summary>
        /// <returns></returns>
        public List<ValueText> GetCompanyCode()
        {
            //需求單號：20241210003 國泰(原國寶)人壽/南山(原朝陽)人壽/國華人壽，因已無合約，故請不要顯示，並增加排序。 BY VITA 2025.01
            string sql = @"
            SELECT term_code as Value, term_meaning as Text
            FROM trmval 
            WHERE term_id = 'company_code' and ISNUMERIC(term_code) = 1
                AND term_code NOT IN('007', '010', '012')
            ORDER BY term_code";
            List<ValueText> result = DbHelper.Query<ValueText>(VLifeRepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 需求單號：202503XX00X 增加險種中文、險種代碼模糊查詢 by vita 2025.03
        /// </summary>
        /// <param name="company_name">保險公司代碼</param>
        /// <param name="plan_title">險種中文名稱 / 險種代碼</param>
        /// <param name="chktype">0:險種中文名稱、1:險種代碼</param>
        /// <returns></returns>
        public List<PlanTitle> GetPlanTitle(string company_name, string plan_title, string chktype)
        {
            List<PlanTitle> result = new List<PlanTitle>();
            string sql = @"EXEC usp_GetPlanSetName @company_name, @plan_title, @chktype";
            var Result = DbHelper.QueryMultiple(VLifeRepository.ConnectionStringName, sql, new
            {
                company_name = company_name.Trim(),
                plan_title = plan_title.Trim(),
                chktype = chktype.Trim()
            },
            resultTypes: new Type[] { typeof(PlanTitle) });
            result = DbHelper.Query<PlanTitle>(VLifeRepository.ConnectionStringName, sql,
                   new
                   {
                       company_name = company_name,
                       plan_title = plan_title,
                       chktype = chktype
                   }).ToList();
            //List<string> dtResult = new List<string>();
            //for (int i = 0; i < result.Count; i++)
            //{
            //    dtResult[i] = result[i].plan_code.Trim() + result[i].plan_title.Trim();
            //}
            return result;
        }

        #endregion

        #region 查詢        
        /// <summary>
        /// 抓取初年度業績換算率(業行部)、保公獎勵內容(業行部)、永達競賽獎勵(業支部)
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public PlanSetAll GetQueryArea(PlanSetMainInput model)
        {
            string sql = "";
            sql = @"EXEC usp_GetAllPlanSetData @company_code, @plan_code, @collect";
            var Result = DbHelper.QueryMultiple(VLifeRepository.ConnectionStringName, sql, new
            {
                company_code = model.company_name.Trim(),
                plan_code = model.plan_code.Trim(),
                collect = model.collect
            },
            resultTypes: new Type[] { typeof(PlanSetArea1), typeof(PlanSetArea2), typeof(PlanSetArea3) });
            PlanSetAll dtResult = new PlanSetAll()
            {
                PlanSetArea1 = (List<PlanSetArea1>)(IEnumerable<PlanSetArea1>)Result[0],
                PlanSetArea2 = (List<PlanSetArea2>)(IEnumerable<PlanSetArea2>)Result[1],
                PlanSetArea3 = (List<PlanSetArea3>)(IEnumerable<PlanSetArea3>)Result[2]
            };

            return dtResult;
        }
        #endregion

        #region 佣酬預估試算 202505 by Fion 20250527001_佣酬預估試算
        /// <summary>
        /// 取得業務員基本資料
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public ExtraData GetAgentData(AgentRewardPolicyCondition condition)
        {
            //在職
            var sql = @"
                        select agent_code AgentCode,[name] agent_name,[name] AgentName ,a.ag_level AgLevel
                        ,b.level_name_bf as AgLevelNameBf,b.level_name_chs as AgLevelName
                        from V_orgin a WITH(NOLOCK)
                        JOIN v_aglevel_occpind b WITH(NOLOCK) ON a.ag_level = b.ag_level and a.ag_occp_ind = b.ag_occp_ind
                        WHERE 1=1 
                        AND a.ag_status_code in ('0','1') 
                    ";

            if (!string.IsNullOrWhiteSpace(condition.AgentCode))
            {
                sql += " AND a.agent_code = @AgentCode";
            }

            if (!string.IsNullOrWhiteSpace(condition.AgentId))
            {
                sql += " AND LEFT(a.agent_code,10) = @AgentId";
            }

            var result = DbHelper.Query<dynamic>(VLifeRepository.ConnectionStringName
                    ,sql
                    , condition
                ).Select(m =>
                {
                    var extraData = new ExtraData();
                    ((IDictionary<string, object>)m).ForEach(item =>
                    {
                        extraData.Add(item.Key, item.Value != null ? Convert.ToString(item.Value) : "");
                    });
                    return extraData;
                }).FirstOrDefault();

            return result;
        }

        /// <summary>
        /// 業務員最新MDRT屆數
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public string GetAgentRewardMDRT(AgentRewardPolicyCondition condition)
        {
            var result = DbHelper.Query<string>(VBEPRepository.ConnectionStringName,
                    @" select convert(varchar,PeriodNum) MDRT
                        from AgentRewardMDRT a WITH(NOLOCK)
                        WHERE 1=1 
                        AND RType = '1'
                        AND client_id = LEFT(@AgentCode,10)
                        ORDER BY RewardYear Desc
                       "
                    , condition
                ).FirstOrDefault();

            if (string.IsNullOrEmpty(result))
            {
                return "-";
            }
            return result;
        }

        /// <summary>
        /// 個人達成報酬率
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public IEnumerable<AgentRewardRange> GetAgentRewardRange(AgentRewardPolicyCondition condition)
        {
            var result = DbHelper.Query<AgentRewardRange>(VBEPRepository.ConnectionStringName,
                    @" select *
                        from AgentRewardRange a WITH(NOLOCK)
                        WHERE 1=1 
                        AND ag_level = @AgLevel
                        AND agtb_key1 >= 0
                       "
                    , condition
                ).ToList();

            return result;
        }

        /// <summary>
        /// 未核保保單資訊
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public IEnumerable<AgentRewardPolicyInfo> QueryUnPaidRewardPolicyData(AgentRewardPolicyCondition condition)
        {
            IEnumerable<AgentRewardPolicyInfo> result = null;

            var sql = @"
                    select rp.*,'' as splitOn ,RTRIM(t0.term_meaning) CompanyName,RTRIM(t1.term_meaning) PolicyStatusDesc
                    , IsNull(Cr.plan_title,'') as INSName
                    , CASE WHEN ISNULL(p.comm_prem_act, 0) <> 0 THEN 'Y' ELSE '' END AS IsPaid
                    , CASE WHEN ISNULL(pr.agent2_id,'') <> '' THEN 'Y' ELSE '' END AS IsHalf
                    FROM AgentRewardPolicy rp with(nolock)
                    LEFT JOIN trmval t0 with(nolock) on t0.term_id='company_code_sm' and t0.term_code = rp.company_code
					LEFT JOIN trmval t1 with(nolock) on t1.term_id='policy_status_code' and t1.term_code = rp.policy_status_code            
					LEFT JOIN policyrelate pr with(nolock) on rp.policy_serial = pr.policy_serial
                    LEFT JOIN policy P with(nolock) on rp.policy_serial = p.policy_serial
                    LEFT OUTER JOIN crat Cr with(nolock) on (p.prouct_type = cr.prouct_type) and (p.prouct_type1 = cr.prouct_type1) and (p.plan_code = cr.plan_code) 
                        and (p.proj_id =cr.proj_id) and (p.company_code = cr.company_code) 
                        and (p.insured_age between cr.plan_start_age and cr.plan_end_age) 
                        and (p.po_issue_date between cr.po_start_date and cr.po_start_end ) 
                        and (p.collect_year between cr.plan_start_date and cr.plan_start_end)     
                    WHERE 1=1
            ";

            if (!string.IsNullOrWhiteSpace(condition.AgentCode))
            {
                sql += " AND rp.agent_code = @AgentCode";
            }

            if (!string.IsNullOrWhiteSpace(condition.CalYM))
            {
                sql += " AND LEFT(rp.begin_date,7) = @CalYM";
            }
            //202510 by Fion 20250901003_佣酬預估試算優化
            if (!string.IsNullOrWhiteSpace(condition.CalYmS))
            {
                sql += " AND LEFT(rp.begin_date,7) >= @CalYmS";
            }

            if (!string.IsNullOrWhiteSpace(condition.CalYmE))
            {
                sql += " AND LEFT(rp.begin_date,7) <= @CalYmE";
            }

            result = DbHelper.Query<AgentRewardPolicy, dynamic, AgentRewardPolicyInfo>(VBEPRepository.ConnectionStringName,
                sql,
                (rp, dyna) =>
                {
                    var extraData = new ExtraData();
                    ((IDictionary<string, object>)dyna).ForEach(item =>
                    {
                        extraData.Add(item.Key, item.Value != null ? Convert.ToString(item.Value) : "");
                    });
                    return new AgentRewardPolicyInfo() { AgentRewardPolicy = rp, ExtraDataInfo = extraData };
                },
                "splitOn",
                condition);

            return result;
        }

        /// <summary>
        /// 取得業務員年度累計所得資料
        /// 202510 by Fion 20250901003_佣酬預估試算優化
        /// </summary>
        /// <param name="agentCode"></param>
        /// <returns></returns>
        public ExtraData QueryAgentRewardTotIncome(string agentCode)
        {
            var sql = @"
                        select *
                        from AgentRewardTotIncome
                        where agent_code = @AgentCode";

            var result = DbHelper.Query<dynamic>(VBEPRepository.ConnectionStringName
                    , sql
                    , new { AgentCode = agentCode }).Select(m => {
                        var extraData = new ExtraData();
                        ((IDictionary<string, object>)m).ForEach(item =>
                        {
                            extraData.Add(item.Key, item.Value != null ? Convert.ToString(item.Value) : "");
                        });
                        return extraData;
                    }).FirstOrDefault();

            return result;
        }        
        #endregion 佣酬預估試算

        #region 競賽計C商品清單
        /// <summary>
        /// 取得報表工作日
        /// </summary>
        /// <returns></returns>
        public string GetWorkDate()
        {
            string sql = @" SELECT TOP 1 FORMAT(DATEADD(day, -1, create_datetime), 'yyyy/MM/dd') AS create_date_previous_day
                            FROM PlanSet_WarptSet";
    
            var result = DbHelper.Query<string>(VLifeRepository.ConnectionStringName, sql).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 取得商品清單資料
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>

        public List<PlanSetWarptSet> GetPlanSetWarptSet(PlanSetWarptSetCondition condition)
        {
            string sql = @" SELECT * FROM PlanSet_WarptSet psws WHERE 1 = 1";

            switch (condition.SelectedProduct)
            {
                case ProductList.Y:
                    sql += " AND IsCredited = 'Y' ";
                    break;
                case ProductList.N:
                    sql += " AND IsCredited = 'N' ";
                    break;
                case ProductList.Maintenance:
                    sql += " AND ISNULL(IsCredited, '') = '' ";
                    break;
                default:
                    //顯示全部
                    break;
            }
            if (condition.CompanyCode != "0")
            {
                sql += " AND psws.company_code = @CompanyCode";
            }
            //if (!string.IsNullOrWhiteSpace(condition.PlanYear))
            //{
            //    sql += " AND psws.plan_year_str <= @PlanYear AND psws.plan_year_end >= @PlanYear";
            //}

            sql += " ORDER BY psws.company_code,psws.plan_code,psws.plan_year_str,psws.set_start_date";

            var result = DbHelper.Query<PlanSetWarptSet>(VLifeRepository.ConnectionStringName, sql, new { CompanyCode = condition.CompanyCode }).ToList();
            return result;
        }

        /// <summary>
        /// 取得報表
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public Stream GetPlanSetWarptSetReport(PlanSetWarptSetCondition condition)
        {
            //取出資料
            var Report = GetPlanSetWarptSet(condition);
            MemoryStream ms = new MemoryStream();
            ExcelPackage excel = new ExcelPackage();
            ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("競賽計C商品清單");
            //字型
            sheet.Cells.Style.Font.Name = "標楷體";
            //文字大小
            sheet.Cells.Style.Font.Size = 12;

            #region 標題
            int row = 1;

            // 第一列右上角：資料日
            sheet.Cells[row, 8].Value = $"資料日：{GetWorkDate()}";
            sheet.Cells[row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            // 第二列右上角：查詢的前一工作日
            //row++;
            //sheet.Cells[row, 8].Value = "查詢的前一工作日";
            //sheet.Cells[row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ////sheet.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //sheet.Cells[row, 8].Style.Font.Color.SetColor(Color.Red);

            // 標題
            row++;
            sheet.Cells[row, 1, row, 8].Value = "競賽計C商品清單";
            sheet.Cells[row, 1, row, 8].Merge = true;
            sheet.Cells[row, 1, row, 8].Style.Font.Size = 16;
            sheet.Cells[row, 1, row, 8].Style.Font.Bold = true;
            sheet.Cells[row, 1, row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // 紅字標題
            row++;
            sheet.Cells[row, 2].Value = "以下四點請呈現紅字";
            sheet.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[row, 2].Style.Font.Color.SetColor(Color.Red);
            sheet.Cells[row, 2].Style.Font.Bold = true;
            sheet.Cells[row, 1, row, 8].Merge = true;
            sheet.Cells[row, 1, row, 8].Style.WrapText = true; // 准許換行
            sheet.Row(row).Height = 50;

            // 說明
            string[] notes = new string[]
            {
                "1.本清單僅呈現「現售」商品，且為查詢時前一工作日建檔資料。\n" +
                "2.新商品資料維護需作業時間，會有時間差，故暫呈現「維護中」。\n" +
                "3.『投資型』的FYC皆不計入競賽業績。\n" 
            };
          
            foreach (var note in notes)
            {
                sheet.Cells[row, 1, row, 8].Value = note;
                //sheet.Cells[row, 1, row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[row, 1, row, 8].Style.Font.Color.SetColor(Color.Red);
                //sheet.Cells[row, 1, row, 8].Style.WrapText = true;
            }

            // 定義要匯出的欄位順序和對應的屬性名稱
            string[] propertiesToExport = new string[]
            {
            //"ID",
            "CompanyName",
            "PlanTitle",
            "PlanCode",
            "PlanYearCond",
            "SetStartDate",
            "SetEndDate",
            "IsCreditedTxt",
            };

            // 標題列
            int headerRow = 4;
            string[] Title = { "序號", "保險公司", "險種名稱", "險種代號", "年期", "要保申請起日", "要保申請迄日", "競賽FYC" };
            for (int i = 0; i < Title.Length; i++)
            {
                ExcelSetCell(sheet, new string[] { Title[i] }, headerRow, i + 1);
                sheet.Cells[headerRow, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            #endregion

            int startDataRow = 5; // 資料起始列
            int startDataCol = 1; // 資料起始欄

            // 設定標題樣式 
            //sheet.Cells[headerRow, startDataCol, headerRow, startDataCol + propertiesToExport.Length - 1].Style.Font.Bold = true;
            //sheet.Cells[headerRow, startDataCol, headerRow, startDataCol + propertiesToExport.Length - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // 迴圈 Report 中的每一行資料
            for (int i = 0; i < Report.Count; i++)
            {
                int currentRow = startDataRow + i; // 計算當前 Excel 列
                var currentItem = Report[i];
                
                ExcelSetCell(sheet, new string[] { (i + 1).ToString() }, currentRow, startDataCol); // 第1列的欄位是序號(流水號)
                sheet.Cells[currentRow, startDataCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // 迴圈遍歷欄位
                for (int j = 0; j < propertiesToExport.Length; j++)
                {
                    int currentCol = startDataCol + j + 1 ; // 計算當前 Excel 欄
                    string propertyName = propertiesToExport[j];

                    // 取得屬性值
                    PropertyInfo prop = typeof(PlanSetWarptSet).GetProperty(propertyName);
                    if (prop != null)
                    {
                        object value = prop.GetValue(currentItem, null);
                        string valueStr = value?.ToString() ?? "";

                        ExcelSetCell(sheet, new string[] { valueStr }, currentRow, currentCol);
                        sheet.Cells[currentRow, currentCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                }
            }

            // 自動調整欄寬
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();         
            excel.SaveAs(ms);
            excel.Dispose();
            ms.Position = 0;
            return ms;
        }
        #endregion

        #region 資料填入Excel的Cell ExcelSetCell
        /// <summary>
        /// 資料填入Excel的Cell
        /// </summary>
        /// <param name="workSheet">sheet object</param>
        /// <param name="valueList">Data Array</param>
        /// <param name="rowStartPosition">Row Start Position</param>
        /// <param name="columnStartPosition">Column Start Position</param>
        private static void ExcelSetCell(ExcelWorksheet workSheet, string[] valueList, int rowStartPosition, int columnStartPosition)
        {
            var orgColPos = columnStartPosition;
            foreach (var value in valueList)
            {
                workSheet.Cells[rowStartPosition, columnStartPosition++].Value = value;

            }
            columnStartPosition = (columnStartPosition != orgColPos ? columnStartPosition - 1 : columnStartPosition);
            //下框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //上框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //右框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //左框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        }

        /// <summary>
        /// 資料填入Excel的Cell
        /// </summary>
        /// <param name="workSheet">sheet object</param>
        /// <param name="valueList">Data Array</param>
        /// <param name="rowStartPosition">Row Start Position</param>
        /// <param name="columnStartPosition">Column Start Position</param>
        private static void ExcelSetCell(ExcelWorksheet workSheet, int[] valueList, int rowStartPosition, int columnStartPosition)
        {
            var orgColPos = columnStartPosition;
            foreach (var value in valueList)
            {
                workSheet.Cells[rowStartPosition, columnStartPosition++].Value = value;

            }
            columnStartPosition = (columnStartPosition != orgColPos ? columnStartPosition - 1 : columnStartPosition);
            //下框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //上框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //右框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //左框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        }

        /// <summary>
        /// 資料填入Excel的Cell
        /// </summary>
        /// <param name="workSheet">sheet object</param>
        /// <param name="valueList">Data Array</param>
        /// <param name="rowStartPosition">Row Start Position</param>
        /// <param name="columnStartPosition">Column Start Position</param>
        private static void ExcelSetCell(ExcelWorksheet workSheet, double[] valueList, int rowStartPosition, int columnStartPosition)
        {
            var orgColPos = columnStartPosition;
            foreach (var value in valueList)
            {
                workSheet.Cells[rowStartPosition, columnStartPosition++].Value = value;

            }
            columnStartPosition = (columnStartPosition != orgColPos ? columnStartPosition - 1 : columnStartPosition);
            //下框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //上框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //右框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //左框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        }

        /// <summary>
        /// 資料填入Excel的Cell
        /// </summary>
        /// <param name="workSheet">sheet object</param>
        /// <param name="valueList">Data Array</param>
        /// <param name="rowStartPosition">Row Start Position</param>
        /// <param name="columnStartPosition">Column Start Position</param>
        private static void ExcelSetCell(ExcelWorksheet workSheet, float[] valueList, int rowStartPosition, int columnStartPosition)
        {
            var orgColPos = columnStartPosition;
            foreach (var value in valueList)
            {
                workSheet.Cells[rowStartPosition, columnStartPosition++].Value = value;

            }
            columnStartPosition = (columnStartPosition != orgColPos ? columnStartPosition - 1 : columnStartPosition);
            //下框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //上框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //右框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //左框線
            workSheet.Cells[rowStartPosition, orgColPos, rowStartPosition, columnStartPosition].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        }
        #endregion
    }
}