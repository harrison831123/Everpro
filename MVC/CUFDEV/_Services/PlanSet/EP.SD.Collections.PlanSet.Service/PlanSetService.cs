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