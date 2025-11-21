//20250326001-保險公司受理前十大商品年報排程 20250507 by Harrison
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Top10ReportProcess.Model;

namespace Top10ReportProcess
{
    public class FileHelper
    {
        /// <summary>
        /// 另存檔案路徑
        /// </summary>
        /// <param name="originalFilePath"></param>
        /// <param name="newFileName"></param>
        /// <returns></returns>
        public string SaveExcelAsNewFile(string originalFilePath, string newFileName)
        {
            if (!File.Exists(originalFilePath))
                throw new FileNotFoundException("找不到原始 Excel 檔案。", originalFilePath);

            string directory = Path.GetDirectoryName(originalFilePath);
            string newFilePath = Path.Combine(directory, newFileName);

            // 開啟並另存 Excel
            using (var workbook = new XLWorkbook(originalFilePath))
            {
                workbook.SaveAs(newFilePath);
            }

            return newFilePath;
        }

        /// <summary>
        /// 保險公司受理前十大產品年報
        /// </summary>
        /// <param name="templatePath"></param>
        /// <param name="data"></param>
        /// <param name="startRow">預設從第2列開始產生資料</param>
        /// <returns>回傳報表資料流</returns>
        public MemoryStream AnnualReportByExcelTemplate(string templatePath, List<List<TopReprotModel>> data, string reportName, int startRow = 2)
        {
            if (!File.Exists(templatePath))
                throw new FileNotFoundException("找不到範本檔案", templatePath);

            using (var workbook = new XLWorkbook(templatePath))
            {
                var stream = new MemoryStream();
                string[] type = new[] { "以FYC排名", "以件數排名", "以FYP排名" };

                for (int sheetIndex = 0; sheetIndex < data.Count; sheetIndex++)
                {
                    var itemList = data[sheetIndex];
                    var sheet = workbook.Worksheet(sheetIndex + 1); // ClosedXML sheetIndex 從 1 開始                                                                 
                    sheet.Range("A1:H1").Merge(); // 合併 A1 到 H1                   
                    sheet.Cell("A1").Value = "保險公司受理前十大產品年報_" + reportName + "(" + type[sheetIndex] + ")"; // 設定內容
                    // 可選：置中對齊（水平+垂直）
                    sheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;                
                    sheet.Cell("A1").Style.Font.Bold = true; // 可選：加粗字體

                    for (int i = 0; i < itemList.Count; i++)
                    {
                        var item = itemList[i];
                        int row = startRow + i;

                        sheet.Cell(row, 1).Value = item.Rank;
                        sheet.Cell(row, 2).Value = item.Company;
                        sheet.Cell(row, 3).Value = item.PlanTitle;
                        sheet.Cell(row, 4).Value = item.PlanCode;
                        sheet.Cell(row, 5).Value = item.FYC;
                        sheet.Cell(row, 6).Value = item.Cnt;
                        sheet.Cell(row, 7).Value = item.FYP;
                        sheet.Cell(row, 8).Value = item.Proportion;
                    }
                }
                workbook.SaveAs(stream);
                stream.Position = 0;
                return stream;
            }
        }
    }
}
