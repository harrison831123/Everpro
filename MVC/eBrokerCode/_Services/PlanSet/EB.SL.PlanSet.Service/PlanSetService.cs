using EB.EBrokerModels;
using EB.Platform.Service;
using EB.SL.PlanSet.Models;
using EB.VLifeModels;
using Microsoft.CUF.Framework.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EB.SL.PlanSet.Service
{
	public class PlanSetService: IPlanSetService
	{
		#region 調整起迄時間設定
		/// <summary>
		/// 取得工作月
		/// </summary>
		/// <param name="strSel">query type</param>
		/// <param name="exclude88">是否排掉 sequence=88</param>
		/// <param name="selectTop">要取得的筆數</param>
		/// <returns></returns>
		public List<agym> GetYMData(string strSel, bool exclude88, int selectTop = 1)
		{
			string sql = @"select top " + selectTop + " production_ym,sequence,agbc_ind from agym where 1=1";

			if (strSel == "agym")
			{
				sql += " and agym_ind='1'";
			}

			if (strSel == "agbc")
			{
				sql += " and agym_ind='1'";
			}

			//排掉88
			if (exclude88)
			{
				sql += " and [sequence] <> 88";
			}

			sql += " order by production_ym desc,sequence desc";

			List<agym> result = DbHelper.Query<agym>(
				  VLifeRepository.ConnectionStringName, sql).ToList();
			if (result != null)
			{
				for (int i = 0; i < result.Count; i++)
				{
					string[] sArray = result[i].ProductionYM.Split('/');
					result[i].ProductionYM = (Convert.ToInt32(sArray[0]) + 1911) + "/" + sArray[1];
				}
			}

			return result;
		}


		/// <summary>
		/// 查詢一筆OpCalendar時間
		/// </summary>
		public OpCalendarViewModel QueryOpCalendar(OpCalendar model)
		{
			string sqlGetData = @"select *
                                  from OpCalendar 
                                  where production_ym=@productionYM and sequence=@sequence ";
			OpCalendarViewModel result = DbHelper.Query<OpCalendarViewModel>(EBrokerRepository.ConnectionStringName, sqlGetData, new
			{
				productionYM = model.ProductionYM,
				sequence = model.Sequence
			}).FirstOrDefault();

			return result;
		}

		/// <summary>
		/// 依照ID查詢OpCalendar
		/// </summary>
		public OpCalendar GetOpCalendarByID(string iden)
		{
			string sqlGetData = @"select *
                                  from OpCalendar 
                                  where iden=@iden ";
			OpCalendar result = DbHelper.Query<OpCalendar>(EBrokerRepository.ConnectionStringName, sqlGetData, new
			{
				iden = iden,
			}).FirstOrDefault();

			return result;
		}

		/// <summary>
		/// 取得當次RUN佣日期
		/// </summary>
		/// <returns></returns>
		public OpCalendar GetsalrundateNow()
		{
			agym ag = new agym();
			OpCalendar calendar = new OpCalendar();
			string sql = "", sql1 = "";
			//取的當次RUN佣日期
			sql = @"select convert(varchar, convert(int, substring(production_ym, 1, charindex('/', production_ym) - 1))+1911)+substring(production_ym, charindex('/', production_ym), 10) as production_ym,sequence from agym where agbc_ind = 0 order by process_ym desc";
			ag = DbHelper.Query<agym>(VLifeRepository.ConnectionStringName, sql).FirstOrDefault();

			sql1 = @"SELECT * FROM OpCalendar where production_ym = @production_ym and sequence = @sequence";
			calendar = DbHelper.Query<OpCalendar>(EBrokerRepository.ConnectionStringName, sql1, new { production_ym = ag.ProductionYM, sequence = ag.sequence }).FirstOrDefault();

			return calendar;
		}

		/// <summary>
		/// 更新一筆OpCalendar紀錄
		/// </summary>
		/// <returns>identity</returns>
		public bool UpdateOpCalendar(OpCalendar model)
		{
			string sql = @"update OpCalendar 
                           set sequence=@sequence, 
                               production_ym=@productionYM,
                               hr_close_date=@hrCloseDate, 
                               sal_run_date=@salRunDate, 
                               sal_pay_date=@salPayDate, 
                               sal_receipt_date=@salReceiptDate,
                               adj_datetime_str=@adjDateTimeStr, 
                               adj_datetime_end=@adjDateTimeEnd, 
                               open_query_date=@OpenQueryDate, 
                               open_query_date_ann=@OpenQueryDateAnn, 
                               remark=@remark, 
                               update_datetime=getdate(),
                               update_user_code=@updateUserCode
                           where iden=@iden";
			int result = -1;
			try
			{
				result = DbHelper.Execute(EBrokerRepository.ConnectionStringName, sql, new
				{
					hrCloseDate = model.HrCloseDate,
					salRunDate = model.SalRunDate,
					salPayDate = model.SalPayDate,
					salReceiptDate = model.SalReceiptDate,
					adjDateTimeStr = model.AdjDateTimeStr,
					adjDateTimeEnd = model.AdjDateTimeEnd,
					remark = model.Remark,
					updateUserCode = model.UpdateUserCode,
					productionYM = model.ProductionYM,
					sequence = model.Sequence,
					iden = model.Iden,
					OpenQueryDateAnn = model.OpenQueryDateAnn,
					OpenQueryDate = model.OpenQueryDate
				});
				return (result > 0);
			}
			catch (Exception x)
			{
				return (false);
			}
		}

		/// <summary>
		/// 新增一筆OpCalendar紀錄
		/// </summary>
		public string InsertOpCalendar(OpCalendar model)
		{
			//1.檢查有無重複的業績年月&序號資料
			int duplicateCount = DbHelper.Query<int>(EBrokerRepository.ConnectionStringName,
				"select iden from OpCalendar where production_ym=@productionYM and [sequence]=@sequence", new
				{
					productionYM = model.ProductionYM,
					sequence = model.Sequence
				}).Count();
			if (duplicateCount > 0)
			{
				return "Duplicate";
			}
			else
			{
				//2.新增
				string sql = @"INSERT INTO OpCalendar (production_ym, [sequence], hr_close_date, sal_run_date, sal_pay_date,sal_receipt_date, adj_datetime_str, adj_datetime_end, remark, create_datetime, create_user_code,open_query_date,open_query_date_ann)
                           VALUES (@productionYM, @sequence, @hrCloseDate, @salRunDate, @salPayDate,@salReceiptDate, @adjDatetimeStr, @adjDatetimeEnd, @remark, getdate(), @createUserCode,@OpenQueryDate,@OpenQueryDateAnn);";
				int result = DbHelper.Execute(EBrokerRepository.ConnectionStringName, sql, new
				{
					productionYM = model.ProductionYM,
					sequence = model.Sequence,
					hrCloseDate = model.HrCloseDate,
					salRunDate = model.SalRunDate,
					salPayDate = model.SalPayDate,
					salReceiptDate = model.SalReceiptDate,
					adjDateTimeStr = model.AdjDateTimeStr,
					adjDateTimeEnd = model.AdjDateTimeEnd,
					remark = model.Remark,
					createUserCode = model.CreateUserCode,
					OpenQueryDateAnn = model.OpenQueryDateAnn,
					OpenQueryDate = model.OpenQueryDate
				});

				return (result > 0) ? "OK" : "Fail";
			}
		}

		/// <summary>
		/// 查詢調整起迄時間LOG
		/// </summary>
		/// <param name="model"></param>
		/// <returns></returns>
		public List<OpCalendarLog> QueryAdjDateTimeUpateLog(OpCalendar model)
		{
			string sql = @"select *
                           from OpCalendarLog 
                           where production_ym=@productionYM and sequence=@sequence and log_type <> 'S'
                           order by log_iden desc";
			List<OpCalendarLog> result = DbHelper.Query<OpCalendarLog>(EBrokerRepository.ConnectionStringName, sql, new
			{
				productionYM = model.ProductionYM,
				sequence = model.Sequence
			}).ToList();

			return result;
		}

		/// <summary>
		/// 刪除OpCalendar紀錄
		/// </summary>
		/// <param name="iden">自動識別碼</param>
		/// <param name="UserID"></param>
		/// <returns></returns>
		public bool DeleteOpCalendarByIden(string iden, string UserID)
		{
			string sql = @"update OpCalendar set update_datetime=getdate(),
                           update_user_code=@updateUserCode
                           where iden=@iden";
			DbHelper.Execute(EBrokerRepository.ConnectionStringName, sql, new
			{
				updateUserCode = UserID,
				iden = iden
			});
			int result = DbHelper.Execute(EBrokerRepository.ConnectionStringName, @"delete OpCalendar where iden=@iden", new
			{
				iden = iden
			});

			return (result > 0);
		}

		/// <summary>
		/// 取得OpCalendar
		/// </summary>
		/// <returns></returns>
		public List<OpCalendar> GetOpCalendar()
		{
			string sql = @"select * from OpCalendar order by production_ym desc,sequence desc";
			List<OpCalendar> op = new List<OpCalendar>();
			//取出資料
			return op = DbHelper.Query<OpCalendar>(EBrokerRepository.ConnectionStringName, sql).ToList();
		}

		/// <summary>
		/// 取得名字
		/// </summary>
		/// <returns></returns>
		public string GetScAccont(string iaccount)
		{
			string sql = @"select nmember from sc_account a join sc_member b on a.imember = b.imember
			where iaccount = @iaccount";

			//取出資料
			return DbHelper.Query<string>(EBrokerRepository.ConnectionStringName, sql, new { iaccount = iaccount }).FirstOrDefault();
		}

		#region 報表

		public Stream GetOpCalendarReportList(string productionYM)
		{
			string sql = @"select * from OpCalendar where production_ym like @productionYM order by production_ym desc,sequence desc";
			//取出資料
			List<OpCalendar> OpCalendar = DbHelper.Query<OpCalendar>(EBrokerRepository.ConnectionStringName, sql, new { productionYM = "%" +  productionYM + "%" }).ToList();

			MemoryStream ms = new MemoryStream();
			ExcelPackage excel = new ExcelPackage();
			ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("佣酬相關日期設定報表");
			//直排
			int a = 1, b = 2, c = 3, d = 4, e = 5, f = 6, g = 7, h = 8, j = 9, k = 10, l = 11, m = 12, n = 13, o = 14, p = 15;
			//橫排
			int aa = 3;

			#region 標題
			ExcelSetCell(sheet, new string[] { "佣酬相關日期設定報表" }, 1, 1);
			ExcelSetCell(sheet, new string[] { "" }, 1, 15);
			sheet.Cells[1, 1, 1, 15].Merge = true;
			sheet.Cells[1, 1, 1, 15].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			//sheet 標題 橫排 直排
			ExcelSetCell(sheet, new string[] { "業績年月" }, 2, a);
			sheet.Cells[2, a].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "次佣" }, 2, b);
			sheet.Cells[2, b].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "人事關檔日期" }, 2, c);
			sheet.Cells[2, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "RUN佣日期" }, 2, d);
			sheet.Cells[2, d].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "發佣日期" }, 2, e);
			sheet.Cells[2, e].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "簽收回條截止日" }, 2, f);
			sheet.Cells[2, f].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "調整起日" }, 2, g);
			sheet.Cells[2, g].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "調整迄日" }, 2, h);
			sheet.Cells[2, h].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "佣酬明細開放查詢日期" }, 2, j);
			sheet.Cells[2, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "年度績效報酬開放查詢日期" }, 2, k);
			sheet.Cells[2, k].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			//ExcelSetCell(sheet, new string[] { "計佣保費" }, 2, k);
			//sheet.Cells[2, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "建立者" }, 2, l);
			sheet.Cells[2, l].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "建立時間" }, 2, m);
			sheet.Cells[2, m].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "修改者" }, 2, n);
			sheet.Cells[2, n].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "修改時間" }, 2, o);
			sheet.Cells[2, o].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			ExcelSetCell(sheet, new string[] { "備註" }, 2, p);
			sheet.Cells[2, p].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

			#endregion

			#region 明細

			for (int i = 0; i < OpCalendar.Count; i++)
			{
				ExcelSetCell(sheet, new string[] { OpCalendar[i].ProductionYM }, aa, a);
				sheet.Cells[aa, a].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].Sequence }, aa, b);
				sheet.Cells[aa, b].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].HrCloseDate }, aa, c);
				sheet.Cells[aa, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].SalRunDate }, aa, d);
				sheet.Cells[aa, d].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].SalPayDate }, aa, e);
				sheet.Cells[aa, e].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].SalReceiptDate }, aa, f);
				sheet.Cells[aa, f].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].AdjDateTimeStr != null ? Convert.ToDateTime(OpCalendar[i].AdjDateTimeStr).ToString("yyyy/MM/dd HH:mm") : "" }, aa, g);
				sheet.Cells[aa, g].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].AdjDateTimeEnd != null ? Convert.ToDateTime(OpCalendar[i].AdjDateTimeEnd).ToString("yyyy/MM/dd HH:mm") : "" }, aa, h);
				sheet.Cells[aa, h].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].OpenQueryDate }, aa, j);
				sheet.Cells[aa, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].OpenQueryDateAnn }, aa, k);
				sheet.Cells[aa, k].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				//ExcelSetCell(sheet, new int[] { OpCalendar[i].CommModePrem }, aa, k);
				//sheet.Cells[aa, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				if (!String.IsNullOrEmpty(OpCalendar[i].CreateUserCode))
				{
					OpCalendar[i].CreateUserName = GetScAccont(OpCalendar[i].CreateUserCode);
				}
				if (!String.IsNullOrEmpty(OpCalendar[i].UpdateUserCode))
				{
					OpCalendar[i].UpdateUserName = GetScAccont(OpCalendar[i].UpdateUserCode);
				}

				ExcelSetCell(sheet, new string[] { OpCalendar[i].CreateUserName }, aa, l);
				sheet.Cells[aa, l].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].CreateDateTime }, aa, m);
				sheet.Cells[aa, m].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].UpdateUserName }, aa, n);
				sheet.Cells[aa, n].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].UpdateDateTime }, aa, o);
				sheet.Cells[aa, o].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				ExcelSetCell(sheet, new string[] { OpCalendar[i].Remark }, aa, p);
				sheet.Cells[aa, p].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

				aa++;
			}
			#endregion

			//ExcelSetCell(sheet, new int[] { TotalCommModePrem }, bb, k);
			//sheet.Cells[bb, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

			//設置列寬
			sheet.Column(1).Width = 10;
			//單元格自動適應大小
			//sheet.Cells.Style.ShrinkToFit = true;
			//字型
			sheet.Cells.Style.Font.Name = "微軟正黑體";
			//文字大小
			sheet.Cells.Style.Font.Size = 12;
			excel.SaveAs(ms);
			excel.Dispose();
			ms.Position = 0;
			return ms;
		}
		#endregion
		#endregion

		#region 資料填入Excel的Cell ExcelSetCell
		/// <summary>
		/// 資料填入Excel的Cell
		/// </summary>
		/// <param name="workSheet">sheet object</param>
		/// <param name="valueList">Data Array</param>
		/// <param name="rowStartPosition">Row Start Position</param>
		/// <param name="columnStartPosition">Column Start Position</param>
		public static void ExcelSetCell(ref ExcelWorksheet workSheet, object[] valueList, int rowStartPosition, int columnStartPosition)
		{
			foreach (var value in valueList)
			{
				// 全部資料範圍設定框線
				//workSheet.Cells[rowStartPosition, columnStartPosition].Style.Border.BorderAround(ExcelBorderStyle.Thin);
				workSheet.Cells[rowStartPosition, columnStartPosition++].Value = value;
			}
		}

		public static void ExcelSetCellStyle(ref ExcelWorksheet sheet, int rowStart = 1, int colStart = 1, int rowEnd = 1, int colEnd = 1, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
		{
			// 全部資料範圍設定框線
			sheet.Cells[rowStart, colStart, rowEnd, colEnd].Style.Border.BorderAround(borderStyle);

			// 自動調整欄位大小
			for (int i = 1; i <= colEnd; i++)
			{
				sheet.Column(i).AutoFit();
			}
		}

		public static void CellAutoFit(ref ExcelWorksheet sheet)
		{
			sheet.Cells.AutoFitColumns();
		}

		public static void CellFillBorderStyle(ref ExcelWorksheet sheet, int? rowStart = null, int? colStart = null, int? rowEnd = null, int? colEnd = null)
		{
			if (!rowStart.HasValue)
			{
				rowStart = 1;
			}
			if (!rowEnd.HasValue)
			{
				rowEnd = sheet.Cells.Rows;
			}

			if (!colStart.HasValue)
			{
				colStart = 1;
			}

			if (!colEnd.HasValue)
			{
				colEnd = sheet.Cells.Columns;
			}
			var cells = sheet.Cells[rowStart.Value, colStart.Value, rowEnd.Value, colEnd.Value];
			cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
			cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
			cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
			cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
			cells.AutoFitColumns();
		}

		#region EPPlus
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

		#endregion
	}
}
