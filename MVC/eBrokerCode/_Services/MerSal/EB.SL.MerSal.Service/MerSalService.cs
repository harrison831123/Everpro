using EB.Common;
using EB.EBrokerModels;
using EB.Platform.Service;
using EB.Platform.Service.Base;
using EB.SL.MerSal.Models;
using EB.VLifeModels;
using EB.WebBrokerModels;
using Microsoft.CUF.Framework.Data;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Drawing;
using Microsoft.CUF;

namespace EB.SL.MerSal.Service
{
    public class MerSalService : IMerSalService
    {
        #region 發佣檢核&紀錄
        /// <summary>
        /// 取得工作月
        /// </summary>
        /// <param name="strSel">query type</param>
        /// <param name="exclude88">是否排掉 sequence=88</param>
        /// <param name="selectTop">要取得的筆數</param>
        /// <returns></returns>
        public List<agym> GetYMData(string strSel, bool exclude88, int selectTop = 1)
        {
            string sql = @"select top " + selectTop + " production_ym,sequence from agym where 1=1";

            if (strSel == "agym")
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
        /// 取得發佣原始檔狀態
        /// </summary>
        /// <param name="ProductionYm"></param>
        /// <param name="Sequence"></param>
        /// <param name="CompanyCode"></param>
        /// <returns></returns>
        public string GetSeqNo(string ProductionYm, string Sequence, string CompanyCode)
        {
            string sql = @"
            select process_flag from RunSal 
            where seq_no in (	
                select MAX(seq_no) from RunSal 
                where production_ym=@ProductionYm and sequence=@Sequence
                and company_code=@CompanyCode and process_flag<>'D')";
            string result = DbHelper.Query<string>(WebBrokerRepository.ConnectionStringName, sql,
                new
                {
                    ProductionYm = ProductionYm,
                    Sequence = Sequence,
                    CompanyCode = CompanyCode
                }).FirstOrDefault();

            return result ?? "";
        }

        /// <summary>
        /// 取得保公
        /// </summary>
        /// <returns></returns>
        public List<TermVal> GetCompanyCode()
        {
            string sql = @"
            select term_code,term_meaning 
            from trmval 
            where term_id='company_code' and term_code LIKE '%[0-9]%' 
            order by term_code";
            List<TermVal> result = DbHelper.Query<TermVal>(VLifeRepository.ConnectionStringName, sql).ToList();

            return result;
        }

        /// <summary>
        /// 佣酬檢核狀態為「4-資料入發佣」
        /// </summary>
        /// <param name="ProductionYm"></param>
        /// <param name="Sequence"></param>
        /// <returns></returns>
        public string Getprocess_status_MSR(string ProductionYm, string Sequence)
        {
            string sql = @"
            select distinct process_status_MSR from MerSalrun 
            where production_ym=@ProductionYm and sequence=@Sequence 
            and process_status_MSR='4'";

            ////5 ★測試
            //string sql = @"
            //select distinct process_status_MSR from MerSalrun 
            //where production_ym='0111/11' and sequence='1' 
            //and process_status_MSR='4'";

            string result = DbHelper.Query<string>(WebBrokerRepository.ConnectionStringName, sql,
                new
                {
                    ProductionYm = ProductionYm,
                    Sequence = Sequence,
                }).FirstOrDefault();

            return result;
        }

        /// <summary>
        /// 佣酬檢核狀態為「3-資料入佣酬調整
        /// </summary>
        /// <param name="ProductionYm"></param>
        /// <param name="Sequence"></param>
        /// <returns></returns>
        public string Getprocess_status_MSRC(string ProductionYm, string Sequence, string CompanyCode)
        {
            string sql = @"
            select distinct process_status_MSR from MerSalrun
            where production_ym=@ProductionYm and sequence=@Sequence
            and company_code=@CompanyCode 
            and process_status_MSR='3'";
            string result = DbHelper.Query<string>(WebBrokerRepository.ConnectionStringName, sql,
                new
                {
                    ProductionYm = ProductionYm,
                    Sequence = Sequence,
                    CompanyCode = CompanyCode
                }).FirstOrDefault();

            return result;
        }

        /// <summary>
        /// Batch中有狀態為R且時間不超過3分鐘
        /// </summary>
        /// <param name="ProductionYm"></param>
        /// <param name="Sequence"></param>
        /// <param name="CompanyCode"></param>
        /// <returns></returns>
        public string Getcm_batch_controlR(string ProductionYm, string Sequence, string CompanyCode)
        {
            //條件
            //1. 小於現在
            //2. 如果dend是NULL,那和現在時間比要小於3分鐘==>可能死在R
            string sql = @"
            select count(*) cnt
            from cm_batch_control
            where nreport_name='佣酬檢核報表'+replace(@ProductionYm,'/','')+@Sequence+'-'+@CompanyCode
            and fbatch_status='R'
            and (
	            dstart<=dateadd(n,1,GETDATE())
            or (dend IS NULL and DATEDIFF(n,dstart,GETDATE())<=3)
            )

            ";

            int dataCount = DbHelper.Query<int>(EBrokerRepository.ConnectionStringName, sql, new
            {
                ProductionYm = ProductionYm,
                Sequence = Sequence,
                CompanyCode = CompanyCode
            }).FirstOrDefault();
            return dataCount.ToString() ?? "";

        }

        /// <summary>
        /// 檢核執行紀錄
        /// </summary>
        /// <param name="ProductionYm"></param>
        /// <param name="Sequence"></param>
        /// <param name="CompanyCode"></param>
        /// <returns></returns>
        public List<MerSalViewModel> GetMerSalRun(string ProductionYm, string Sequence, string CompanyCode)
        {
            string sql = @"
            select a.production_ym,a.sequence,a.company_code
            ,a.imp_count,a.err_count,a.war_count,a.formal_count,a.cut_count
            ,tv2.term_meaning ag_fromName
            ,tv1.term_meaning process_status_MSRName
            ,a.data_checktime_s,a.data_checktime_e,a.create_user_code
            from MerSalRun a
            left join trmval tv1 on a.process_status_MSR =tv1.term_code and tv1.term_id='process_status_MSR'
            left join trmval tv2 on a.ag_from =tv2.term_code and tv2.term_id='ag_from'
            where production_ym=@ProductionYm
            and sequence=@Sequence
            and company_code=@CompanyCode
            order by run_seq desc
            ";
            List<MerSalViewModel> result = DbHelper.Query<MerSalViewModel>(WebBrokerRepository.ConnectionStringName, sql,
                new
                {
                    ProductionYm = ProductionYm,
                    Sequence = Sequence,
                    CompanyCode = CompanyCode
                }).ToList();

            return result;
        }

        /// <summary>
        /// 目前人事關檔年月-次薪
        /// </summary>
        /// <returns></returns>
        public string GetAgymIndOne()
        {
            string sql = @"
            select production_ym+'-'+convert(varchar,sequence) 
            from agym
            where production_ym in (select MAX(production_ym) production_ym from agym where agym_ind='1' and sequence<>88)
            and sequence in (
                select MAX(sequence) from agym 
                where agym_ind='1' 
                and production_ym in (select MAX(production_ym) production_ym from agym where agym_ind='1' and sequence<>88)
                and sequence<>88)";

            ////1 ★測試
            //string sql = @"select '0111/11-1'";

            string result = DbHelper.Query<string>(VLifeRepository.ConnectionStringName, sql).FirstOrDefault();
            return result ?? "";
        }

        /// <summary>
        /// 目前業績關檔年月-次薪
        /// </summary>
        /// <returns></returns>
        public string GetAgbcIndOne()
        {
            string sql = @"
            select production_ym+'-'+convert(varchar,sequence) from agym
            where production_ym in (select MAX(production_ym) production_ym from agym where agym_ind='1' and agbc_ind='1' and sequence<>88)
            and sequence in (
                select MAX(sequence) from agym
                where agym_ind='1' and agbc_ind='1'
                and production_ym in (select MAX(production_ym) production_ym from agym where agym_ind='1' and agbc_ind='1' and sequence<>88)
                and sequence<>88)";

            ////4 ★測試
            //string sql = @"select '0111/10-2'";

            string result = DbHelper.Query<string>(VLifeRepository.ConnectionStringName, sql).FirstOrDefault();
            return result ?? "";
        }

        /// <summary>
        /// 目前工作月
        /// </summary>
        /// <returns></returns>
        public string GetProductionYm()
        {
            string sql = @"select MAX(production_ym) production_ym from agym where sequence<>88";

            ////2 ★測試
            //string sql = @"select '0111/11'";

            string result = DbHelper.Query<string>(VLifeRepository.ConnectionStringName, sql).FirstOrDefault();
            return result ?? "";
        }

        /// <summary>
        /// 目前次薪
        /// </summary>
        /// <returns></returns>
        public string GetSequence()
        {
            string sql = @"select MAX(sequence) sequence from agym where production_ym in (select MAX(production_ym) production_ym from agym where sequence<>88)";

            ////3 ★測試
            //string sql = @"select '1'";

            short result = DbHelper.Query<short>(VLifeRepository.ConnectionStringName, sql).FirstOrDefault();
            return result.ToString() ?? "";
        }

        /// <summary>
        /// 執行SQL(無回傳值)
        /// </summary>
        /// <param name="sql"></param>
        public void ExcuteWebBrokerSQL(string sql)
        {
            DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql);
        }
        #endregion

        #region 發佣即時批次
        public Stream BatchQueryReport(string jsonParams)
        {
            JObject stuff = JObject.Parse(JsonConvert.DeserializeObject(jsonParams).ToString());
            //@ProductionYm, @Sequence, @CompanyCode, @FileSeq, @CreateUserCode
            ExcuteWebBrokerSQL(string.Format("exec  usp_MerSalCheckProcess01 '{0}','{1}','{2}','{3}','{4}'", stuff["ProductionYm"].ToString(),
                stuff["Sequence"].ToString(), stuff["CompanyCode"].ToString(), stuff["FileSeq"].ToString(), stuff["CreateUserCode"].ToString()));

            //取得報表的資料
            var dataList = DbHelper.Query<MerSalReportModel>(WebBrokerRepository.ConnectionStringName,
                @"select iden,run_seq,production_ym,sequence,company_code,file_seq,policy_no2,chk_code,chk_code_type
                ,error_type_msg,error_value,create_datetime,c.nmember
                from MerSalCheckError a join [SL_EBROKERDB_EBROKER].EBROKER.dbo.sc_account b on a.create_user_code = b.iaccount
				join [SL_EBROKERDB_EBROKER].EBROKER.dbo.sc_member c on b.imember = c.imember
                where run_seq in(
                SELECT MAX(run_seq) 
                FROM MerSalRun 
                WHERE production_ym=@ProductionYm AND sequence=@Sequence 
                AND company_code=@CompanyCode AND file_seq=@FileSeq 
                AND process_status_MSR<>'D')
                order by iden", new
                {
                    Sequence = stuff["Sequence"].ToString(),
                    ProductionYm = stuff["ProductionYm"].ToString(),
                    FileSeq = stuff["FileSeq"].ToString(),
                    CompanyCode = stuff["CompanyCode"].ToString(),
                }).ToList();

            ExcelPackage excel = new ExcelPackage();

            //設定Excel中的Sheet名稱
            ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("佣酬檢核訊息");

            int row = 1;

            #region 設定報表欄位表頭
            var titles = new object[] { "序號", "工作月", "序號", "保險公司", "檔案序號",
                                        "保單號碼", "檢核代碼", "錯誤代碼", "錯誤訊息", "錯誤資料顯示",
                                        "執行時間", "執行人員" };
            // 設定報表欄位表頭
            ExcelSetCell(ref sheet, titles, row, 1);
            #endregion

            #region 設定報表內容
            //設定報表內容
            foreach (var item in dataList)
            {
                var valueList = new object[] { item.RunSeq, item.ProductionYm, item.Sequence, item.CompanyCode, item.FileSeq, item.PolicyNo2,
                    item.ChkCode, item.ChkCodeType, item.ErrorTypeMsg, item.ErrorValue,item.CreateDatetime, item.nmember };
                row++;

                ExcelSetCell(ref sheet, valueList, row, 1);

            }

            //if (dataList.Any())
            //{
            //    sheet.Cells[2, 32, row, 38].Style.Numberformat.Format = "#,##0";
            //}

            CellFillBorderStyle(ref sheet, 1, 1, row, titles.Length);

            #endregion

            var result = new MemoryStream();
            excel.SaveAs(result);
            excel.Dispose();
            result.Position = 0;

            return result;

        }
        #endregion

        #region 佣收報表
        /// <summary>
        /// 入佣資料報表
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public byte[] GetMerSalDReportList(MerSalViewModel condition)
        {
            //EXEC usp_MerSalDRpt @ProductionYm,@Sequence,@CompanyCode, 'Detail'
            string sql = @"EXEC usp_MerSalDRpt @ProductionYm,@Sequence,@CompanyCode, Detail";

            List<MerSalDReportModel> MerSalReport = new List<MerSalDReportModel>();
            //取出資料
            MerSalReport = DbHelper.Query<MerSalDReportModel>(WebBrokerRepository.ConnectionStringName, sql, new
            {
                ProductionYM = condition.ProductionYM,
                Sequence = condition.Sequence.ToString(),
                CompanyCode = condition.CompanyCode

            }).ToList();

            //匯出excel格式
            MemoryStream ms = null;
            string HeadName = "入佣資料報表                                                         查詢人員：" + condition.QueryUser+ "   查詢日期：" +condition.QueryDate;
            List<EB.Common.CustHeader> headList = new List<EB.Common.CustHeader>();
            CustHeader c1 = new CustHeader();
            c1.CellsNum = new List<int>() { 45};
            c1.CellsObj = new ArrayList() { HeadName };

            CustHeader c2 = new CustHeader();
            c2.CellsNum = new List<int> { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1
                , 1, 1, 1, 1 ,1, 1, 1, 1, 1, 1
                , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1
                , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1
                , 1, 1, 1, 1, 1};
            c2.CellsObj = new ArrayList() { "序號","工作月", "檔案序號", "序次", "保公代碼", "保單號碼", "永達-繳別繳次", "保公-繳別", "保公-繳次"
                , "保公-年度", "永達-年期", "保公-險種代碼(無年期)", "保公-險種代碼(有年期)", "保額", "保費","保公佣金", "金額比率","金額類別"
                ,"永達佣金率","業績換算率","永達佣金", "永達原因碼","保公原因碼", "生效日", "保公版次", "保險年齡", "被保人", "被保人id"
                , "主附約碼", "支票日期", "國寶D1保費", "主約年期", "費用類別", "保險類別註記", "替換件註記" , "職業等級", "集彙代碼"
                , "系統執行項目(業行)","發放註記", "業務員ID1", "業務員證號1", "業務員ID2", "業務員證號2","half件","助理簽收日期"};
            headList.AddRange(new CustHeader[] { c1, c2 });
            if (MerSalReport.Count != 0)
            {
                ms = NPoiHelper.Export(MerSalReport, headList, HeadName);
                return ms.ToArray();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 保險公司佣酬檢核表
        /// </summary>
        /// <param name="year">年度</param>
        /// <returns>Stream</returns>
        public Stream GetCompanyMerSalDReportList(MerSalViewModel condition)
        {
            MemoryStream ms = new MemoryStream();
            ExcelPackage excel = new ExcelPackage();

            #region SQL資料
            string WebBroker = WebBrokerRepository.DbInfo.ConnectionString;
            DataSet ds = new DataSet();
            using (SqlConnection sqlcon = new SqlConnection(WebBroker))
            {
                using (SqlCommand cmd = new SqlCommand("usp_MerSalDRpt", sqlcon))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add("@ProductionYm", SqlDbType.VarChar).Value = condition.ProductionYM;
                    cmd.Parameters.Add("@Sequence", SqlDbType.VarChar).Value = condition.Sequence;
                    cmd.Parameters.Add("@CompanyCode", SqlDbType.VarChar).Value = condition.CompanyCode ?? "";
                    cmd.Parameters.Add("@Rpttype", SqlDbType.VarChar).Value = "Item";

                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(ds);
                    }
                }
            }

            // 定義集合    
            List<MerSalSystemReportModel> merSalSystemReports = new List<MerSalSystemReportModel>();
            List<MerSalCompanyReportModel> merSalCompanyReports = new List<MerSalCompanyReportModel>();
            MerSalSystemReportModel t = new MerSalSystemReportModel();
            MerSalCompanyReportModel t2 = new MerSalCompanyReportModel();
            PropertyInfo[] prop = t.GetType().GetProperties();
            PropertyInfo[] prop2 = t2.GetType().GetProperties();

            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                t = new MerSalSystemReportModel();
                //通過反射獲取T類型的所有成員
                foreach (PropertyInfo pi in prop)
                {
                    //DataTable列名=屬性名
                    if (ds.Tables[0].Columns.Contains(pi.Name))
                    {
                        //屬性值不為空
                        if (dr[pi.Name] != DBNull.Value)
                        {
                            object value = Convert.ChangeType(dr[pi.Name], pi.PropertyType);
                            //給T類型字段賦值
                            pi.SetValue(t, value, null);
                        }
                    }
                }
                //將T類型添加到集合list
                merSalSystemReports.Add(t);
            }

            foreach (DataRow dr in ds.Tables[1].Rows)
            {
                t2 = new MerSalCompanyReportModel();
                //通過反射獲取T類型的所有成員
                foreach (PropertyInfo pi in prop2)
                {
                    //DataTable列名=屬性名
                    if (ds.Tables[1].Columns.Contains(pi.Name))
                    {
                        //屬性值不為空
                        if (dr[pi.Name] != DBNull.Value)
                        {
                            object value = Convert.ChangeType(dr[pi.Name], pi.PropertyType);
                            //給T類型字段賦值
                            pi.SetValue(t2, value, null);
                        }
                    }
                }
                //將T類型添加到集合list
                merSalCompanyReports.Add(t2);
            }
            #endregion

            //ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("保險公司佣酬檢核報表");
            ExcelWorksheet sheet = excel.Workbook.Worksheets.Add(condition.CompanyCode ?? "all");

            //直排
            int a = 1, b = 2, c = 3, d = 4, e = 5, f = 6, g = 7, h = 8, j = 9, k = 10, l = 11, m = 12, n = 13, o = 14, p = 15, q = 16, r = 17;
            //橫排
            int aa = 3, bb = 15;

            #region 報表抬頭-系統執行項目
            for (int i = 0; i < merSalSystemReports.Count; i++)
            {
                //佣酬類別
                ExcelSetCell(sheet, new string[] { merSalSystemReports[i].amount_type }, aa, a);
                sheet.Cells[aa, a].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //年期
                ExcelSetCell(sheet, new string[] { merSalSystemReports[i].modx_year_name }, aa, b);
                sheet.Cells[aa, b].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //系統(保留)筆數
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].Cnt_Cut00) }, aa, c);
                sheet.Cells[aa, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, c].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //系統(保留)保費
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].ModePrem_Cut00) }, aa, d);
                sheet.Cells[aa, d].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, d].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //系統(保留)佣金
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].CommPrem_Cut00) }, aa, e);
                sheet.Cells[aa, e].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, e].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //系統(人工調帳)筆數
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].Cnt_Cut01) }, aa, f);
                sheet.Cells[aa, f].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, f].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //永達-繳別繳次系統 (人工調帳)保費
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].ModePrem_Cut01) }, aa, g);
                sheet.Cells[aa, g].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, g].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //系統(人工調帳)佣金
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].CommPrem_Cut01) }, aa, h);
                sheet.Cells[aa, h].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, h].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //轉入佣酬計算筆數
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].Cnt_Check) }, aa, j);
                sheet.Cells[aa, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, j].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //轉入佣酬計算保費
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].ModePrem_Check) }, aa, k);
                sheet.Cells[aa, k].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, k].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //轉入佣酬計算佣金
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].CommPrem_Check) }, aa, l);
                sheet.Cells[aa, l].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, l].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //系統(不處理)筆數
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].Cnt_N) }, aa, m);
                sheet.Cells[aa, m].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, m].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //系統(不處理)保費
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].ModePrem_N) }, aa, n);
                sheet.Cells[aa, n].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, n].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //系統(不處理)佣金
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].CommPrem_N) }, aa, o);
                sheet.Cells[aa, o].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, o].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //合計筆數
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].Cnt_Total) }, aa, p);
                sheet.Cells[aa, p].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, p].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //合計保費
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].ModePrem_Total) }, aa, q);
                sheet.Cells[aa, q].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, q].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //合計佣金
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalSystemReports[i].CommPrem_Total) }, aa, r);
                sheet.Cells[aa, r].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[aa, r].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                aa++;
            }
            #endregion

            #region 報表抬頭-保險公司佣酬檢核表
            for (int i = 0; i < merSalCompanyReports.Count; i++)
            {
                //工作月
                ExcelSetCell(sheet, new string[] { merSalCompanyReports[i].production_ym }, bb, a);
                sheet.Cells[bb, a].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //次薪
                ExcelSetCell(sheet, new string[] { merSalCompanyReports[i].sequence }, bb, b);
                sheet.Cells[bb, b].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //保險公司
                ExcelSetCell(sheet, new string[] { merSalCompanyReports[i].company_code }, bb, c);
                sheet.Cells[bb, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //筆數
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalCompanyReports[i].cnt) }, bb, d);
                sheet.Cells[bb, d].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[bb, d].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //保費
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalCompanyReports[i].mode_prem) }, bb, e);
                sheet.Cells[bb, e].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[bb, e].Style.Numberformat.Format = "#,##0;[Red]-#,##0";
                //佣金
                ExcelSetCell(sheet, new int[] { Convert.ToInt32(merSalCompanyReports[i].comm_prem) }, bb, f);
                sheet.Cells[bb, f].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[bb, f].Style.Numberformat.Format = "#,##0;[Red]-#,##0";

                bb++;
            }
            #endregion

            #region 標題
            //sheet 標題 橫排 直排
            ExcelSetCell(sheet, new string[] { "系統執行項目" }, 1, 1);
            ExcelSetCell(sheet, new string[] { "" }, 1, 17);
            sheet.Cells[1, 1, 1, 17].Merge = true;
            sheet.Cells[1, 1, 1, 17].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //sheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //sheet.Cells[1, 1].Style.Font.Color.SetColor(Color.Blue);

            ExcelSetCell(sheet, new string[] { "佣酬類別" }, 2, 1);
            ExcelSetCell(sheet, new string[] { "年期" }, 2, 2);
            ExcelSetCell(sheet, new string[] { "系統(保留)筆數" }, 2, 3);
            ExcelSetCell(sheet, new string[] { "系統(保留)保費" }, 2, 4);
            ExcelSetCell(sheet, new string[] { "系統(保留)佣金" }, 2, 5);
            ExcelSetCell(sheet, new string[] { "系統(人工調帳)筆數" }, 2, 6);
            sheet.Cells[2, 7].Style.WrapText = true;//自动换行
            ExcelSetCell(sheet, new string[] { "永達-繳別繳次系統\n(人工調帳)保費" }, 2, 7);
            ExcelSetCell(sheet, new string[] { "系統(人工調帳)佣金" }, 2, 8);
            ExcelSetCell(sheet, new string[] { "轉入佣酬計算筆數" }, 2, 9);
            ExcelSetCell(sheet, new string[] { "轉入佣酬計算保費" }, 2, 10);
            ExcelSetCell(sheet, new string[] { "轉入佣酬計算佣金" }, 2, 11);
            ExcelSetCell(sheet, new string[] { "系統(不處理)筆數" }, 2, 12);
            ExcelSetCell(sheet, new string[] { "系統(不處理)保費" }, 2, 13);
            ExcelSetCell(sheet, new string[] { "系統(不處理)佣金" }, 2, 14);
            ExcelSetCell(sheet, new string[] { "合計筆數" }, 2, 15);
            ExcelSetCell(sheet, new string[] { "合計保費" }, 2, 16);
            ExcelSetCell(sheet, new string[] { "合計佣金" }, 2, 17);
            
            //系統執行項目的表頭格式
            for (int y = 1; y <= 17; y++)
            {
                sheet.Cells[2, y].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[2, y].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);//设置单元格背景色
                sheet.Cells[2, y].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
            }

            ExcelSetCell(sheet, new string[] { "保險公司佣酬檢核表" }, 13, 1);
            ExcelSetCell(sheet, new string[] { "" }, 13, 2);
            ExcelSetCell(sheet, new string[] { "" }, 13, 3);
            ExcelSetCell(sheet, new string[] { "" }, 13, 4);
            ExcelSetCell(sheet, new string[] { "" }, 13, 5);
            ExcelSetCell(sheet, new string[] { "" }, 13, 6);
            sheet.Cells[13, 1, 13, 6].Merge = true;////合并单元格
            sheet.Cells[13, 1, 13, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中

            ExcelSetCell(sheet, new string[] { "工作月" }, 14, 1);
            ExcelSetCell(sheet, new string[] { "次薪" }, 14, 2);
            ExcelSetCell(sheet, new string[] { "保險公司" }, 14, 3);
            ExcelSetCell(sheet, new string[] { "筆數" }, 14, 4);
            ExcelSetCell(sheet, new string[] { "保費" }, 14, 5);
            ExcelSetCell(sheet, new string[] { "佣金" }, 14, 6);

            //保險公司佣酬檢核表的表頭格式
            for (int y = 1; y <= 6; y++)
            {
                sheet.Cells[14, y].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[14, y].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                sheet.Cells[14, y].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            #endregion

            sheet.Column(1).Width = 20;
            sheet.Cells.Style.ShrinkToFit = true;
            //字型
            sheet.Cells.Style.Font.Name = "微軟正黑體";
            //文字大小
            sheet.Cells.Style.Font.Size = 10;
            excel.SaveAs(ms);
            excel.Dispose();
            ms.Position = 0;

            if (merSalSystemReports.Count != 0 || merSalCompanyReports.Count != 0)
            {
                return ms;
            }
            else
            {
                return Stream.Null;
            }
        }

        #endregion

        #region 佣酬發放調整----------------------------------------------------------S

        /// <summary>
        /// 目前業績工作月
        /// </summary>
        /// <returns></returns>
        public string GetProductionYmNow()
        {
            string sql = @"
            select production_ym+'-'+convert(varchar,sequence) 
            from (
	            select production_ym,max(sequence) sequence 
                from agym
	            where production_ym in (select max(production_ym) production_ym from agym)
	            group by production_ym
            )a";
            string result = DbHelper.Query<string>(VLifeRepository.ConnectionStringName, sql).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 入Run薪工作月
        /// </summary>
        /// <returns></returns>
        public string GetProductionYmRun()
        {
            string sql = @"
            select production_ym+'-'+convert(varchar,sequence) PMS_Comm 
            from WebBroker.dbo.MerSalRun 
            where production_ym=(select max(production_ym) production_ym from agym)
            and sequence=(
                select max(sequence) sequence 
                from agym 
	            where production_ym in (select max(production_ym) production_ym from agym) 
                group by production_ym
             )
            and process_status_MSR='4' 
            GROUP BY production_ym+'-'+convert(varchar,sequence)
            ";
            string result = DbHelper.Query<string>(VLifeRepository.ConnectionStringName, sql).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 佣酬類別
        /// </summary>
        /// <returns></returns>
        public List<Trmval> GetAmountType()
        {
            string sql = @"select * from trmval where term_id='amount_type' ORDER BY term_sequence";
            List<Trmval> result = DbHelper.Query<Trmval>(WebBrokerRepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 檢核特殊資料類別
        /// </summary>
        /// <returns></returns>
        public List<Trmval> GetChkCodeActType()
        {
            string sql = @"select * from trmval where term_id='chkcode_acttype' ORDER BY term_sequence ";
            List<Trmval> result = DbHelper.Query<Trmval>(WebBrokerRepository.ConnectionStringName, sql).ToList();
            return result;
        }

        ///// <summary>
        ///// 發放註記
        ///// </summary>
        ///// <returns></returns>
        //public List<Trmval> GetPayType()
        //{
        //    string sql = @"select * from trmval where term_id='pay_type' ORDER BY term_sequence";
        //    List<Trmval> result = DbHelper.Query<Trmval>(WebBrokerRepository.ConnectionStringName, sql).ToList();
        //    return result;
        //}

        /// <summary>
        /// 佣酬發放調整查詢
        /// </summary>
        /// <returns></returns>
        public List<MerSalCheckViewModel> GetMerSalCheck(MerSalCheckViewModel model)
        {
            #region SQL

            //paytype_show---------------------------------------------------------------------------------------------------------
            //case
            //     1.目前年月NULL==>X(不可能沒有agym資料)
            //當工作月入的資料
            //     2.列YM=目前年月 & 目前年月<>人事關檔年月                ==>X[人事未關先不用改，改了也沒用，會重跑]
            //     3.列YM=目前年月 & 已入發薪EBroker                      ==>X
            //     4.列YM=目前年月 & 目前年月= 人事關檔年月 & 未入發薪      ==>Y
            //先前工作月入的資料
            //     5.列YM<目前年月 & 目前年月<>人事關檔年月                ==>X只要是人事還沒關就不能做更改 #536琍婷要求
            //     6.列YM<目前年月 & NotPay                              ==>Y
            //     7.列YM<目前年月 & Pay & pay_month=目前年月 & 未入發薪   ==>Y
            //     8.列YM<目前年月 & Pay                                 ==>X
            //未來工作月入的資料...
            //      .列YM>目前年月 & Pay                                 ==>N[修改列年月比agym大...應該不可能吧]刪除
            //paytype_show---------------------------------------------------------------------------------------------------------

            //            ,main.memo
            string sql = $@"
            SELECT main.iden
            ,main.production_ym,main.sequence
            ,main.company_code,main.file_seq
            ,main.amount_type+'-'+rtrim(tv1.term_meaning) amount_type
            ,main.policy_no2,main.plan_code,main.collect_year,main.modx_sequence
            ,po_issue_date,age
            ,case when isnull(main.reason_code_c,'')<>'' and isnull(rc.reason_code_cname,'')='' then main.reason_code_c
                  when isnull(main.reason_code_c,'')<>'' and isnull(rc.reason_code_cname,'')<>'' then main.reason_code_c+'-'+rc.reason_code_cname else '' 
             end reason_code_c
            ,main.mode_prem,main.comm_prem_c,main.amount,main.comm_prem
            ,nn1.names agent_name1,main.agent_code1,nn2.names agent_name2,main.agent_code2
            ,rtrim(vp.policy_status_name) policy_status_name
            ,pt1.receipt_date
            ,case when isnull(main.pay_month,'')<>'' and isnull(main.pay_seq,'')<>'' then main.pay_month+'-'+main.pay_seq else '' end pay_ym
            ,main.pay_type,main.pay_type+'-'+rtrim(tv2.term_meaning) pay_type_name
            ,main.rpt_include_flag
            ,main.remark
            ,main.memo
            ,main.not_pay_YMS
            ,case when main.ag_cnt='2' then 'Y' else '' end 'ag_cnt'
            ,CASE	
	            WHEN '{model.YMNow}' IS NULL THEN 'X'
                WHEN main.production_ym+CONVERT(CHAR(1),main.sequence)='{model.YMNow}' AND '{model.YMNow}'<>'{model.YmAgymCloseNow}' THEN 'X'
                WHEN main.production_ym+CONVERT(CHAR(1),main.sequence)='{model.YMNow}' AND ISNULL(RunMSR.process_status_MSR,'')= '4' THEN 'X'
	            WHEN main.production_ym+CONVERT(CHAR(1),main.sequence)='{model.YMNow}' AND '{model.YMNow}'= '{model.YmAgymCloseNow}' AND ISNULL(RunMSR.process_status_MSR,'')<>'4' THEN 'Y'
                WHEN main.production_ym+CONVERT(CHAR(1),main.sequence)<'{model.YMNow}' AND '{model.YMNow}'<>'{model.YmAgymCloseNow}' THEN 'X'
	            WHEN main.production_ym+CONVERT(CHAR(1),main.sequence)<'{model.YMNow}' AND LEFT(main.pay_type,6)= 'NotPay' THEN 'Y'
                WHEN main.production_ym+CONVERT(CHAR(1),main.sequence)<'{model.YMNow}' AND LEFT(main.pay_type,6)<>'NotPay' AND isnull(main.pay_month,'')+CONVERT(CHAR(1),isnull(main.pay_seq,''))='{model.YMNow}' AND ISNULL(RunMSRPay.process_status_MSR,'')<>'4' THEN 'Y'
	            WHEN main.production_ym+CONVERT(CHAR(1),main.sequence)<'{model.YMNow}' AND LEFT(main.pay_type,6)<>'NotPay' THEN 'X'
	            ELSE '-'
            END paytype_show
            FROM MerSalCheck main
            LEFT JOIN trmval tv1 on main.amount_type=tv1.term_code and tv1.term_id='amount_type'
            LEFT JOIN trmval tv2 on main.pay_type=tv2.term_code and tv2.term_id='pay_type'
            LEFT JOIN (			
                select company_code,file_seq,amount_type,reason_code_c,reason_code_cname,minus_flag	from MerSalCommReasonCode where is_delete='N'
			    group by company_code,file_seq,amount_type,reason_code_c,reason_code_cname,minus_flag
            )rc on main.amount_type=rc.amount_type
                and main.company_code=rc.company_code 
                and main.file_seq=rc.file_seq 
                and main.reason_code_c=rc.reason_code_c 
                and rc.minus_flag = CASE WHEN rc.minus_flag <> '' THEN (CASE WHEN main.amount < 0 THEN 'Y' ELSE 'N' END) ELSE rc.minus_flag END
                and rc.reason_code_cname<>''
            LEFT JOIN (SELECT * FROM MerSalRun a2 WHERE process_status_MSR<>'D') RunMSR ON main.production_ym=RunMSR.production_ym and main.sequence=RunMSR.sequence and main.company_code=RunMSR.company_code and main.file_seq=RunMSR.file_seq
            LEFT JOIN (SELECT * FROM MerSalRun a3 WHERE process_status_MSR<>'D') RunMSRPay ON main.pay_month=RunMSRPay.production_ym and main.pay_seq=RunMSRPay.sequence and main.company_code=RunMSRPay.company_code and main.file_seq=RunMSRPay.file_seq
            LEFT JOIN vlife.dbo.podt_1 pt1 on main.policy_no=pt1.policy_no
            LEFT JOIN MerSalVisualPolicy vp on main.company_code=vp.company_code AND main.policy_no=vp.policy_no
			LEFT JOIN vlife.dbo.nain nn1 on left(main.agent_code1,10)=nn1.client_id
			LEFT JOIN vlife.dbo.nain nn2 on left(main.agent_code2,10)=nn2.client_id
            WHERE 1 = 1";

            if (!String.IsNullOrEmpty(model.ProductionYM))
            {
                sql += " AND main.production_ym=@ProductionYm";
            }
            if (model.Sequence != 0)
            {
                sql += " AND main.sequence=@Sequence";
            }
            if (!String.IsNullOrEmpty(model.PayMonth))
            {
                sql += " AND main.pay_month=@PayMonth";
            }
            if (!String.IsNullOrEmpty(model.PaySeq))
            {
                sql += " AND main.pay_seq=@PaySeq";
            }
            if (!String.IsNullOrEmpty(model.NotPayYMSS))
            {
                sql += " AND main.not_pay_YMS>=@NotPayYMSS";
            }
            if (!String.IsNullOrEmpty(model.NotPayYMSE))
            {
                sql += " AND main.not_pay_YMS<=@NotPayYMSE";
            }
            if (model.CompanyCode != "0")
            {
                sql += " AND main.company_code=@CompanyCode";
            }
            if (!String.IsNullOrEmpty(model.FileSeq))
            {
                sql += " AND main.file_seq=@FileSeq";
            }
            if (!String.IsNullOrEmpty(model.PolicyNo2))
            {
                sql += " AND main.policy_no2=@PolicyNo2";
            }
            if (!String.IsNullOrEmpty(model.PlanCode))
            {
                sql += " AND main.plan_code=@PlanCode";
            }
            if (!String.IsNullOrEmpty(model.AmountType))
            {
                sql += " AND main.amount_type=@AmountType";
            }
            if (!String.IsNullOrEmpty(model.PayType))
            {
                if (model.PayType == "Pay" )
                {
                    sql += " AND main.pay_type=@PayType";
                }
                else if (model.PayType == "NotPay-00")//小琍婷來電說簽未回不含不結轉下期 20240327 
                {
                    sql += " AND main.pay_type=@PayType AND main.rpt_include_flag<>'N' ";
                }
                else
                {
                    //sql += " AND main.pay_type not in ('NotPay-00','Pay')";
                    sql += " AND main.pay_type not in ('Pay')";//不發放-ALL含簽未回
                }
                
            }

            sql += " ORDER BY main.policy_no2,main.production_ym,main.sequence,main.plan_code,main.amount_type";
            #endregion

            List<MerSalCheckViewModel> result = DbHelper.Query<MerSalCheckViewModel>(WebBrokerRepository.ConnectionStringName, sql
                , new
                {
                    model.ProductionYM,
                    model.Sequence,
                    model.CompanyCode,
                    model.FileSeq,
                    model.PolicyNo2,
                    model.PlanCode,
                    model.AmountType,
                    model.PayType,
                    model.PayMonth,
                    model.PaySeq,
                    model.NotPayYMSS,
                    model.NotPayYMSE
                }).ToList();
            return result;
        }

        /// <summary>
        /// 更新佣酬發放調整
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public bool UpdateMerSalCheck(MerSalCheckViewModel model, string sA1, string sA2, string sA3)
        {
            string sql = "";
            sql = @" UPDATE MerSalCheck SET process_user_code=@ProcessUserCode";

            //只改1 ==>修改 佣酬發放
            if (sA1 != "X1")                                
            {
                sql += @" ,check_ind=@CheckInd";
                sql += @" ,pay_type = @PayType";
                sql += @" ,pay_month = @PayMonth";
                sql += @" ,pay_seq = @PaySeq";
                sql += @" ,not_pay_YMS=@NotPayYMS";
                sql += @" ,chk_code=chk_code+@ChkCode";
            }

            //改2 ==>修改 結轉下期
            if (sA2 != "X2")                 
            {
                sql += @" ,rpt_include_flag=@RptIncludeFlag";
            }

            //改3 ==>修改 人工備註
            if (sA3 != "X3")  
            {
                sql += @" ,memo=@Memo";
            }

            sql += @" WHERE Iden = @Iden";
            int result = -1;
            try
            {
                if (sA1 != "X1" && sA2 != "X2" && sA3 != "X3") //修改 1 VVV
                {
                    result = DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql, new
                    {
                        Iden = model.Iden,
                        CheckInd = model.CheckInd,
                        PayType = model.PayType,
                        PayMonth = model.PayMonth,
                        PaySeq = model.PaySeq,
                        NotPayYMS = model.NotPayYMS,
                        ChkCode = model.ChkCode,
                        RptIncludeFlag = model.RptIncludeFlag,
                        Memo = model.Memo,
                        ProcessUserCode=model.ProcessUserCode
                    });
                }
                else if (sA1 != "X1" && sA2 != "X2" && sA3 == "X3") //修改 2 VVX
                {
                    result = DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql, new
                    {
                        Iden = model.Iden,
                        CheckInd = model.CheckInd,
                        PayType = model.PayType,
                        PayMonth = model.PayMonth,
                        PaySeq = model.PaySeq,
                        NotPayYMS = model.NotPayYMS,
                        ChkCode = model.ChkCode,
                        RptIncludeFlag = model.RptIncludeFlag,
                        ProcessUserCode = model.ProcessUserCode
                    });
                }
                else if (sA1 != "X1" && sA2 == "X2" && sA3 != "X3") //修改 3 VXV
                {
                    result = DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql, new
                    {
                        Iden = model.Iden,
                        CheckInd = model.CheckInd,
                        PayType = model.PayType,
                        PayMonth = model.PayMonth,
                        PaySeq = model.PaySeq,
                        ChkCode = model.ChkCode,
                        NotPayYMS = model.NotPayYMS,
                        Memo = model.Memo,
                        ProcessUserCode = model.ProcessUserCode
                    });
                }
                else if (sA1 == "X1" && sA2 != "X2" && sA3 != "X3") //修改 4 XVV
                {
                    result = DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql, new
                    {
                        Iden = model.Iden,
                        RptIncludeFlag = model.RptIncludeFlag,
                        Memo = model.Memo,
                        ProcessUserCode = model.ProcessUserCode
                    });
                }
                else if (sA1 == "X1" && sA2 == "X2" && sA3 != "X3") //修改 5 XXV
                {
                    result = DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql, new
                    {
                        Iden = model.Iden,
                        Memo = model.Memo,
                        ProcessUserCode = model.ProcessUserCode
                    });
                }
                else if (sA1 == "X1" && sA2 != "X2" && sA3 == "X3") //修改 6 XVX
                {
                    result = DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql, new
                    {
                        Iden = model.Iden,
                        RptIncludeFlag = model.RptIncludeFlag,
                        ProcessUserCode = model.ProcessUserCode
                    });
                }
                else if (sA1 != "X1" && sA2 == "X2" && sA3 == "X3") //修改 7 VXX
                {
                    result = DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql, new
                    {
                        Iden = model.Iden,
                        CheckInd = model.CheckInd,
                        PayType = model.PayType,
                        PayMonth = model.PayMonth,
                        PaySeq = model.PaySeq,
                        NotPayYMS = model.NotPayYMS,
                        ChkCode = model.ChkCode,
                        ProcessUserCode = model.ProcessUserCode
                    });
                }

                //換成Trigger
                //if (result > 0)
                //{
                //    string sql1 = @"Insert Into MerSalCheckLog 
                //    Select Iden,run_seq,IDD,production_ym,sequence,company_code,file_seq,policy_no,policy_no2
                //    ,receipt_date,check_ind,pay_type,memo,pay_month,pay_seq,chk_code,not_pay_YMS,rpt_include_flag
                //    ,GETDATE(),@create_user_code,'Updated' 
                //    From MerSalCheck WITH (NOLOCK)
                //    Where Iden=@Iden";
                //    DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql1, new
                //    {
                //        create_user_code = model.CreateUserCode,
                //        Iden = model.Iden
                //    });
                //}
                return (result > 0);
            }
            catch (Exception ex)
            {
                Throw.BusinessError(ex.Message);
                Throw.BusinessError(ex.StackTrace);
                return (result < 0);
            }
        }

        #region 佣酬發放調整報表
        /// <summary>
        /// 佣酬發放調整報表
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public byte[] GetMerSalCheckReportList(MerSalCheckViewModel condition,string ShowTitle)
        {
            //取出資料
            //需求單號：20240715004 新增佣酬發放調整報表欄位，生效日、年齡、招攬人姓名及ID等欄位 vicky

            List<MerSalCheckViewModel> MerSalReport = GetMerSalCheck(condition) ?? new List<MerSalCheckViewModel>();
            List<MerSalCheckReportModel> MerSalCheckReports = new List<MerSalCheckReportModel>();
            if (MerSalReport.Count != 0)
            {
                foreach (var item in MerSalReport)
                {
                    MerSalCheckReportModel model = new MerSalCheckReportModel();
                    model.ProductionYM = item.ProductionYM;
                    model.Sequence = item.Sequence;
                    model.CompanyCode = item.CompanyCode;
                    model.FileSeq = item.FileSeq;
                    model.AmountType = item.AmountType;
                    model.PolicyNo2 = item.PolicyNo2;
                    model.PlanCode = item.PlanCode;
                    model.CollectYear = item.CollectYear;
                    model.ModxSequence = item.ModxSequence;
                    model.PoIssueDate = item.PoIssueDate;
                    model.Age = item.Age;
                    model.ReasonCodeC = item.ReasonCodeC;
                    model.ModePrem = item.ModePrem;
                    model.CommPremC = item.CommPremC;
                    model.Amount = item.Amount;
                    model.CommPrem = item.CommPrem;
                    model.AgentName1 = item.AgentName1;
                    model.AgentCode1 = item.AgentCode1;
                    model.AgentName2 = item.AgentName2;
                    model.AgentCode2 = item.AgentCode2;
                    model.ReceiptDate = item.ReceiptDate;
                    model.PayYm = item.PayYm;
                    model.PayTypeName = item.PayTypeName;
                    model.RptInclideFlag = item.RptIncludeFlag;
                    model.AgCnt = item.AgCnt;
                    model.PolicyStatusName = item.PolicyStatusName;
                    model.NotPayYMS = item.NotPayYMS;
                    model.Memo = item.Memo;
                    model.Remark = item.Remark;

                    MerSalCheckReports.Add(model);
                }

                //匯出excel格式
                MemoryStream ms = null;
                string HeadName = "";
                if (ShowTitle == "請選擇")
                {
                    HeadName = "佣酬發放調整報表";
                }
                else
                {
                    HeadName = "佣酬發放調整報表_" + ShowTitle;
                }
                List<EB.Common.CustHeader> headList = new List<EB.Common.CustHeader>();
                CustHeader c1 = new CustHeader();
                c1.CellsNum = new List<int>() { 29 };
                c1.CellsObj = new ArrayList() { HeadName };
                CustHeader c2 = new CustHeader();
                c2.CellsNum = new List<int> { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
                c2.CellsObj = new ArrayList() { "轉入工作月", "轉入序號", "保險公司", "檔案序號", "佣酬類別", "保單號碼", "保公險種", "年期"
                , "繳別繳次","生效日","年齡"
                , "保公原因碼", "保費", "保公佣金", "永達佣金", "永達FYC"
                , "招攬人1姓名", "招攬人1ID", "招攬人2姓名", "招攬人2ID"
                , "簽收回條日", "發放工作月", "發放註記","結轉下期","half件","保單狀態","異動工作月","人工備註","系統備註"};
                headList.AddRange(new CustHeader[] { c1, c2 });

                ms = NPoiHelper.Export(MerSalCheckReports, headList, "佣酬發放調整報表");
                return ms.ToArray();
            }
            else
            {
                return null;
            }
        }
        #endregion

        #endregion

        #region 檢核特殊資料設定
        /// <summary>
        /// 目前業績未關檔年月-次薪
        /// </summary>
        /// <returns></returns>
        public string GetNotYetAgbcInd()
        {
            string sql = @"select production_ym+'-'+convert(varchar,sequence) from agym
            where production_ym in (
            select MAX(production_ym) production_ym from agym where agbc_ind='0' and sequence<>88)
            and sequence in (
            select MAX(sequence) from agym 
            where agbc_ind='0' 
            and production_ym in (
            select MAX(production_ym) production_ym from agym where agbc_ind='0' and sequence<>88
            )
            and sequence<>88)";
            string result = DbHelper.Query<string>(VLifeRepository.ConnectionStringName, sql).FirstOrDefault();
            return result ?? "";
        }

        /// <summary>
        /// 最大業績已關檔年月
        /// </summary>
        /// <returns></returns>
        public string GetYmClose()
        {
            string sql = @"select MAX(production_ym) production_ym 
			from agym 
			where agbc_ind='1' and sequence<>88
			";
            string result = DbHelper.Query<string>(VLifeRepository.ConnectionStringName, sql).FirstOrDefault();
            return result ?? "";
        }

        /// <summary>
        /// 最大業績已關檔序號
        /// </summary>
        /// <returns></returns>
        public string GetSeqClose()
        {
            string sql = @"select MAX(sequence) sequence
			from agym 
			where agbc_ind='1' 
			and production_ym in (
			select MAX(production_ym) production_ym from agym where agbc_ind='1' and sequence<>88)
			";
            int result = DbHelper.Query<Int16>(VLifeRepository.ConnectionStringName, sql).FirstOrDefault();
            return result.ToString() ?? "";
        }

        /// <summary>
        /// 取得保公
        /// </summary>
        /// <returns></returns>
        public List<TermVal> GetALLCompanyCode()
        {
            string sql = @"select term_code,term_meaning,ROW_NUMBER()over(order by term_code) row_seq
            from trmval where term_id='company_code' and term_code LIKE '%[0-9]%' 
            union
            select '' term_code,'全部保險公司' term_meaning ,'99' row_seq
            order by row_seq";
            List<TermVal> result = DbHelper.Query<TermVal>(VLifeRepository.ConnectionStringName, sql).ToList();

            return result;
        }

        /// <summary>
        /// 檢核特殊資料設定查詢
        /// </summary>
        /// <returns></returns>
        public List<MerSalCheckSPruleViewModel> GetMerSalCheckSPrule(MerSalCheckSPruleViewModel model, string ClickItem)
        {
            #region SQL


            string sql = "";
            sql = @"";
            sql += " select ";
            if (ClickItem == "Query")
            {
                sql += " iden, ";
            }
            sql += "sp.production_ym_s+'-'+sp.sequence_s production_ym_s";
            sql += ",sp.production_ym_e+'-'+sp.sequence_e production_ym_e";
            sql += ",sp.company_code";
            sql += ",sp.file_seq";
            sql += ",sp.amount_type+'-'+rtrim(tv1.term_meaning) amount_type_name";
            sql += ",rtrim(tv2.term_meaning) chk_code_name";
            sql += ",sp.act_name";
            sql += ",sp.rule_01,sp.rule_02,sp.rule_03,sp.rule_04";
            sql += ",sp.remark";
            sql += ",sp.create_datetime,sp.create_user_code,udc.UserName create_user_name";
            sql += ",sp.update_datetime,sp.update_user_code,udu.UserName update_user_name";
            sql += ",sp.is_delete";
            sql += " from MerSalCheckSPrule sp";
            sql += " left join trmval tv1 on tv1.term_id='amount_type' and tv1.term_code=sp.amount_type";
            sql += " left join trmval tv2 on tv2.term_id='chkcode_acttype' and tv2.term_code=sp.chk_code+'|'+sp.act_type";
            sql += " left join UserData udc on udc.UserId=isnull(sp.create_user_code,'')";
            sql += " left join UserData udu on udu.UserId=isnull(sp.update_user_code,'')";
            sql += " where 1=1 ";

            if (!String.IsNullOrEmpty(model.ProductionYmS))
            {
                sql += " and production_ym_s+sequence_s>=replace(@ProductionYmS,'-','')";
            }
            if (!String.IsNullOrEmpty(model.ProductionYmE))
            {
                sql += " and production_ym_e+sequence_e<= replace(@ProductionYmE,'-','')";
            }
            if (!String.IsNullOrEmpty(model.CompanyCode))
            {
                sql += " and company_code=@CompanyCode";
            }
            if (!String.IsNullOrEmpty(model.FileSeq))
            {
                sql += "  and file_seq=@FileSeq";
            }
            if (!String.IsNullOrEmpty(model.ChkCodeActType))
            {
                sql += "  and chk_code+'|'+act_type=@ChkCodeActType";
            }
            if (!String.IsNullOrEmpty(model.AmountType))
            {
                sql += " and amount_type=@AmountType";
            }
            if (!String.IsNullOrEmpty(model.Rule01))
            {
                sql += " and rule_01 like @Rule01";
            }
            if (!String.IsNullOrEmpty(model.Rule02))
            {
                sql += " and rule_02  like @Rule02";
            }
            if (!String.IsNullOrEmpty(model.Rule03))
            {
                sql += " and rule_03  like @Rule03";
            }
            if (!String.IsNullOrEmpty(model.Rule04))
            {
                sql += " and rule_04  like @Rule04";
            }
            if (!String.IsNullOrEmpty(model.IsDelete))
            {
                sql += " and is_delete = @IsDelete";
            }
            sql += " order by sp.company_code,sp.file_seq,sp.amount_type,sp.chk_code,sp.act_type,sp.rule_01,sp.rule_02,sp.rule_03,sp.production_ym_e";
            #endregion

            List<MerSalCheckSPruleViewModel> result = DbHelper.Query<MerSalCheckSPruleViewModel>(WebBrokerRepository.ConnectionStringName, sql
                , new
                {
                    ProductionYmS = model.ProductionYmS,
                    ProductionYmE = model.ProductionYmE,
                    CompanyCode = model.CompanyCode,
                    FileSeq = model.FileSeq,
                    ChkCodeActType = model.ChkCodeActType,
                    AmountType = model.AmountType,
                    IsDelete = model.IsDelete,
                    Rule01 = "%" + model.Rule01 + "%",
                    Rule02 = "%" + model.Rule02 + "%",
                    Rule03 = "%" + model.Rule03 + "%",
                    Rule04 = "%" + model.Rule04 + "%",
                }).ToList();
            return result;
        }

        #region 檢核特殊資料設定報表
        /// <summary>
        /// 檢核特殊資料設定報表
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public byte[] GetMerSalCheckSPruleReportList(MerSalCheckSPruleViewModel condition)
        {
            //取出資料
            List<MerSalCheckSPruleViewModel> MerSalSPruleReport = GetMerSalCheckSPrule(condition, "Report") ?? new List<MerSalCheckSPruleViewModel>();
            List<MerSalCheckSPruleReportModel> MerSalCheckSPruleReports = new List<MerSalCheckSPruleReportModel>();
            if (MerSalSPruleReport.Count != 0)
            {
                foreach (var item in MerSalSPruleReport)
                {
                    MerSalCheckSPruleReportModel model = new MerSalCheckSPruleReportModel();
                    model.ProductionYmS = item.ProductionYmS;
                    model.ProductionYmE = item.ProductionYmE;
                    model.CompanyCode = item.CompanyCode;
                    model.FileSeq = item.FileSeq;
                    model.AmountTypeName = item.AmountTypeName;
                    model.ChkCodeName = item.ChkCodeName;
                    model.ActName = item.ActName;
                    model.Rule01 = item.Rule01;
                    model.Rule02 = item.Rule02;
                    model.Rule03 = item.Rule03;
                    model.Rule04 = item.Rule04;
                    model.Remark = item.Remark;
                    model.CreateDatetime = item.CreateDatetime;
                    model.CreateUserName = item.CreateUserName;
                    model.UpdateDatetime = item.UpdateDatetime;
                    model.UpdateUserName = item.UpdateUserName;
                    model.IsDelete = item.IsDelete;
                    MerSalCheckSPruleReports.Add(model);
                }

                //匯出excel格式
                MemoryStream ms = null;
                string HeadName = "檢核特殊資料設定報表";
                List<EB.Common.CustHeader> headList = new List<EB.Common.CustHeader>();
                CustHeader c1 = new CustHeader();
                c1.CellsNum = new List<int>() { 17 };
                c1.CellsObj = new ArrayList() { HeadName };

                CustHeader c2 = new CustHeader();
                c2.CellsNum = new List<int> { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
                c2.CellsObj = new ArrayList() { "工作月序號起", "工作月序號迄", "保險公司", "檔案序號", "佣酬類別", "特殊檢核設定類別", "特殊檢核顯示中文"
                , "規則01", "規則02", "規則03", "規則04","備註","建檔時間","建檔人員","更新時間","更新人員","狀態"};
                headList.AddRange(new CustHeader[] { c1, c2 });

                ms = NPoiHelper.Export(MerSalCheckSPruleReports, headList, "檢核特殊資料設定報表");
                return ms.ToArray();
            }
            else
            {
                return null;
            }
        }
        #endregion

        /// <summary>
        /// 更新停用狀態
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public bool UpdateMerSalCheckSPruleIsDelete(MerSalCheckSPruleViewModel model)
        {
            string sql = @"UPDATE MerSalCheckSPrule SET production_ym_e=@YmClose,sequence_e=@SeqClose,is_delete='Y',update_user_code = @UpdateUserCode ,update_datetime = getdate()
            FROM MerSalCheckSPrule
            WHERE iden=@Iden";
            int result = -1;
            try
            {
                result = DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql, new
                {
                    UpdateUserCode = model.UpdateUserCode,
                    YmClose = model.YmClose,
                    SeqClose = model.SeqClose,
                    Iden = model.iden
                });
                return (result > 0);
            }
            catch (Exception ex)
            {
                Throw.BusinessError(ex.Message);
            }

            return (result < 0);
        }

        /// <summary>
        /// 新增MerSalCheckSPrule
        /// </summary>
        /// <param name="model">data</param>
        public string InsertMerSalCheckSPrule(MerSalCheckSPrule model)
        {
            //1.檢查有無重複的資料
            string sql = @"select count(*) cnt from MerSalCheckSPrule 
                    WHERE @ProductionYmS between production_ym_s and production_ym_e 
                    AND company_code=@CompanyCode AND file_seq=@FileSeq
                    AND amount_type=@AmountType
                    AND chk_code=@ChkCode
                    AND act_type=@ActType
                    AND rule_01=@Rule01
                    AND rule_02=@Rule02
                    AND rule_03=@Rule03
                    AND rule_04=@Rule04
                    AND is_delete='N'";
            int duplicateCount = DbHelper.Query<int>(WebBrokerRepository.ConnectionStringName, sql, new
            {
                ProductionYmS = model.ProductionYmS,
                CompanyCode = model.CompanyCode,
                FileSeq = model.FileSeq,
                AmountType = model.AmountType,
                ChkCode = model.ChkCode,
                ActType = model.ActType,
                Rule01 = model.Rule01,
                Rule02 = model.Rule02,
                Rule03 = model.Rule03,
                Rule04 = model.Rule04
            }).FirstOrDefault();

            if (duplicateCount > 0)
            {
                return "Duplicate";
            }
            else
            {
                //2.新增
                string sql2 = @"insert into MerSalCheckSPrule values
                    (@ProductionYmS,@SequenceS,@ProductionYmE,@SequenceE,@CompanyCode,@FileSeq,@AmountType,@ChkCode,@ActType,@ActName
                    ,ISNULL(@Rule01,''),ISNULL(@Rule02,''),ISNULL(@Rule03,''),ISNULL(@Rule04,''),@Remark,getdate(),@CreateUserCode,NULL,NULL,'N');";
                int result = DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql2, model);

                return (result > 0) ? "OK" : "Fail";
            }


            //try
            //{
            //    string sql = " insert into MerSalCheckSPrule values" +
            //        "(@ProductionYmS,@SequenceS,@ProductionYmE,@SequenceE,@CompanyCode,@FileSeq,@AmountType,@ChkCode,@ActType,@ActName,@Rule01,@Rule02,@Rule03,@Remark,getdate(),@CreateUserCode,NULL,NULL,'N')";
            //    DbHelper.Execute(WebBrokerRepository.ConnectionStringName, sql, model);
            //}
            //catch (Exception ex)
            //{
            //    throw new Exception(ex.ToString());
            //}
        }
        #endregion

        #region 原始檔系統保留暨人工調帳查詢

        #region 原始檔系統保留暨人工調帳-查詢
        /// <summary>
        /// 原始檔系統保留暨人工調帳查詢
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public List<MerSalCutViewModel> GetMerSalCut(MerSalCutViewModel model)
        {
            string sql = "";
            sql = @"";
            sql += " select ";
            if (model.BtnType == "btnReport")//查詢結果下載
            {
                sql += " a.IDD,a.production_ym,a.sequence ";
            }
            else
            {
                sql += " a.production_ym,a.sequence";
            }

            if (model.BtnType == "btnRptPayRoll")//大批人工調帳格式下載
            {
                sql += " ,cast('' as char(1)) EmpCol1 ";//型態
            }

            sql += " ,a.company_code";
            sql += " ,agent_code1,nn1.names names1,agent_code2,ISNULL(nn2.names,'') names2";

            //if (model.BtnType == "btnRptPayRoll")
            //{
            //sql += " ,cast('' as char(1)) EmpCol2";//調整原因碼
            sql += " ,reason_code";//調整原因碼
            //}

            sql += " ,policy_no2,a.plan_code,a.collect_year,a.modx_sequence";

            if (model.BtnType == "btnRptPayRoll")//大批人工調帳格式下載
            {
                sql += " , cast('' as char(1)) EmpCol3, cast('' as char(1)) EmpCol4,cast('' as char(1)) EmpCol5";//空白格-金額;FYC;FYP  20230901刪除-計佣保費, cast('' as char(1)) EmpCol6
                sql += " , case when a.amount_type in('01','02') then rtrim(a.insured_name)+' '+a.po_issue_date+' '+a.plan_code+' '+convert(varchar,convert(int,a.collect_year))+'年 $'+convert(varchar,a.mode_prem) else '' end EmpCol7";//原因說明
            }

            sql += " ,insured_name,a.age,a.po_issue_date";
            sql += " ,mode_prem,comm_prem_c";

            if (model.BtnType == "btnReport")//查詢結果下載
            {
                sql += " ,case when a.pay_month='' then '' else a.pay_month+'-'+a.pay_seq end 'pay_month_seq'";

            }
            else
            {
                sql += " ,a.IDD ";
            }
            sql += " ,rtrim(tv1.term_meaning) check_ind_name";
            sql += " from MerSalCut a";
            sql += " left join vlife.dbo.nain nn1 on left(agent_code1,10)=nn1.client_id";
            sql += " left join vlife.dbo.nain nn2 on left(agent_code2,10)=nn2.client_id";
            sql += " left join trmval tv1 on a.check_ind=tv1.term_code and tv1.term_id='check_ind'";

            sql += " where 1 = 1";
            if (!String.IsNullOrEmpty(model.ProductionYM))
            {
                sql += " and production_ym = @ProductionYm";
            }
            if (!String.IsNullOrEmpty(model.Sequence))
            {
                sql += " and sequence = @Sequence";
            }
            if (!String.IsNullOrEmpty(model.CompanyCode))
            {
                sql += " and a.company_code = @CompanyCode";
            }
            if (!String.IsNullOrEmpty(model.CheckInd))
            {
                sql += " and check_ind = @CheckInd";
            }

            List<MerSalCutViewModel> result = DbHelper.Query<MerSalCutViewModel>(WebBrokerRepository.ConnectionStringName, sql,
                new
                {
                    ProductionYm = model.ProductionYM,
                    Sequence = model.Sequence,
                    CompanyCode = model.CompanyCode,
                    CheckInd = model.CheckInd,
                    BtnType = model.BtnType
                }).ToList();

            return result;
        }
        #endregion

        #region 原始檔系統保留暨人工調帳-報表
        /// <summary>
        /// 原始檔系統保留暨人工調帳報表
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        public byte[] GetMerSalCutReportList(MerSalCutViewModel condition, string FileName)
        {
            //取出資料

            List<MerSalCutViewModel> MerSalReport = GetMerSalCut(condition) ?? new List<MerSalCutViewModel>();


            if (MerSalReport.Count != 0)
            {
                if (condition.BtnType == "btnReport")
                {
                    List<MerSalCutReportModel> MerSalCheckReports1 = new List<MerSalCutReportModel>();
                    foreach (var item in MerSalReport)
                    {
                        MerSalCutReportModel model1 = new MerSalCutReportModel();
                        model1.IDD = item.IDD;
                        model1.ProductionYM = item.ProductionYM;
                        model1.Sequence = item.Sequence;
                        model1.CompanyCode = item.CompanyCode;
                        model1.AgentCode1 = item.AgentCode1;
                        model1.Names1 = item.Names1;
                        model1.AgentCode2 = item.AgentCode2;
                        model1.Names2 = item.Names2;
                        model1.ReasonCode = item.ReasonCode;
                        model1.PolicyNo2 = item.PolicyNo2;
                        model1.PlanCode = item.PlanCode;
                        model1.CollectYear = item.CollectYear;
                        model1.ModxSequence = item.ModxSequence;
                        model1.InsuredName = item.InsuredName;
                        model1.Age = item.Age;
                        model1.PoIssueDate = item.PoIssueDate;
                        model1.ModePrem = item.ModePrem;
                        model1.CommPremC = item.CommPremC;
                        model1.PayMonth = item.PayMonth;
                        model1.PaySeq = item.PaySeq;
                        model1.CheckIndName = item.CheckIndName;

                        MerSalCheckReports1.Add(model1);
                    }
                    MemoryStream ms = null;
                    string HeadName = FileName.Replace(".xlsx", "");//檔名
                    List<EB.Common.CustHeader> headList = new List<EB.Common.CustHeader>();
                    CustHeader c1 = new CustHeader();//1. 表頭
                    CustHeader c2 = new CustHeader();//2. 報表內容

                    //查詢結果下載
                    //1. 表頭
                    c1.CellsNum = new List<int>() { 21 };//合併儲存格
                    c1.CellsObj = new ArrayList() { HeadName };

                    //2. 報表內容
                    c2.CellsNum = new List<int> { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,1 };
                    c2.CellsObj = new ArrayList() { "資料序號","轉入工作月", "轉入序號", "公司碼"
                                                , "業務員(一)代碼", "業務員(一)姓名", "業務員(二)代碼", "業務員(二)姓名"
                                                , "調整原因碼"
                                                , "保單號碼", "險種代碼", "繳費年期", "繳別繳次"
                                                , "被保險人", "年齡", "生效日"
                                                , "保公保費", "保公佣金"
                                                , "大批調整年月", "大批調整序號","類型" };
                    headList.AddRange(new CustHeader[] { c1, c2 });
                    ms = NPoiHelper.Export(MerSalCheckReports1, headList, HeadName); //報表產生
                    return ms.ToArray();
                }
                else
                {
                    List<MerSalCutRptPayRollModel> MerSalCheckReports2 = new List<MerSalCutRptPayRollModel>();
                    foreach (var item in MerSalReport)
                    {
                        MerSalCutRptPayRollModel model2 = new MerSalCutRptPayRollModel();
                        model2.ProductionYM = item.ProductionYM;
                        model2.Sequence = item.Sequence;
                        model2.EmpCol1 = item.EmpCol1; //型態
                        model2.CompanyCode = item.CompanyCode;
                        model2.AgentCode1 = item.AgentCode1;
                        model2.Names1 = item.Names1;
                        model2.AgentCode2 = item.AgentCode2;
                        model2.Names2 = item.Names2;
                        model2.ReasonCode = item.ReasonCode;//調整原因碼
                        model2.PolicyNo2 = item.PolicyNo2;
                        model2.PlanCode = item.PlanCode;
                        model2.CollectYear = item.CollectYear;
                        model2.ModxSequence = item.ModxSequence;
                        model2.EmpCol3 = item.EmpCol3;//金額
                        model2.EmpCol4 = item.EmpCol4;//FYC
                        model2.EmpCol5 = item.EmpCol5;//FYP
                        //model2.EmpCol6 = item.EmpCol6;//計佣保費
                        model2.EmpCol7 = item.EmpCol7;//原因說明
                        model2.InsuredName = item.InsuredName;
                        model2.Age = item.Age;
                        model2.PoIssueDate = item.PoIssueDate;
                        model2.ModePrem = item.ModePrem;
                        model2.CommPremC = item.CommPremC;
                        model2.IDD = item.IDD;
                        model2.CheckIndName = item.CheckIndName;

                        MerSalCheckReports2.Add(model2);
                    }
                    MemoryStream ms = null;
                    string HeadName = FileName.Replace(".xlsx", "");//檔名
                    List<EB.Common.CustHeader> headList = new List<EB.Common.CustHeader>();
                    CustHeader c1 = new CustHeader();//1. 表頭
                    CustHeader c2 = new CustHeader();//2. 報表內容

                    //大批人工調帳轉格式下載
                    c1.CellsNum = new List<int>() { 24 };
                    c1.CellsObj = new ArrayList() { HeadName };
                    c2.CellsNum = new List<int> { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,1 };
                    c2.CellsObj = new ArrayList() { "轉入工作月", "轉入序號","型態", "公司碼"
                                                , "業務員(一)代碼", "業務員(一)姓名", "業務員(二)代碼", "業務員(二)姓名"
                                                , "調整原因碼"
                                                , "保單號碼", "險種代碼", "繳費年期", "繳別繳次"
                                                , "金額", "FYC", "FYP", "原因說明"
                                                , "被保險人", "年齡", "生效日"
                                                , "保公保費", "保公佣金"
                                                , "資料序號","類型" };

                    headList.AddRange(new CustHeader[] { c1, c2 });
                    ms = NPoiHelper.Export(MerSalCheckReports2, headList, HeadName);//報表產生
                    return ms.ToArray();
                }


            }
            else
            {
                return null;
            }
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

        #endregion

        #endregion

    }
}

