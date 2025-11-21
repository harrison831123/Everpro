using EP.CUFModels;
using EP.H2OModels;
using EP.VLifeModels;
using EP.Platform.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EP.SD.SalesSupport.LAW.Models;
using System.IO;
using EP.Common;
using System.Collections;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace EP.SD.SalesSupport.LAW.Service
{
    public class LAWService : ILAWService
    {
        #region 系統設定作業
        #region 管理人員設定
        /// <summary>
        /// 管理人員設定列表
        /// </summary>
        public List<LawSys> GetLawSys()
        {
            List<LawSys> result = new List<LawSys>();
            string sql = @"select * from law_sys";

            result = DbHelper.Query<LawSys>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 新增管理人員
        /// </summary>
        public void InsertLawSys(List<LawSys> model)
        {
            string sql = @"INSERT INTO law_sys(sys_name,create_name,create_date,SysMemberID,SysCreatorID)
                           VALUES (@SysName, @Createname, @CreateDate, @SysMemberID, @SysCreatorID);";

            for (int i = 0; i < model.Count; i++)
            {
                DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
                {
                    CreateDate = model[i].CreateDate,
                    Createname = model[i].Createname,
                    SysCreatorID = model[i].SysCreatorID,
                    SysMemberID = model[i].SysMemberID,
                    SysName = model[i].SysName,
                });
            }
        }

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawSysByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawSys>(new { ID = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }

        /// <summary>
        /// 取得部門level
        /// </summary>
        /// <param name="unitid">部門編號</param>
        public int GetUnitlevel(string unitid)
        {
            string sqlGetData = @"select * from sc_unit a join sc_unit_ext b on a.uunit = b.uunit
                                where iunit = @iunit
                                order by organization_level";
            sc_unit_ext result = DbHelper.Query<sc_unit_ext>(CUFRepository.ConnectionStringName, sqlGetData, new
            {
                iunit = unitid
            }).FirstOrDefault();

            return Convert.ToInt32(result.sign_level);
        }
        #endregion

        #region 經辦人員設定
        /// <summary>
        /// 經案人員設定列表
        /// </summary>
        public List<LawDoUser> GetLawDouser()
        {
            List<LawDoUser> result = new List<LawDoUser>();
            string sql = @"select * from law_douser order by douser_sort asc";

            result = DbHelper.Query<LawDoUser>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 新增經辦人員
        /// </summary>
        public void InsertLawDouser(List<LawDoUser> model)
        {
            string sql = @"INSERT INTO law_douser(douser_name,douser_phone_ext,create_name,create_date,douser_sort,DouserMemberID,DouserCreatorID)
                                          VALUES (@DouserName, @DouserPhoneExt, @CreateName, @CreateDate, @DouserSort,@DouserMemberID,@DouserCreatorID);";

            for (int i = 0; i < model.Count; i++)
            {
                DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
                {
                    DouserName = model[i].DouserName,
                    DouserPhoneExt = model[i].DouserPhoneExt,
                    CreateDate = model[i].CreateDate,
                    Createname = model[i].CreateName,
                    DouserSort = model[i].DouserSort,
                    DouserCreatorID = model[i].DouserCreatorID,
                    DouserMemberID = model[i].DouserMemberID,
                });
            }
        }

        /// <summary>
        /// 經辦人員詳細資料
        /// </summary>
        public List<LawDoUser> GetLawDouserByID(string ID)
        {
            List<LawDoUser> result = new List<LawDoUser>();
            string sql = @"select * from law_douser  where law_douser_id = @LawDouserId ";

            result = DbHelper.Query<LawDoUser>(H2ORepository.ConnectionStringName, sql, new { LawDouserId = ID }).ToList();
            return result;
        }

        /// <summary>
        /// 更新經辦人員
        /// </summary>
        public void UpdateLawDouser(LawDoUser model)
        {
            var sql = @"update law_douser 
            set douser_phone_ext = @douser_phone_ext,douser_sort = @douser_sort
            where law_douser_id = @LawDouserId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { douser_phone_ext = model.DouserPhoneExt, douser_sort = model.DouserSort, LawDouserId = model.LawDouserId });
        }

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawDouserByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawDoUser>(new { LawDouserId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }

        /// <summary>
        /// 確認經辦順位
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public string CheckDouserSort(string DouserSort)
        {
            int result;
            string sql = @"select douser_sort from law_douser  where douser_sort = @DouserSort ";
            result = DbHelper.Query<int>(H2ORepository.ConnectionStringName, sql, new { DouserSort = DouserSort }).FirstOrDefault();
            return result.ToString();
        }

        /// <summary>
        /// 用ID確認經辦順位
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public string CheckDouserSortByID(string DouserSort, string LawDouserId)
        {
            int result;
            string sql = @"select douser_sort from law_douser  where douser_sort = @DouserSort and law_douser_id = @LawDouserId ";
            result = DbHelper.Query<int>(H2ORepository.ConnectionStringName, sql, new { DouserSort = DouserSort, LawDouserId = LawDouserId }).FirstOrDefault();
            return result.ToString();
        }

        #endregion

        #region 承辦單位設定
        /// <summary>
        /// 承辦單位設定列表
        /// </summary>
        public List<LawDoUnit> GetLawDoUnit()
        {
            List<LawDoUnit> result = new List<LawDoUnit>();
            string sql = @"select * from law_do_unit";

            result = DbHelper.Query<LawDoUnit>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 新增承辦單位
        /// </summary>
        public void InsertLawDoUnit(LawDoUnit model)
        {
            string sql = @"INSERT INTO law_do_unit(unit_name,create_name,create_date,status_type,DounitCreatorID)
            VALUES (@UnitName, @CreateName, @CreateDate, @StatusType,@DounitCreatorID);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                UnitName = model.UnitName,
                CreateName = model.CreateName,
                CreateDate = model.CreateDate,
                StatusType = model.StatusType,
                DounitCreatorID = model.DounitCreatorID,
            });
        }

        /// <summary>
        /// 承辦單位詳細資料
        /// </summary>
        public List<LawDoUnit> GetLawDoUnitByID(string ID)
        {
            List<LawDoUnit> result = new List<LawDoUnit>();
            string sql = @"select * from law_do_unit  where law_do_unit_id = @LawDoUnitId ";

            result = DbHelper.Query<LawDoUnit>(H2ORepository.ConnectionStringName, sql, new { LawDoUnitId = ID }).ToList();
            return result;
        }

        /// <summary>
        /// 更新承辦單位
        /// </summary>
        public void UpdateLawDoUnit(LawDoUnit model)
        {
            var sql = @"update law_do_unit 
            set unit_name = @UnitName
            where law_do_unit_id = @LawDoUnitId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { UnitName = model.UnitName, LawDoUnitId = model.LawDoUnitId });
        }

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawDoUnitByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawDoUnit>(new { LawDoUnitId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }

        /// <summary>
        /// 修改承辦狀態
        /// </summary>
        public void ChangeStatusType(string StatusType, string ID)
        {
            var sql = @"update law_do_unit 
            set status_type = @StatusType
            where law_do_unit_id = @LawDoUnitId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { StatusType = StatusType, LawDoUnitId = ID });
        }
        #endregion

        #region 相關費率設定
        /// <summary>
        /// 相關費率設定列表
        /// </summary>
        public List<LawInterestRates> GetLawInterestRates()
        {
            List<LawInterestRates> result = new List<LawInterestRates>();
            string sql = @"select * from law_interest_rates";

            result = DbHelper.Query<LawInterestRates>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 新增相關費率
        /// </summary>
        public void InsertLawInterestRates(LawInterestRates model)
        {
            string sql = @"INSERT INTO law_interest_rates(interest_rates,create_name,create_date,interest_rates_type,InterestRatesCreateID)
            VALUES (@InterestRates, @CreateName, @CreateDate, @InterestRatesType,@InterestRatesCreateID);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                InterestRates = model.InterestRates,
                CreateName = model.CreateName,
                CreateDate = model.CreateDate,
                InterestRatesType = model.InterestRatesType,
                InterestRatesCreateID = model.InterestRatesCreateID,
            });
        }

        /// <summary>
        /// 相關費率詳細資料
        /// </summary>
        public List<LawInterestRates> GetLawInterestRatesByID(string ID)
        {
            List<LawInterestRates> result = new List<LawInterestRates>();
            string sql = @"select * from law_interest_rates  where law_interest_rates_id = @LawInterestRatesId ";

            result = DbHelper.Query<LawInterestRates>(H2ORepository.ConnectionStringName, sql, new { LawInterestRatesId = ID }).ToList();
            return result;
        }

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawInterestRatesByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawInterestRates>(new { LawInterestRatesId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }

        /// <summary>
        /// 修改相關費率狀態
        /// </summary>
        public void ChangeInterestRatesType(string InterestRatesType, string ID)
        {
            var sql = @"update law_interest_rates 
            set interest_rates_type = @InterestRatesType
            where law_interest_rates_id = @LawInterestRatesId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { InterestRatesType = InterestRatesType, LawInterestRatesId = ID });
        }

        /// <summary>
        /// 律師費率設定列表
        /// </summary>
        public List<LawLawyerServiceRates> GetLawLawyerServiceRates()
        {
            List<LawLawyerServiceRates> result = new List<LawLawyerServiceRates>();
            string sql = @"select * from law_lawyer_service_rates";

            result = DbHelper.Query<LawLawyerServiceRates>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 新增律師費率
        /// </summary>
        public void InsertLawLawyerServiceRates(LawLawyerServiceRates model)
        {
            string sql = @"INSERT INTO law_lawyer_service_rates(lawyer_service_rates,create_name,create_date,lawyer_service_rates_type,LawyerServiceRatesCreatorID)
            VALUES (@LawyerServiceRates, @CreateName, @CreateDate, @LawyerServiceRatesType,@LawyerServiceRatesCreatorID);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawyerServiceRates = model.LawyerServiceRates,
                CreateName = model.CreateName,
                CreateDate = model.CreateDate,
                LawyerServiceRatesType = model.LawyerServiceRatesType,
                LawyerServiceRatesCreatorID = model.LawyerServiceRatesCreatorID,
            });
        }

        /// <summary>
        /// 律師費率詳細資料
        /// </summary>
        public List<LawLawyerServiceRates> GetLawLawyerServiceRatesByID(string ID)
        {
            List<LawLawyerServiceRates> result = new List<LawLawyerServiceRates>();
            string sql = @"select * from law_lawyer_service_rates  where law_lawyer_service_rates_id = @LawInterestRatesId ";

            result = DbHelper.Query<LawLawyerServiceRates>(H2ORepository.ConnectionStringName, sql, new { LawInterestRatesId = ID }).ToList();
            return result;
        }

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawLawyerServiceRatesByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawLawyerServiceRates>(new { LawLawyerServiceRatesId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }

        /// <summary>
        /// 修改律師費率狀態
        /// </summary>
        public void ChangeLawyerServiceRatesType(string LawyerServiceRatesType, string ID)
        {
            var sql = @"update law_lawyer_service_rates 
            set lawyer_service_rates_type = @LawyerServiceRatesType
            where law_lawyer_service_rates_id = @LawLawyerServiceRatesId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { LawyerServiceRatesType = LawyerServiceRatesType, LawLawyerServiceRatesId = ID });
        }
        #endregion

        #region 結案狀態設定
        /// <summary>
        /// 結案狀態設定列表
        /// </summary>
        public List<LawCloseType> GetLawCloseType()
        {
            List<LawCloseType> result = new List<LawCloseType>();
            string sql = @"select * from law_close_type";

            result = DbHelper.Query<LawCloseType>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 新增結案狀態
        /// </summary>
        public void InsertLawCloseType(LawCloseType model)
        {
            string sql = @"INSERT INTO law_close_type(close_type_name,create_name,create_date,status_type,count_type,CloseTypeCreatorID)
            VALUES (@CloseTypeName, @CreateName, @CreateDate, @StatusType,@CountType,@CloseTypeCreatorID);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                CloseTypeName = model.CloseTypeName,
                CreateName = model.CreateName,
                CreateDate = model.CreateDate,
                StatusType = model.StatusType,
                CountType = model.CountType,
                CloseTypeCreatorID = model.CloseTypeCreatorID,
            });
        }

        /// <summary>
        /// 結案狀態詳細資料
        /// </summary>
        public List<LawCloseType> GetLawCloseTypeByID(string ID)
        {
            List<LawCloseType> result = new List<LawCloseType>();
            string sql = @"select * from law_close_type where close_type_id = @CloseTypeId ";

            result = DbHelper.Query<LawCloseType>(H2ORepository.ConnectionStringName, sql, new { CloseTypeId = ID }).ToList();
            return result;
        }

        /// <summary>
        /// 更新結案狀態
        /// </summary>
        public void UpdateLawCloseType(LawCloseType model)
        {
            var sql = @"update law_close_type 
            set close_type_name = @CloseTypeName,count_type = @CountType
            where close_type_id = @CloseTypeId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { CloseTypeName = model.CloseTypeName, CountType = model.CountType, CloseTypeId = model.CloseTypeId });
        }

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawCloseTypeByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawCloseType>(new { CloseTypeId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }

        /// <summary>
        /// 修改結案狀態
        /// </summary>
        public void ChangeLawCloseTypeStatusType(string StatusType, string ID)
        {
            var sql = @"update law_close_type 
            set status_type = @StatusType
            where close_type_id = @CloseTypeId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { StatusType = StatusType, CloseTypeId = ID });
        }
        #endregion

        #region 契撤變原因設定
        /// <summary>
        /// 契撤變原因設定列表
        /// </summary>
        public List<LawEvidType> GetLawEvidType()
        {
            List<LawEvidType> result = new List<LawEvidType>();
            string sql = @"select * from law_evid_type";

            result = DbHelper.Query<LawEvidType>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 新增契撤變原因
        /// </summary>
        public void InsertLawEvidType(LawEvidType model)
        {
            string sql = @"INSERT INTO law_evid_type(evid_type_name,create_name,create_date,evid_status_type,EvidTypeCreatorID)
            VALUES (@EvidTypeName, @CreateName, @CreateDate, @EvidStatusType,@EvidTypeCreatorID);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                EvidTypeName = model.EvidTypeName,
                CreateName = model.CreateName,
                CreateDate = model.CreateDate,
                EvidStatusType = model.EvidStatusType,
                EvidTypeCreatorID = model.EvidTypeCreatorID,
            });
        }

        /// <summary>
        /// 契撤變原因詳細資料
        /// </summary>
        public List<LawEvidType> GetLawEvidTypeByID(string ID)
        {
            List<LawEvidType> result = new List<LawEvidType>();
            string sql = @"select * from law_evid_type where evid_type_id = @EvidTypeId ";

            result = DbHelper.Query<LawEvidType>(H2ORepository.ConnectionStringName, sql, new { EvidTypeId = ID }).ToList();
            return result;
        }

        /// <summary>
        /// 更新契撤變原因
        /// </summary>
        public void UpdateLawEvidType(LawEvidType model)
        {
            var sql = @"update law_evid_type 
            set evid_type_name = @EvidTypeName
            where evid_type_id = @EvidTypeId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { EvidTypeName = model.EvidTypeName, EvidTypeId = model.EvidTypeId });
        }

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawEvidTypeByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawEvidType>(new { EvidTypeId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }

        /// <summary>
        /// 修改契撤變原因狀態
        /// </summary>
        public void ChangeLawEvidTypeStatusType(string StatusType, string ID)
        {
            var sql = @"update law_evid_type 
            set evid_status_type = @EvidStatusType
            where evid_type_id = @EvidTypeId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { EvidStatusType = StatusType, EvidTypeId = ID });
        }

        /// <summary>
        /// 確認契撤變原因有無重複
        /// </summary>
        public bool chkEvidTypeNameRepeat(string EvidTypeName)
        {
            string resultsql = string.Empty;
            bool result = false;
            string sql = @"select * from law_evid_type where evid_type_id = @EvidTypeId ";

            resultsql = DbHelper.Query<string>(H2ORepository.ConnectionStringName, sql, new { EvidTypeName = EvidTypeName }).FirstOrDefault();

            if (!string.IsNullOrEmpty(resultsql))
            {
                return result;
            }
            else
            {
                result = true;
                return result;
            }
        }
        #endregion

        #region 報表排序設定
        /// <summary>
        /// 確認報表排序有沒有設定新年度的資料，沒有則新增
        /// </summary>
        public void CheckLawReportSortBySortYear(string SortYear)
        {
            List<LawReportSort> result = new List<LawReportSort>();
            string sql = @"select sort_year,sort_vm,sort_vm_name from law_report_sort where sort_year= @SortYear
                        group by sort_year,sort_vm,sort_vm_name order by sort_vm,sort_vm_name";

            result = DbHelper.Query<LawReportSort>(H2ORepository.ConnectionStringName, sql, new { SortYear = SortYear }).ToList();

            if (!result.Any())
            {
                string sqlsVM = @"select @SortYear as year,a.production_ym,c.vm_code,d.vm_name, b.top_center,c.center_name as sm_name,a.center_code,b.center_name from(
                Select production_ym,center_code From vlife.dbo.aginb as a               
                Where a.production_ym in (select max(production_ym) from aginb) and (ag_status_code in ('0','1')) 
                and occp_ind not in ('0','2','3','4','5','6','7','15') and center_code not in ('0100','0200','1009','1663','1200','1201','8888')
                group by center_code,production_ym) a 
                Inner Join acccb as b          
                on a.center_code=b.center_code and a.production_ym=b.production_ym 
                Inner Join agtcb as c    
                on b.top_center=c.top_center and b.production_ym=c.production_ym 
                inner join mis.dbo.org_vm as d 
                on c.vm_code=d.vm_code 
                group by a.production_ym,c.vm_code,d.vm_name, b.top_center,c.center_name, a.center_code,b.center_name 
                order by a.production_ym,c.vm_code,d.vm_name, b.top_center,c.center_name, a.center_code,b.center_name";
                List<LawReportSortDetail> lawReportSortDetailVM = new List<LawReportSortDetail>();
                lawReportSortDetailVM = DbHelper.Query<LawReportSortDetail>(VLifeRepository.ConnectionStringName, sqlsVM, new { SortYear = SortYear }).ToList();

                for (int i = 0; i < lawReportSortDetailVM.Count; i++)
                {
                    string sqlChkVM = "select * from law_report_sort where 2>1 ";
                    string vm_name = string.Empty;
                    if (lawReportSortDetailVM[i].vm_name.Length > 4){
                        vm_name = lawReportSortDetailVM[i].vm_name.Substring(0, lawReportSortDetailVM[i].vm_name.Length - 2).Trim();
                    }
					else
					{                        
                        vm_name = lawReportSortDetailVM[i].vm_name.Substring(0, 2);
                    }
                    sqlChkVM = sqlChkVM + " and sort_vm_name = @vm_name";
                    sqlChkVM = sqlChkVM + " and sort_sm_name = @sm_name";
                    sqlChkVM = sqlChkVM + " and sort_center_name = @center_name ";
                    sqlChkVM = sqlChkVM + " and sort_year = @year";
                    var ChklawReportSortVM = new LawReportSort();

                    
                    ChklawReportSortVM = DbHelper.Query<LawReportSort>(H2ORepository.ConnectionStringName, sqlChkVM, new
                    {
                        year = lawReportSortDetailVM[i].year,
                        sm_name = lawReportSortDetailVM[i].sm_name.Substring(0, 2),
                        vm_name = vm_name,
                        center_name = lawReportSortDetailVM[i].center_name,
                    }).FirstOrDefault();
                    if(ChklawReportSortVM == null)
					{
                        string sqliVM = @"INSERT INTO law_report_sort(sort_year,sort_vm_name,sort_sm_name,sort_center_name)
                        VALUES (@year, @vm_name, @sm_name, @center_name);";
                        DbHelper.Execute(H2ORepository.ConnectionStringName, sqliVM, new
                        {
                            year = lawReportSortDetailVM[i].year,
                            sm_name = lawReportSortDetailVM[i].sm_name.Substring(0, 2).Trim(),
                            vm_name = vm_name.Trim(),
                            center_name = lawReportSortDetailVM[i].center_name.Trim(),
                        });
                    }
                }

                string sqlsSM = @"select t.year,case right(rtrim(t.vm_name),2) when '副總' then left(t.vm_name,len(t.vm_name)-2) when '區塊' then left(t.vm_name,len(t.vm_name)-2) else t.vm_name end as vm_name
                ,case right(rtrim(t.sm_name),2) when '副總' then left(t.sm_name,len(t.sm_name)-2) else t.sm_name end as sm_name
                ,t.center_name,t.vm_code,t.sm_code,t.center_code from (
                 Select @year as year,vm_name 
                ,sm_name,center_name,vm_code,sm_code,center_code From mis.dbo.orgin_month 
                 where proc_ym like @SortYear
                 and ag_level>='65' 
                 group by vm_name,sm_name,center_name,vm_code,sm_code,center_code ) t  order by vm_code,sm_code,center_code";
                List<LawReportSortDetail> lawReportSortDetailSM = new List<LawReportSortDetail>();
                lawReportSortDetailSM = DbHelper.Query<LawReportSortDetail>(VLifeRepository.ConnectionStringName, sqlsSM, new { year = SortYear, SortYear = "%" + SortYear + "%" }).ToList();

                for (int i = 0; i < lawReportSortDetailSM.Count; i++)
                {
                    string sqlChkSM = "select * from law_report_sort where sort_vm_name = @vm_name ";
                    sqlChkSM = sqlChkSM + " and sort_sm_name = @sm_name";
                    sqlChkSM = sqlChkSM + " and sort_center_name = @center_name";
                    sqlChkSM = sqlChkSM + " and sort_year = @year";
                    var ChklawReportSortSM = new LawReportSort();
                    ChklawReportSortSM = DbHelper.Query<LawReportSort>(H2ORepository.ConnectionStringName, sqlChkSM, new
                    {
                        year = lawReportSortDetailSM[i].year,
                        sm_name = lawReportSortDetailSM[i].sm_name.Trim(),
                        vm_name = lawReportSortDetailSM[i].vm_name.Trim(),
                        center_name = lawReportSortDetailSM[i].center_name.Trim(),
                    }).FirstOrDefault();

                    if(ChklawReportSortSM == null)
					{
                        string sqliSM = @"INSERT INTO law_report_sort(sort_year,sort_vm_name,sort_sm_name,sort_center_name)
                        VALUES (@year, @vm_name, @sm_name, @center_name);";
                        DbHelper.Execute(H2ORepository.ConnectionStringName, sqliSM, new
                        {
                            year = lawReportSortDetailSM[i].year,
                            sm_name = lawReportSortDetailSM[i].sm_name.Trim(),
                            vm_name = lawReportSortDetailSM[i].vm_name.Trim(),
                            center_name = lawReportSortDetailSM[i].center_name.Trim(),
                        });
                    }
                }
                //update團隊有相同名稱的資料
                var sqluVM = @"update a set a.sort_vm=b.sort_vm
                 from law_report_sort as a,law_report_sort as b
                 where a.sort_vm_name=b.sort_vm_name
                 and a.sort_year=b.sort_year
                 and b.sort_year is not null";
                DbHelper.Execute(H2ORepository.ConnectionStringName, sqluVM);
                //update體系有相同名稱的資料
                var sqluSM = @"update a set a.sort_sm=b.sort_sm
                 from law_report_sort as a,law_report_sort as b
                 where a.sort_sm_name=b.sort_sm_name
                 and a.sort_vm_name=b.sort_vm_name
                 and a.sort_year=b.sort_year
                 and b.sort_year is not null";
                DbHelper.Execute(H2ORepository.ConnectionStringName, sqluSM);
            }
        }

        /// <summary>
        /// 報表排序列表VM
        /// </summary>
        public List<LawReportSort> GetLawReportSortVM(string SortYear)
        {
            List<LawReportSort> result = new List<LawReportSort>();
            string sql = @"select sort_year,sort_vm,sort_vm_name from law_report_sort where sort_year= @SortYear
                        group by sort_year,sort_vm,sort_vm_name order by sort_vm,sort_vm_name";

            result = DbHelper.Query<LawReportSort>(H2ORepository.ConnectionStringName, sql, new { SortYear = SortYear }).ToList();
            return result;
        }

        /// <summary>
        /// 報表排序列表SM
        /// </summary>
        public List<LawReportSort> GetLawReportSortSM(string SortYear)
        {
            List<LawReportSort> result = new List<LawReportSort>();
            string sql = @"select sort_year,sort_vm,sort_vm_name,sort_sm,sort_sm_name from law_report_sort
                    where sort_year = @SortYear group by sort_year,sort_vm,sort_vm_name,sort_sm,sort_sm_name order by sort_vm,sort_sm";
            result = DbHelper.Query<LawReportSort>(H2ORepository.ConnectionStringName, sql, new { SortYear = SortYear }).ToList();
            return result;
        }

        /// <summary>
        /// 更新VM報表排序
        /// </summary>
        public bool UpdateLawReportSortVM(string SortVmName, int? SortVm, string SortYear)
        {
            var sql = @"update law_report_sort 
            set sort_vm = @SortVm
            where sort_vm_name = @SortVmName and sort_year = @SortYear";
            if (SortVm == 0)
            {
                SortVm = null;
            }
            int result = -1;
            try
            {
                result = DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
                {
                    SortVm = SortVm,
                    SortVmName = SortVmName,
                    SortYear = SortYear
                });
                return (result > 0);
            }
            catch (Exception x)
            {

            }
            return (result < 0);
        }

        /// <summary>
        /// 更新SM報表排序
        /// </summary>
        public bool UpdateLawReportSortSM(string SortSmName, string SortVmName, int? SortSm, string SortYear)
        {
            var sql = @"update law_report_sort 
            set sort_sm = @SortSm
            where sort_sm_name = @SortSmName and sort_year = @SortYear and sort_vm_name = @SortVmName";
            if (SortSm == 0)
            {
                SortSm = null;
            }
            int result = -1;
            try
            {
                result = DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
                {
                    SortSm = SortSm,
                    SortSmName = SortSmName,
                    SortVmName = SortVmName,
                    SortYear = SortYear
                });
                return (result > 0);
            }
            catch (Exception x)
            {

            }
            return (result < 0);

        }

        /// <summary>
        /// 刪除年度排序
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawReportSortBySortYear(string SortYear)
        {
            int result = -1;
            try
            {
                if (!string.IsNullOrEmpty(SortYear))
                {
                    var sql = @"delete law_report_sort where sort_year= @SortYear";
                    result = DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new { SortYear = SortYear });
                    return (result > 0);
                }
            }
            catch (Exception x)
            {

            }
            return (result < 0);
        }
        #endregion
        #endregion

        #region 受理作業
        /// <summary>
        /// 檢查照會單號是否存在與流水號大小
        /// </summary>
        public LawNoteYmCounter CheckLawNoteYmCounter(string LawNoteYear, string LawNoteMonth)
        {
            LawNoteYmCounter noteYmCounter = new LawNoteYmCounter();
            string sql = @"select * from law_note_ym_counter where law_note_year = @LawNoteYear and law_note_month = @LawNoteMonth";
            noteYmCounter = DbHelper.Query<LawNoteYmCounter>(H2ORepository.ConnectionStringName, sql, new { LawNoteYear = LawNoteYear, LawNoteMonth = LawNoteMonth }).FirstOrDefault();

            return noteYmCounter;
        }

        /// <summary>
        /// 檢查個人資料是否重複
        /// </summary>
        public LawAgentData CheckLawAgentDataRepeat(string agent_code)
        {
            LawAgentData lawAgentData = new LawAgentData();
            string sql = @"select * from law_agent_data where agent_code = @agent_code";
            lawAgentData = DbHelper.Query<LawAgentData>(H2ORepository.ConnectionStringName, sql, new { agent_code = agent_code }).FirstOrDefault();

            return lawAgentData;
        }

        /// <summary>
        /// 檢查業務員相關資料
        /// </summary>
        public OrgibDetail CheckOrgib(string agent_code, string production_ym, string sequence)
        {
            OrgibDetail noteYmCounter = new OrgibDetail();
            string sql = @"select a.sm_code,a.sm_name,a.wc_center,a.wc_center_name,a.center_code,a.center_name,a.administrat_id,a.admin_name,a.admin_level,a.agent_code,rtrim(c.names) as names,a.ag_status_code,a.ag_level
            ,b.level_name_chs,c.birth,c.cellur_phone_no,d.record_date,e.register_date,e.ag_status_date from orgib a
            left join agln b on a.ag_level=b.ag_level 
            left join nain c on left(a.agent_code,10)=c.client_id 
            left join agid d on a.agent_code=d.agent_code 
            left join agin e on a.agent_code=e.agent_code 
            where a.agent_code= @agent_code and a.production_ym = @production_ym and a.sequence= @sequence";

            noteYmCounter = DbHelper.Query<OrgibDetail>(VLifeRepository.ConnectionStringName, sql, new { agent_code = agent_code, production_ym = production_ym, sequence = sequence }).FirstOrDefault();

            return noteYmCounter;
        }

        /// <summary>
        /// 檢查八大區塊
        /// </summary>
        public orgsmb Checkorgsmb(string agent_code)
        {
            orgsmb orgsmb = new orgsmb();
            string sql = "select * from org_smb where production_ym=(select max(production_ym) from org_smb where sm_code in(select sm_code from orgin where agent_code = @agent_code)";
            sql = sql + " ) and sm_code=(select sm_code from orgin where agent_code = @agent_code)";

            orgsmb = DbHelper.Query<orgsmb>(VLifeRepository.ConnectionStringName, sql, new { agent_code = agent_code }).FirstOrDefault();

            return orgsmb;
        }

        /// <summary>
        /// 檢查個人資料是否重複
        /// </summary>
        public LawAgentData CheckLawAgentData(string agent_code)
        {
            LawAgentData lawAgentData = new LawAgentData();
            string sql = "select * from law_agent_data where agent_code = @agent_code";

            lawAgentData = DbHelper.Query<LawAgentData>(VLifeRepository.ConnectionStringName, sql, new { agent_code = agent_code }).FirstOrDefault();

            return lawAgentData;
        }

        /// <summary>
        /// 抓取第一順位經辦人員
        /// </summary>
        public LawDoUser GetLawDouserTopOne()
        {
            LawDoUser result = new LawDoUser();
            string sql = @"select top 1 law_douser_id,douser_orgid,substring(douser_name,patindex('% %',douser_name)+1,len(douser_name)-patindex('% %',douser_name)) as douser_name,douser_phone_ext,douser_sort from law_douser order by douser_sort";

            result = DbHelper.Query<LawDoUser>(H2ORepository.ConnectionStringName, sql).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 新增照會單號編號記錄檔
        /// </summary>
        public void InsertLawNoteYmCounter(LawNoteYmCounter noteYmCounter)
        {
            string sqli = @"insert into law_note_ym_counter(law_note_year, law_note_month, law_note_counter) values(@LawNoteYear,@LawNoteMonth, 1)";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sqli, new
            {
                LawNoteYear = noteYmCounter.LawNoteYear,
                LawNoteMonth = noteYmCounter.LawNoteMonth,
            });
        }

        /// <summary>
        /// 更新照會單號編號記錄檔
        /// </summary>
        public void UpdateLawNoteYmCounter(LawNoteYmCounter noteYmCounter)
        {
            string sqli = @"update law_note_ym_counter 
                set law_note_counter = @LawNoteCounter
                where law_note_year = @LawNoteYear and law_note_month = @LawNoteMonth";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sqli, new
            {
                LawNoteYear = noteYmCounter.LawNoteYear,
                LawNoteMonth = noteYmCounter.LawNoteMonth,
                LawNoteCounter = noteYmCounter.LawNoteCounter + 1
            });
        }

        /// <summary>
        /// 新增法追主檔
        /// </summary>
        public void InsertLawContent(LawContent lawContent)
        {
            string sql = @"insert into law_content (law_note_no,law_year,law_month,law_pay_sequence,law_file_no,law_due_name,law_due_agentid,law_due_money,law_super_account,law_due_reason,law_status_type,LawContentCreatorID,law_content_create_name,law_content_create_date,law_do_unit_id,law_do_unit_name)
            values(@LawNoteNo,@LawYear,@LawMonth,@LawPaySequence,@LawFileNo,@LawDueName,@LawDueAgentId,@LawDueMoney,@LawSuperAccount,@LawDueReason,@LawStatusType,@LawContentCreatorID,@LawContentCreateName,@LawContentCreateDate,@LawDoUnitId,@LawDoUnitName)";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawNoteNo = lawContent.LawNoteNo,
                LawYear = lawContent.LawYear,
                LawMonth = lawContent.LawMonth,
                LawPaySequence = lawContent.LawPaySequence,
                LawFileNo = lawContent.LawFileNo,
                LawDueName = lawContent.LawDueName,
                LawDueAgentId = lawContent.LawDueAgentId,
                LawDueMoney = lawContent.LawDueMoney,
                LawSuperAccount = lawContent.LawSuperAccount,
                LawDueReason = lawContent.LawDueReason,
                LawStatusType = lawContent.LawStatusType,
                LawContentCreatorID = lawContent.LawContentCreatorID,
                LawContentCreateName = lawContent.LawContentCreateName,
                LawContentCreateDate = lawContent.LawContentCreateDate,
                LawDoUnitId = lawContent.LawDoUnitId,
                LawDoUnitName = lawContent.LawDoUnitName
            });
        }

        /// <summary>
        /// 取得業務員地址(戶籍、住家)與電話
        /// </summary>
        public adin Getadin(string agent_code)
        {
            adin adin = new adin();
            string sql = @"select * from adin where client_id=@agent_code";

            adin = DbHelper.Query<adin>(VLifeRepository.ConnectionStringName, sql, new { agent_code = agent_code }).FirstOrDefault();

            return adin;
        }

        /// <summary>
        /// 取得三代輔導主管
        /// </summary>
        public OrginDetail Getorgin(string agent_code)
        {
            OrginDetail orginDetail = new OrginDetail();
            string sql = @"select a.center_name,a.ag_status_code,a.name
            ,a.director_id ,a.director_name,b.ag_status_code as dir_status_code
            ,case when d.term_code not in('0','1','2','3','4','5','6','7','10','15') then rtrim(term_meaning) else rtrim(c.level_name_chs) end as level_name_chs
            ,d.term_code,d.term_meaning 
            from orgin as a,orgin as b 
            left join agln as c on b.ag_level=c.ag_level 
            left join trmval as d on term_id='occp_ind' and b.ag_occp_ind=d.term_code 
            where a.agent_code = @agent_code
            and a.director_id =b.agent_code";

            orginDetail = DbHelper.Query<OrginDetail>(VLifeRepository.ConnectionStringName, sql, new { agent_code = agent_code }).FirstOrDefault();

            return orginDetail;
        }

        /// <summary>
        /// 依照照會單號取得主檔ID
        /// </summary>
        public LawContent GetLawContentByLawNoteNo(string LawNoteNo)
        {
            LawContent lawContent = new LawContent();
            string sql = @"select * from law_content where law_note_no = @LawNoteNo";

            lawContent = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { LawNoteNo = LawNoteNo }).FirstOrDefault();

            return lawContent;
        }

        /// <summary>
        /// 新增業務員資料檔
        /// </summary>
        public void InsertLawAgentContent(LawAgentContent model)
        {
            string sql = @"insert into law_agent_content (law_note_no,production_ym,sequence,vm_code,vm_name,sm_code,sm_name,wc_center,wc_center_name,center_code,center_name,administrat_id,admin_name,admin_level,agent_code,name,ag_status_code,ag_level,ag_level_name,record_date,register_date,ag_status_date,super_account,create_date)
                                                    values(@LawNoteNo,@ProductionYm,@Sequence,@VmCode,@VmName,@SmCode,@SmName,@WcCenter,@WcCenterName,@CenterCode,@CenterName,@AdministratId,@AdminName,@AdminLevel,@AgentCode,@Name,@AgStatusCode,@AgLevel,@AgLevelName,@RecordDate,@RegisterDate,@AgStatusDate,@SuperAccount,@CreateDate)";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawNoteNo = model.LawNoteNo,
                ProductionYm = model.ProductionYm,
                Sequence = model.Sequence,
                VmCode = model.VmCode,
                VmName = model.VmName,
                SmCode = model.SmCode,
                SmName = model.SmName,
                WcCenter = model.WcCenter,
                WcCenterName = model.WcCenterName,
                CenterCode = model.CenterCode,
                CenterName = model.CenterName,
                AdministratId = model.AdministratId,
                AdminName = model.AdminName,
                AdminLevel = model.AdminLevel,
                AgentCode = model.AgentCode,
                Name = model.Name,
                AgStatusCode = model.AgStatusCode,
                AgLevel = model.AgLevel,
                AgLevelName = model.AgLevelName,
                RecordDate = model.RecordDate,
                RegisterDate = model.RegisterDate,
                AgStatusDate = model.AgStatusDate,
                SuperAccount = model.SuperAccount,
                CreateDate = model.CreateDate,
            });
        }

        /// <summary>
        /// 新增業務員個人資料檔
        /// </summary>
        public void InsertLawAgentData(LawAgentData model)
        {
            string sql = @"insert into law_agent_data(agent_code,name,birth,cell_01,phone_01,address,address_01,super_account,create_date) 
            values(@AgentCode,@Name,@Birth,@Cell01,@Phone01,@Address,@Address01,@SuperAccount,@CreateDate)";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                AgentCode = model.AgentCode,
                Name = model.Name,
                Birth = model.Birth,
                Cell01 = model.Cell01,
                Phone01 = model.Phone01,
                Address = model.Address,
                Address01 = model.Address01,
                SuperAccount = model.SuperAccount,
                CreateDate = model.CreateDate,
            });
        }

        /// <summary>
        /// 更新業務員個人資料檔
        /// </summary>
        public void UpdateLawAgentData(string agentcode, string name, string Rename)
        {
            string sqli = @"update law_agent_data set name = @name 
                where agent_code = @agentcode and name = @Rename ";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sqli, new
            {
                agentcode = agentcode,
                name = name,
                Rename = Rename
            });
        }

        /// <summary>
        /// 更新業務員資料檔
        /// </summary>
        public void UpdateLawAgentContent(string agentcode, string name, string Rename)
        {
            string sqli = @"update law_agent_content set name = @name 
                where left(agent_code,10) = @agentcode and name = @Rename";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sqli, new
            {
                agentcode = agentcode,
                name = name,
                Rename = Rename
            });
        }

        /// <summary>
        /// 新增業務員個人資料檔
        /// </summary>
        public void InsertLawNote(LawNote model)
        {
            string sql = @"insert into law_note(law_note_no,law_note_km,law_note_center,law_note_to,law_note_level,law_note_name,law_note_dep,law_note_pro,law_note_tel,law_note_type,law_note_creatername,law_note_createdate,NoteCreatorID)
            values(@LawNoteNo,@LawNoteKm,@LawNoteCenter,@LawNoteTo,@LawNoteLevel,@LawNoteName,@LawNoteDep,@LawNotePro,@LawNoteTel,@LawNoteType,@LawNoteCreatername,@LawNoteCreatedate,@NoteCreatorID)";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawNoteNo = model.LawNoteNo,
                LawNoteKm = model.LawNoteKm,
                LawNoteCenter = model.LawNoteCenter,
                LawNoteTo = model.LawNoteTo,
                LawNoteLevel = model.LawNoteLevel,
                LawNoteName = model.LawNoteName,
                LawNoteDep = model.LawNoteDep,
                LawNotePro = model.LawNotePro,
                LawNoteTel = model.LawNoteTel,
                LawNoteType = model.LawNoteType,
                LawNoteCreatername = model.LawNoteCreatername,
                LawNoteCreatedate = model.LawNoteCreatedate,
                NoteCreatorID = model.NoteCreatorID,
            });
        }

        /// <summary>
        /// 新增進度表訴訟程序記錄檔
        /// </summary>
        public void InsertLawLitigationProgress(LawLitigationProgress model)
        {
            string sql = @"insert into law_litigation_progress(law_note_no,law_id,law_litigation_progress,create_date,LawLitigationProgressCreatorID)
            values(@LawNoteNo,@LawId,@LawLitigationprogress,@CreateDate,@LawLitigationProgressCreatorID)";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawNoteNo = model.LawNoteNo,
                LawId = model.LawId,
                LawLitigationprogress = model.LawLitigationprogress,
                CreateDate = model.CreateDate,
                LawLitigationProgressCreatorID = model.LawLitigationProgressCreatorID,
            });
        }

        /// <summary>
        /// 新增進度表執行程序紀錄
        /// </summary>
        public void InsertLawDoProgress(LawDoProgress model)
        {
            string sql = @"insert into law_do_progress(law_note_no,law_id,law_do_progress,create_date,LawDoProgressCreatorID)
            values(@LawNoteNo,@LawId,@LawDoprogress,@CreateDate,@LawDoProgressCreatorID)";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawNoteNo = model.LawNoteNo,
                LawId = model.LawId,
                LawDoprogress = model.LawDoprogress,
                CreateDate = model.CreateDate,
                LawDoProgressCreatorID = model.LawDoProgressCreatorID,
            });
        }
        #endregion

        #region 通知作業
        /// <summary>
        /// 通知作業查詢列表
        /// </summary>
        public List<LawNote> GetLawNote(string LawNoteNo, int sys, string MemberID, string LawNoteName, string orgID)
        {
            List<LawNote> result = new List<LawNote>();
            string sql = @"Select distinct law_note_no,law_note_name,CONVERT(varchar(12),law_note_createdate,111) as law_note_createdate  From law_note where 2>1";
            if (sys == 1)
            {
                sql += " and (law_note_no in (select mg_no from k_message_to where mg_no <> '' and(mg_no like '%照%') group by mg_no) or law_note_no in (select MSGOBJNote from MessageTo where MSGOBJNote<>'' and (MSGOBJNote like '%照%') group by MSGOBJNote))";
            }
            else
            {
                sql += " and (law_note_no in (select mg_no from k_message_to where mo_orgid in ('" + orgID + "') and mg_no<>'' and (mg_no like '%照%' ) group by mg_no) or law_note_no in (select MSGOBJNote from MessageTo where MSGOBJID in (@MemberID) and MSGOBJNote<>'' and (MSGOBJNote like '%照%' ) group by MSGOBJNote))";
            }

            if (!string.IsNullOrWhiteSpace(LawNoteNo))
            {
                sql += " and law_note_no like @LawNoteNo";
            }
            if (!string.IsNullOrWhiteSpace(LawNoteName))
            {
                sql += " and law_note_name like @LawNoteName";
            }

            sql += " group by law_note_no,law_note_name,law_note_createdate order by law_note_no desc";

            result = DbHelper.Query<LawNote>(H2ORepository.ConnectionStringName, sql, new { LawNoteNo = "%" + LawNoteNo + "%", MemberID = MemberID, LawNoteName = "%" + LawNoteName + "%" }).ToList();
            return result;
        }

        /// <summary>
        /// 系統設定作業目錄,判斷是否為系統管理人員權限
        /// </summary>
        public int CheckLawNoteByMemberID(string MemberID)
        {
            LawNote lawnote = new LawNote();
            int result = -1;
            string sql = @"select * from law_sys where SysMemberID = @MemberID";

            lawnote = DbHelper.Query<LawNote>(H2ORepository.ConnectionStringName, sql, new { MemberID = MemberID }).FirstOrDefault();
            if (lawnote != null)
            {
                result = 1;
            }
            return result;
        }

        #endregion

        #region 維護作業
        #region 維護作業查詢頁面&個人資料&進度表
        /// <summary>
        /// 維護作業查詢列表
        /// </summary>
        public List<LawContent> GetLawContent(LawContent model)
        {
            List<LawContent> result = new List<LawContent>();
            string sql = @"select * from law_content where 2>1";
            switch (model.LawStatusType)
            {
                case 1:
                    sql += " and law_status_type not in (2) and law_note_no not in(select law_note_no from law_content where law_not_close_type_name='執行無實益'  and law_status_type not in (2))";
                    break;
                case 2:
                    sql += " and law_status_type in (2)";
                    break;
                case 3:
                    sql += " and law_close_type like('%6%') and law_status_type not in (2) or (law_not_close_type_name='執行無實益' and law_status_type not in (2))";
                    break;
            }
            if (!string.IsNullOrWhiteSpace(model.LawNoteNo))
            {
                sql += " and law_note_no like @LawNoteNo";
            }
            if (!string.IsNullOrWhiteSpace(model.LawDueName))
            {
                sql += " and law_due_name like @LawDueName";
            }
            if (!string.IsNullOrWhiteSpace(model.LawDueAgentId))
            {
                sql += " and law_due_agentid like @LawDueAgentId";
            }
            if (model.LawDueMoney != 0)
            {
                sql += " and law_due_money like @LawDueMoney";
            }
            if (!string.IsNullOrWhiteSpace(model.LawDoUnitName))
            {
                sql += " and law_do_unit_name like @LawDoUnitName";
            }

            sql += " order by law_year,law_note_no desc";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new
            {
                LawNoteNo = "%" + model.LawNoteNo + "%",
                LawDueName = "%" + model.LawDueName + "%",
                LawDueAgentId = "%" + model.LawDueAgentId + "%",
                LawDueMoney = "%" + model.LawDueMoney + "%",
                LawDoUnitName = "%" + model.LawDoUnitName + "%"
            }).ToList();
            return result;
        }

        /// <summary>
        /// 取得啟用狀態承辦單位
        /// </summary>
        public List<LawDoUnit> GetLawDoUnitByEnable()
        {
            List<LawDoUnit> result = new List<LawDoUnit>();
            string sql = @"select * from law_do_unit where status_type = 1";

            result = DbHelper.Query<LawDoUnit>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 取得個人資料頁面
        /// </summary>
        public LawAgentDetail GetLawAgentDetail(string AgentID, string NoteNo)
        {
            LawAgentDetail result = new LawAgentDetail();
            string sql = @"select * from law_agent_content a left join law_agent_data b on left(a.agent_code,10)=b.agent_code where a.agent_code = @AgentID and a.law_note_no = @NoteNo";

            result = DbHelper.Query<LawAgentDetail>(H2ORepository.ConnectionStringName, sql, new { AgentID = AgentID, NoteNo = NoteNo }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 取得前案情形
        /// </summary>
        public List<LawContent> GetLawContentByIDNo(string AgentID, string NoteNo)
        {
            List<LawContent> result = new List<LawContent>();
            string sql = @"select * from law_content where left(law_due_agentid,10) = @AgentID and law_note_no <> @NoteNo";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { AgentID = AgentID, NoteNo = NoteNo }).ToList();
            return result;
        }

        /// <summary>
        /// 更新業務員資料檔
        /// </summary>
        public void UpdateLawAgentDetail(LawAgentDetail agentDetail)
        {
            if (agentDetail.oldname != agentDetail.Name)
            {
                string sqld = @"update law_agent_data set cell_01=@Cell01,cell_02=@Cell02,phone_01=@Phone01,phone_02=@Phone02,address=@Address,address_01=@Address01,address_02=@Address02,name = @Name " +
                    " where agent_code=@AgentID and law_agent_data_id=@LawAgentDataId";
                DbHelper.Execute(H2ORepository.ConnectionStringName, sqld, new
                {
                    Name = agentDetail.Name,
                    Cell01 = agentDetail.Cell01,
                    Cell02 = agentDetail.Cell02,
                    Phone01 = agentDetail.Phone01,
                    Phone02 = agentDetail.Phone02,
                    Address = agentDetail.Address,
                    Address01 = agentDetail.Address01,
                    Address02 = agentDetail.Address02,
                    LawAgentDataId = agentDetail.LawAgentDataId,
                    AgentID = agentDetail.AgentID
                });

                string sqlc = @"update law_agent_content set name=@Name" +
                    " where left(agent_code,10)=@AgentID and name=@oldName";
                DbHelper.Execute(H2ORepository.ConnectionStringName, sqlc, new
                {
                    Name = agentDetail.Name,
                    AgentID = agentDetail.AgentID,
                    oldName = agentDetail.oldname
                });
            }
            else
            {
                string sqla = @"update law_agent_data set cell_01=@Cell01,cell_02=@Cell02,phone_01=@Phone01,phone_02=@Phone02,address=@Address,address_01=@Address01,address_02=@Address02 " +
                    " where agent_code=@AgentID and law_agent_data_id=@LawAgentDataId";
                DbHelper.Execute(H2ORepository.ConnectionStringName, sqla, new
                {
                    Cell01 = agentDetail.Cell01,
                    Cell02 = agentDetail.Cell02,
                    Phone01 = agentDetail.Phone01,
                    Phone02 = agentDetail.Phone02,
                    Address = agentDetail.Address,
                    Address01 = agentDetail.Address01,
                    Address02 = agentDetail.Address02,
                    LawAgentDataId = agentDetail.LawAgentDataId,
                    AgentID = agentDetail.AgentID
                });
            }
        }

        /// <summary>
        /// 取得進度表頁面
        /// </summary>
        public LawContent GetSchedule(string AgentID, string NoteNo)
        {
            LawContent result = new LawContent();
            string sql = @"select law_id,law_note_no,law_year,law_month,law_pay_sequence,law_file_no,law_due_name,law_due_agentid,law_due_money,law_interest_sdate,law_interest_edate,law_interest_days,law_interest_money,law_total_money
            ,law_super_account,law_due_reason,law_status_type,law_do_unit_id,law_do_unit_name,law_douser_orgid,law_douser_name,law_not_close_type_name,law_close_type,convert(varchar(50),law_close_date,111) as law_close_date
            ,law_closer_orgid,law_closer_name,law_content_create_orgid,law_content_create_name,convert(varchar(50),law_content_create_date,111) as law_content_create_date,law_phone_call1_desc,convert(varchar(50)
            ,law_phone_call1_date,111) as law_phone_call1_date,law_phone_call2_desc,convert(varchar(50),law_phone_call2_date,111) as law_phone_call2_date,law_content_cancel_type,law_content_lastchange_date
             from law_content where left(law_due_agentid,10) = @AgentID and law_note_no = @NoteNo";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { AgentID = AgentID, NoteNo = NoteNo }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 取得進度表存證信函備註記
        /// </summary>
        public List<LawEvidenceDesc> GetLawEvidenceDesc(string LawId)
        {
            List<LawEvidenceDesc> result = new List<LawEvidenceDesc>();
            string sql = @"select * from law_evidence_desc where law_id= @LawId";

            result = DbHelper.Query<LawEvidenceDesc>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId }).ToList();
            return result;
        }

        /// <summary>
        /// 取得分案日期
        /// </summary>
        public LawDoUnitLog GetLawDoUnitLog(string LawId)
        {
            LawDoUnitLog result = new LawDoUnitLog();
            string sql = @"select *,convert(varchar(50),create_date,111) as createdate from law_do_unit_log where law_dounit_log_id in(select max(law_dounit_log_id) from law_do_unit_log where law_id=@LawId and law_dounit_log_id is not null )";

            result = DbHelper.Query<LawDoUnitLog>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 取得啟用狀態承辦單位ID
        /// </summary>
        public LawDoUnit GetLawDoUnitByUnitName(string UnitName)
        {
            LawDoUnit result = new LawDoUnit();
            string sql = @"select * from law_do_unit where unit_name = @UnitName";

            result = DbHelper.Query<LawDoUnit>(H2ORepository.ConnectionStringName, sql, new { UnitName = UnitName }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 取得承辦單位歷史紀錄
        /// </summary>
        public List<LawDoUnitLog> GetLawDoUnitLogByLawId(string LawId)
        {
            List<LawDoUnitLog> result = new List<LawDoUnitLog>();
            string sql = @"select * from law_do_unit_log where law_id=@LawId order by create_date desc";

            result = DbHelper.Query<LawDoUnitLog>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId }).ToList();
            return result;
        }

        /// <summary>
        /// 取得進度表訴訟程序
        /// </summary>
        public List<LawLitigationProgress> GetLawLitigationProgressByLawIdNotoNo(string LawId, string LawNoteNo)
        {
            List<LawLitigationProgress> result = new List<LawLitigationProgress>();
            string sql = @"select * from law_litigation_progress where law_id=@LawId and law_note_no=@LawNoteNo order by law_litigation_id";

            result = DbHelper.Query<LawLitigationProgress>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId, LawNoteNo = LawNoteNo }).ToList();
            return result;
        }

        /// <summary>
        /// 取得進度表執行程序
        /// </summary>
        public List<LawDoProgress> GetLawDoProgressByLawIdNotoNo(string LawId, string LawNoteNo)
        {
            List<LawDoProgress> result = new List<LawDoProgress>();
            string sql = @"select * from law_do_progress where law_id=@LawId and law_note_no=@LawNoteNo order by law_do_id";

            result = DbHelper.Query<LawDoProgress>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId, LawNoteNo = LawNoteNo }).ToList();
            return result;
        }

        /// <summary>
        /// 取得進度表其他說明
        /// </summary>
        public List<LawOtherDescDetail> GetLawOtherDescByLawIdNotoNo(string LawId, string LawNoteNo)
        {
            List<LawOtherDescDetail> result = new List<LawOtherDescDetail>();
            string sql = @"select a.law_other_id,a.law_id,a.law_note_no,a.law_other_desc,a.create_orgid,a.create_date,b.law_repayment_id,convert(varchar(50),b.law_repayment_date,111)as law_repayment_date,b.law_repayment_money,b.law_comm_deduction 
            from law_other_desc a left join law_repayment_list b on a.law_id=b.law_id and a.law_repayment_id=b.law_repayment_id where a.law_id=@LawId and a.law_note_no=@LawNoteNo";

            result = DbHelper.Query<LawOtherDescDetail>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId, LawNoteNo = LawNoteNo }).ToList();
            return result;
        }

        /// <summary>
        /// 取得進度表案件狀態其他說明
        /// </summary>
        public LawCloseOtherLog GetLawCloseOtherLogByLawId(string LawId)
        {
            LawCloseOtherLog result = new LawCloseOtherLog();
            string sql = @"select * from law_close_other_log where law_id= @LawId";

            result = DbHelper.Query<LawCloseOtherLog>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 取得進度表清償金額
        /// </summary>
        public List<LawRepaymentList> GetLawRepaymentListByLawIdNotoNo(string AgentID, string LawId, string LawNoteNo)
        {
            List<LawRepaymentList> result = new List<LawRepaymentList>();
            string sql = @"select law_repayment_id,law_id,law_note_no,law_due_agentid,convert(varchar(50),law_repayment_date,111)as law_repayment_date,law_repayment_money,law_comm_deduction,create_orgid,create_date 
            from law_repayment_list where law_due_agentid=@AgentID and law_id=@LawId and law_repayment_id not in (
            select law_repayment_id from law_other_desc where law_id = @LawId and law_note_no = @LawNoteNo and(law_repayment_id <> '' or law_repayment_id is not null))";

            result = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql, new { AgentID = AgentID, LawId = LawId, LawNoteNo = LawNoteNo }).ToList();
            return result;
        }

        /// <summary>
        /// 取得進度表備註
        /// </summary>
        public List<LawDescDesc> GetLawDescDescByLawIdNotoNo(string LawId, string LawNoteNo)
        {
            List<LawDescDesc> result = new List<LawDescDesc>();
            string sql = @"select * from law_desc_desc where law_id = @LawId and law_note_no = @LawNoteNo";

            result = DbHelper.Query<LawDescDesc>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId, LawNoteNo = LawNoteNo }).ToList();
            return result;
        }

        /// <summary>
        /// 取得啟用狀態結案狀態設定
        /// </summary>
        public List<LawCloseType> GetLawCloseTypeByEnable()
        {
            List<LawCloseType> result = new List<LawCloseType>();
            string sql = @"select * from law_close_type where status_type=1";

            result = DbHelper.Query<LawCloseType>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 存證信函編輯頁面
        /// </summary>
        public LawEvidenceDesc GetLawEvidenceDescById(int LawEvidenceId)
        {
            LawEvidenceDesc result = new LawEvidenceDesc();
            string sql = @"select * from law_evidence_desc where law_evidence_id = @LawEvidenceId";

            result = DbHelper.Query<LawEvidenceDesc>(H2ORepository.ConnectionStringName, sql, new { LawEvidenceId = LawEvidenceId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 訴訟程序編輯頁面
        /// </summary>
        public LawLitigationProgress GetLawLitigationProgressById(int LawLitigationId)
        {
            LawLitigationProgress result = new LawLitigationProgress();
            string sql = @"select * from law_litigation_progress where law_litigation_id= @LawLitigationId";

            result = DbHelper.Query<LawLitigationProgress>(H2ORepository.ConnectionStringName, sql, new { LawLitigationId = LawLitigationId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 執行程序編輯頁面
        /// </summary>
        public LawDoProgress GetLawDoProgressById(int LawDoId)
        {
            LawDoProgress result = new LawDoProgress();
            string sql = @"select * from law_do_progress where law_do_id = @LawDoId";

            result = DbHelper.Query<LawDoProgress>(H2ORepository.ConnectionStringName, sql, new { LawDoId = LawDoId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 其他說明編輯頁面
        /// </summary>
        public LawOtherDesc GetLawOtherDescById(int LawOtherId)
        {
            LawOtherDesc result = new LawOtherDesc();
            string sql = @"select * from law_other_desc where law_other_id = @LawOtherId";

            result = DbHelper.Query<LawOtherDesc>(H2ORepository.ConnectionStringName, sql, new { LawOtherId = LawOtherId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 其他金額編輯頁面
        /// </summary>
        public LawOtherDesc GetLawOtherDescByLawId(string LawDueAgentId, int LawId, int LawRepaymentId)
        {
            LawOtherDesc result = new LawOtherDesc();
            string sql = @"select * from law_other_desc where  law_id=@LawId and left(law_due_agentid,10)=@LawDueAgentId and law_repayment_id is not null and law_repayment_id = @LawRepaymentId and law_repayment_money is not null ";

            result = DbHelper.Query<LawOtherDesc>(H2ORepository.ConnectionStringName, sql, new { LawDueAgentId = LawDueAgentId, LawId = LawId, LawRepaymentId = LawRepaymentId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 其他金額說明編輯頁面
        /// </summary>
        public LawRepaymentList GetLawRepaymentListById(int LawRepaymentId)
        {
            LawRepaymentList result = new LawRepaymentList();
            string sql = @"select * from law_repayment_list where law_repayment_id = @LawRepaymentId";

            result = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql, new { LawRepaymentId = LawRepaymentId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 備註編輯頁面
        /// </summary>
        public LawDescDesc GetLawLawDescDescById(int LawDescId)
        {
            LawDescDesc result = new LawDescDesc();
            string sql = @"select * from law_desc_desc where law_desc_id = @LawDescId";

            result = DbHelper.Query<LawDescDesc>(H2ORepository.ConnectionStringName, sql, new { LawDescId = LawDescId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 新增進度表存證信函備註記
        /// </summary>
        public void InsertLawEvidenceDesc(LawEvidenceDesc model)
        {
            string sql = @"INSERT INTO law_evidence_desc(law_id,law_note_no,law_evidence_desc,create_date,EvidenceDescCreatorID)
                           VALUES (@LawId, @LawNoteNo, @LawEvidencedesc, @CreateDate, @EvidenceDescCreatorID);";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                CreateDate = model.CreateDate,
                EvidenceDescCreatorID = model.EvidenceDescCreatorID,
                LawEvidencedesc = model.LawEvidencedesc,
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
            });
        }

        /// <summary>
        /// 新增其他
        /// </summary>
        public void InsertLawOtherDesc(LawOtherDesc model)
        {
            string sql = @"INSERT INTO law_other_desc(law_id,law_note_no,law_other_desc,create_date,law_due_agentid,OtherDescCreatorID)
                           VALUES (@LawId, @LawNoteNo, @LawOtherdesc, @CreateDate, @LawDueAgentId,@OtherDescCreatorID);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
                LawOtherdesc = model.LawOtherdesc,
                CreateDate = model.CreateDate,
                OtherDescCreatorID = model.OtherDescCreatorID,
                LawDueAgentId = model.LawDueAgentId
            });
        }

        /// <summary>
        /// 新增其他金額說明
        /// </summary>
        public void InsertLawOtherDescHaveMoney(LawOtherDesc model)
        {
            string sql = @"INSERT INTO law_other_desc(law_id,law_note_no,law_other_desc,create_date,law_due_agentid,OtherDescCreatorID,law_repayment_id,law_repayment_money)
                           VALUES (@LawId, @LawNoteNo, @LawOtherdesc, @CreateDate, @LawDueAgentId,@OtherDescCreatorID,@LawRepaymentId,@LawRepaymentMoney);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawRepaymentMoney = model.LawRepaymentMoney,
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
                LawOtherdesc = model.LawOtherdesc,
                CreateDate = model.CreateDate,
                OtherDescCreatorID = model.OtherDescCreatorID,
                LawDueAgentId = model.LawDueAgentId,
                LawRepaymentId = model.LawRepaymentId
            });
        }

        /// <summary>
        /// 新增備註
        /// </summary>
        public void InsertLawDescDesc(LawDescDesc model)
        {
            string sql = @"INSERT INTO law_desc_desc(law_id,law_note_no,law_desc_desc,create_date,LawDescDescCreatorID)
                           VALUES (@LawId, @LawNoteNo, @LawDescdesc, @CreateDate, @LawDescDescCreatorID);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
                LawDescdesc = model.LawDescdesc,
                CreateDate = model.CreateDate,
                LawDescDescCreatorID = model.LawDescDescCreatorID,
            });
        }

        /// <summary>
        /// 新增承辦單位設定記錄log和更新主檔承辦單位
        /// </summary>
        public void InsertLawDoUnitLogUpdateLawContentByLawDoUnitId(LawDoUnitLog model, LawContent content)
        {
            string sql = @"INSERT INTO law_do_unit_log(law_id,law_note_no,law_do_unit_id,case_date,create_name,create_date,DounitLogCreatorID,unit_name)
                           VALUES (@LawId, @LawNoteNo, @LawDoUnitId,@CaseDate,@CreateName, @CreateDate, @DounitLogCreatorID,@UnitName);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
                UnitName = content.LawDoUnitName,
                LawDoUnitId = content.LawDoUnitId,
                CaseDate = model.CaseDate,
                CreateName = model.CreateName,
                CreateDate = model.CreateDate,
                DounitLogCreatorID = model.DounitLogCreatorID,
            });

            var sql1 = @"update law_content set law_do_unit_id=@LawDoUnitId,law_do_unit_name=@LawDoUnitName where law_id=@LawId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql1, new
            { LawId = content.LawId, LawDoUnitName = content.LawDoUnitName, LawDoUnitId = content.LawDoUnitId });
        }

        /// <summary>
        /// 新增電催通知記錄log和更新主檔承辦單位
        /// </summary>
        public void InsertLawPhoneCallLogUpdateLawContentByLawId(LawPhoneCallLog model, LawContent content)
        {
            string sql = @"INSERT INTO law_phone_call_log(law_id,law_note_no,law_phone_call_no,law_phone_call_date,law_phone_call_limited_date,law_phone_call_read_log,PhoneCallCreatorID)
                           VALUES (@LawId, @LawNoteNo, @LawPhoneCallNo,@LawPhoneCallDate,@LawPhoneCallLimitedDate, @LawPhoneCallReadLog, @PhoneCallCreatorID);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
                LawPhoneCallNo = model.LawPhoneCallNo,
                LawPhoneCallDate = model.LawPhoneCallDate,
                LawPhoneCallLimitedDate = model.LawPhoneCallLimitedDate,
                LawPhoneCallReadLog = model.LawPhoneCallReadLog,
                PhoneCallCreatorID = model.PhoneCallCreatorID,
            });

            var sql1 = @"update law_content set law_phone_call1_desc = @LawPhoneCall1Desc ,law_phone_call1_date=@LawPhoneCall1Date where law_id=@LawId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql1, new
            { LawId = content.LawId, LawPhoneCall1Desc = content.LawPhoneCall1Desc, LawPhoneCall1Date = content.LawPhoneCall1Date });
        }

        /// <summary>
        /// 新增地2次電催通知記錄log和更新主檔承辦單位
        /// </summary>
        public void InsertLawPhoneCallLog2UpdateLawContentByLawId(LawPhoneCallLog model, LawContent content)
        {
            string sql = @"INSERT INTO law_phone_call_log(law_id,law_note_no,law_phone_call_no,law_phone_call_date,law_phone_call_limited_date,law_phone_call_read_log,PhoneCallCreatorID)
                           VALUES (@LawId, @LawNoteNo, @LawPhoneCallNo,@LawPhoneCallDate,@LawPhoneCallLimitedDate, @LawPhoneCallReadLog, @PhoneCallCreatorID);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
                LawPhoneCallNo = model.LawPhoneCallNo,
                LawPhoneCallDate = model.LawPhoneCallDate,
                LawPhoneCallLimitedDate = model.LawPhoneCallLimitedDate,
                LawPhoneCallReadLog = model.LawPhoneCallReadLog,
                PhoneCallCreatorID = model.PhoneCallCreatorID,
            });

            var sql1 = @"update law_content set law_phone_call2_desc = @LawPhoneCall2Desc ,law_phone_call2_date=@LawPhoneCall2Date where law_id=@LawId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql1, new
            { LawId = content.LawId, LawPhoneCall2Desc = content.LawPhoneCall2Desc, LawPhoneCall2Date = content.LawPhoneCall2Date });
        }

        /// <summary>
        /// 新增結案狀況其他欄位說明記錄log
        /// </summary>
        public void InsertLawCloseOtherLog(LawCloseOtherLog model)
        {
            string sql = @"INSERT INTO law_close_other_log(law_id,law_note_no,law_close_other_desc,create_date,CloseOtherLogCreatorID)
                           VALUES (@LawId, @LawNoteNo, @LawCloseOtherDesc,@CreateDate,@CloseOtherLogCreatorID);";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
                LawCloseOtherDesc = model.LawCloseOtherDesc,
                CreateDate = model.CreateDate,
                CloseOtherLogCreatorID = model.CloseOtherLogCreatorID,
            });
        }

        /// <summary>
        /// 更新存證信函
        /// </summary>
        public void UpdateLawEvidenceDesc(LawEvidenceDesc model)
        {
            var sql = @"update law_evidence_desc set law_evidence_desc = @LawEvidencedesc,EvidenceDescCreatorID = @EvidenceDescCreatorID,create_date = GETDATE() 
                where law_evidence_id = @LawEvidenceId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { LawEvidenceId = model.LawEvidenceId, LawEvidencedesc = model.LawEvidencedesc, EvidenceDescCreatorID = model.EvidenceDescCreatorID });
        }

        /// <summary>
        /// 更新訴訟程序
        /// </summary>
        public void UpdateLawLitigationProgress(LawLitigationProgress model, LawContent content)
        {
            var sql = @"update law_litigation_progress set law_litigation_progress = @LawLitigationprogress ,LawLitigationProgressCreatorID = @LawLitigationProgressCreatorID,create_date = GETDATE() 
                where law_litigation_id = @LawLitigationId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { LawLitigationId = model.LawLitigationId, LawLitigationprogress = model.LawLitigationprogress, LawLitigationProgressCreatorID = model.LawLitigationProgressCreatorID });

            var sql1 = @"update law_content set law_content_lastchange_date = @LawContentLastchangeDate where law_id= @LawId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql1, new
            { LawId = content.LawId, LawContentLastchangeDate = content.LawContentLastchangeDate });
        }

        /// <summary>
        /// 更新執行程序
        /// </summary>
        public void UpdateLawDoProgress(LawDoProgress model, LawContent content)
        {
            var sql = @"update law_do_progress set law_do_progress = @LawDoprogress ,LawDoProgressCreatorID = @LawDoProgressCreatorID,create_date=GETDATE() 
            where law_do_id = @LawDoId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { LawDoId = model.LawDoId, LawDoprogress = model.LawDoprogress, LawDoProgressCreatorID = model.LawDoProgressCreatorID });

            var sql1 = @"update law_content set law_content_lastchange_date = @LawContentLastchangeDate where law_id = @LawId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql1, new
            { LawId = content.LawId, LawContentLastchangeDate = content.LawContentLastchangeDate });
        }

        /// <summary>
        /// 更新其他
        /// </summary>
        public void UpdateLawOtherDesc(LawOtherDesc model)
        {
            var sql = @"update law_other_desc set law_other_desc = @LawOtherdesc,OtherDescCreatorID = @OtherDescCreatorID,create_date=GETDATE() where law_other_id =@LawOtherId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { LawOtherdesc = model.LawOtherdesc, OtherDescCreatorID = model.OtherDescCreatorID, LawOtherId = model.LawOtherId });
        }

        /// <summary>
        /// 更新備註
        /// </summary>
        public void UpdateLawDescDesc(LawDescDesc model)
        {
            var sql = @"update law_desc_desc set law_desc_desc = @LawDescdesc,LawDescDescCreatorID = @LawDescDescCreatorID,create_date=GETDATE() where law_desc_id =@LawDescId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { LawDescdesc = model.LawDescdesc, LawDescDescCreatorID = model.LawDescDescCreatorID, LawDescId = model.LawDescId });
        }

        /// <summary>
        /// 更新分案日期
        /// </summary>
        public void UpdateCaseDate(LawDoUnitLog model)
        {
            var sql = @"update law_do_unit_log set case_date= @CaseDate where law_dounit_log_id= @LawDoUnitId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { CaseDate = model.CaseDate, LawDoUnitId = model.LawDoUnitId });
        }

        /// <summary>
        /// 更新結案原因和卷宗
        /// </summary>
        public void UpdateLawContentReasonfile(LawContent model)
        {
            var sql = @"update law_content set law_due_reason = @LawDueReason,law_file_no = @LawFileNo where law_id=@LawId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { LawFileNo = model.LawFileNo, LawDueReason = model.LawDueReason, LawId = model.LawId });
        }

        /// <summary>
        /// 更新未結案狀態
        /// </summary>
        public void UpdateLawContentNotCloseType(LawContent model)
        {
            var sql = @"update law_content set law_status_type=@LawStatusType,law_not_close_type_name=@LawNotCloseTypeName where  law_id = @LawId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { LawStatusType = model.LawStatusType, LawNotCloseTypeName = model.LawNotCloseTypeName, LawId = model.LawId });
        }

        /// <summary>
        /// 更新為案件取消
        /// </summary>
        public void UpdateLawContentCancelType(LawContent model)
        {
            var sql = @"update law_content set law_content_cancel_type = @LawContentCancelType where law_id=@LawId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { LawContentCancelType = model.LawContentCancelType, LawId = model.LawId });
        }

        /// <summary>
        /// 更新結案狀態
        /// </summary>
        public void UpdateLawContentCloseType(LawContent model)
        {
            var sql = @"update law_content set law_status_type=@LawStatusType,law_close_type=@LawCloseType,law_close_date=@LawCloseDate,LawContentCloserID = @LawContentCloserID
                ,law_closer_name=@LawCloserName,law_content_lastchange_date=@LawContentLastchangeDate where law_id=@LawId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawStatusType = model.LawStatusType,
                LawCloseType = model.LawCloseType,
                LawCloseDate = model.LawCloseDate,
                LawContentCloserID = model.LawContentCloserID,
                LawCloserName = model.LawCloserName,
                LawContentLastchangeDate = model.LawContentLastchangeDate,
                LawId = model.LawId
            });
        }

        /// <summary>
        /// 用ID刪除單筆存證信函資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawEvidenceDescByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawEvidenceDesc>(new { LawEvidenceId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }

        /// <summary>
        /// 用ID刪除單筆訴訟程序資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawLitigationProgressByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawLitigationProgress>(new { LawLitigationId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }

        /// <summary>
        /// 用ID刪除單筆執行程序資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawDoProgressByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawDoProgress>(new { LawDoId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }

        /// <summary>
        /// 用ID刪除單筆其他資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawOtherDescByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawOtherDesc>(new { LawOtherId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }

        /// <summary>
        /// 用ID刪除單筆備註資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawDescDescByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawDescDesc>(new { LawDescId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }
        #endregion

        #region 結欠金額表
        /// <summary>
        /// 結欠明細
        /// </summary>
        public List<LawContent> GetDueMoneyDetail(string AgentID, string NoteNo)
        {
            List<LawContent> result = new List<LawContent>();
            string sql = @"select * from law_content where left(law_due_agentid,10) = @AgentID and law_note_no = @NoteNo";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { AgentID = AgentID, NoteNo = NoteNo }).ToList();
            return result;
        }

        /// <summary>
        /// 查詢主檔
        /// </summary>
        public LawContent GetLawContentByLawId(int LawId)
        {
            LawContent result = new LawContent();
            string sql = @"select * from law_content where law_id= @LawId";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 查詢律師服務費 
        /// </summary>
        public List<LawLawyerReward> GetLawLawyerReward(int LawId, string LawDueAgentId)
        {
            List<LawLawyerReward> result = new List<LawLawyerReward>();
            string sql = @"select * from law_lawyer_reward where law_due_agentid=@LawDueAgentId and law_id= @LawId";

            result = DbHelper.Query<LawLawyerReward>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId, LawDueAgentId = LawDueAgentId }).ToList();
            return result;
        }

        /// <summary>
        /// 律師服務費編輯頁面
        /// </summary>
        public LawLawyerReward GetLawLawyerRewardById(int LawLawyerPayId)
        {
            LawLawyerReward result = new LawLawyerReward();
            string sql = @"select * from law_lawyer_reward where law_lawyer_pay_id= @LawLawyerPayId";

            result = DbHelper.Query<LawLawyerReward>(H2ORepository.ConnectionStringName, sql, new { LawLawyerPayId = LawLawyerPayId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 查詢0利率
        /// </summary>
        public LawInterestRates GetLawInterestRatesByZero()
        {
            LawInterestRates result = new LawInterestRates();
            string sql = @"select * from law_interest_rates where interest_rates=0.00";

            result = DbHelper.Query<LawInterestRates>(H2ORepository.ConnectionStringName, sql, new { }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 清償明細
        /// </summary>
        public List<LawRepaymentList> GetLawRepaymentListByLawId(int LawId)
        {
            List<LawRepaymentList> result = new List<LawRepaymentList>();
            string sql = @"select * from law_repayment_list where law_id = @LawId order by law_repayment_date";

            result = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId }).ToList();
            return result;
        }

        /// <summary>
        /// 確認清償明細可否新增
        /// </summary>
        public LawContent ChkIRInsert(int LawId, string LawDueAgentId)
        {
            LawContent result = new LawContent();
            string sql = @"select sum(law_total_money) as law_total_money from law_content where law_due_agentid=@LawDueAgentId and law_id=@LawId";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId, LawDueAgentId = LawDueAgentId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 累計清償
        /// </summary>
        public LawRepaymentList GetTotalLawRepaymentMoney(string LawDueAgentId, IEnumerable<int> LawRepaymentId)
        {
            LawRepaymentList result = new LawRepaymentList();
            string sql = @"select sum(a.law_repayment_money) as law_repayment_money from (select sum(law_repayment_money) as law_repayment_money from law_repayment_list where law_due_agentid=@LawDueAgentId and law_repayment_id in @LawRepaymentId 
                union select sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_due_agentid= @LawDueAgentId and law_repayment_id in @LawRepaymentId ) a ";

            result = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql, new { LawDueAgentId = LawDueAgentId, LawRepaymentId = LawRepaymentId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 剩餘結欠金額
        /// </summary>
        public LawContent GetRemainLawRepaymentMoney(string LawDueAgentId, int LawId)
        {
            LawContent result = new LawContent();
            string sql = @"select sum(law_total_money) as law_total_money from law_content where law_due_agentid=@LawDueAgentId and law_id=@LawId";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { LawDueAgentId = LawDueAgentId, LawId = LawId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 新增結欠明細Log
        /// </summary>
        public void InsertLawInterestLog(LawInterestLog model)
        {
            string sql = @"INSERT INTO law_interest_log(law_id,law_note_no,law_due_money,law_interest_sdate,law_interest_edate,law_interest_days,law_interest_rates_id,law_interest_money,law_total_money,create_date,InterestLogCreateID)
            VALUES (@LawId, @LawNoteNo, @LawDueMoney,@LawInterestSdate,@LawInterestEdate,@LawInterestDays,@LawInterestRatesId,@LawInterestMoney,@LawTotalMoney, @CreateDate, @InterestLogCreateID);";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                CreateDate = model.CreateDate,
                InterestLogCreateID = model.InterestLogCreateID,
                LawDueMoney = model.LawDueMoney,
                LawInterestSdate = model.LawInterestSdate,
                LawInterestEdate = model.LawInterestEdate,
                LawInterestDays = model.LawInterestDays,
                LawInterestRatesId = model.LawInterestRatesId,
                LawInterestMoney = model.LawInterestMoney,
                LawTotalMoney = model.LawTotalMoney,
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
            });
        }

        /// <summary>
        /// 新增清償明細
        /// </summary>
        public void InsertLawRepaymentList(LawRepaymentList model)
        {
            string sql = @"INSERT INTO law_repayment_list(law_id,law_note_no,law_repayment_date,law_due_agentid,law_repayment_money,law_repayment_capital,law_repayment_interest,law_repayment_court,law_repayment_other,law_comm_deduction,create_date,RepaymentListCreatorID)
            VALUES (@LawId, @LawNoteNo, @LawRepaymentDate,@LawDueAgentId,@LawRepaymentMoney,@LawRepaymentCapital,@LawRepaymentInterest,@LawRepaymentCourt,@LawRepaymentOther,@LawCommDeduction, @CreateDate, @RepaymentListCreatorID);";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                CreateDate = model.CreateDate,
                RepaymentListCreatorID = model.RepaymentListCreatorID,
                LawDueAgentId = model.LawDueAgentId,
                LawRepaymentDate = model.LawRepaymentDate,
                LawRepaymentMoney = model.LawRepaymentMoney,
                LawRepaymentCapital = model.LawRepaymentCapital,
                LawRepaymentInterest = model.LawRepaymentInterest,
                LawRepaymentCourt = model.LawRepaymentCourt,
                LawRepaymentOther = model.LawRepaymentOther,
                LawCommDeduction = model.LawCommDeduction,
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
            });
        }

        /// <summary>
        /// 新增清償明細Log
        /// </summary>
        public void InsertLawRepaymentListLog(LawRepaymentListLog model)
        {
            string sql = @"INSERT INTO law_repayment_list_log(law_id,law_note_no,law_repayment_date,law_due_agentid,law_repayment_money,law_repayment_capital,law_repayment_interest,law_repayment_court,law_repayment_other,law_comm_deduction,create_date,RepaymentListLogCreatorID,RepaymentListLogDelID,del_name,del_time,law_repayment_id)
            VALUES (@LawId, @LawNoteNo, @LawRepaymentDate,@LawDueAgentId,@LawRepaymentMoney,@LawRepaymentCapital,@LawRepaymentInterest,@LawRepaymentCourt,@LawRepaymentOther,@LawCommDeduction, @CreateDate, @RepaymentListLogCreatorID,@RepaymentListLogDelID,@DelName,@DelTime,@LawRepaymentId);";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                DelTime = model.DelTime,
                RepaymentListLogDelID = model.RepaymentListLogDelID,
                DelName = model.DelName,
                CreateDate = model.CreateDate,
                RepaymentListLogCreatorID = model.RepaymentListLogCreatorID,
                LawDueAgentId = model.LawDueAgentId,
                LawRepaymentDate = model.LawRepaymentDate,
                LawRepaymentMoney = model.LawRepaymentMoney,
                LawRepaymentCapital = model.LawRepaymentCapital,
                LawRepaymentInterest = model.LawRepaymentInterest,
                LawRepaymentCourt = model.LawRepaymentCourt,
                LawRepaymentOther = model.LawRepaymentOther,
                LawCommDeduction = model.LawCommDeduction,
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
                LawRepaymentId = model.LawRepaymentId
            });
        }

        /// <summary>
        /// 新增律師服務費
        /// </summary>
        public void InsertLawLawyerReward(LawLawyerReward model)
        {

            string sql = @"INSERT INTO law_lawyer_reward(law_id,law_note_no,law_due_agentid,law_repayment_money,law_reward_pay_year,law_reward_pay_month,law_fees,law_rates,law_service_reward,create_date,LawyerRewardCreateID)
            VALUES (@LawId, @LawNoteNo,@LawDueAgentId,@LawRepaymentMoney,@LawRewardPayYear,@LawRewardPayMonth,@LawFees,@LawRates,@LawServiceReward,@CreateDate,@LawyerRewardCreateID);";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawRepaymentMoney = model.LawRepaymentMoney,
                LawRewardPayYear = model.LawRewardPayYear,
                LawRewardPayMonth = model.LawRewardPayMonth,
                CreateDate = model.CreateDate,
                LawFees = model.LawFees,
                LawDueAgentId = model.LawDueAgentId,
                LawRates = model.LawRates,
                LawServiceReward = model.LawServiceReward,
                LawyerRewardCreateID = model.LawyerRewardCreateID,
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
            });

        }

        /// <summary>
        /// 更新為案件取消
        /// </summary>
        public void UpdateLawContentDueMoney(LawContent model)
        {
            var sql = @"update law_content set law_due_money=@LawDueMoney,law_interest_sdate=@LawInterestSdate,law_interest_edate=@LawInterestEdate,law_interest_rates_id=@LawInterestRatesId,law_interest_money=@LawInterestMoney,law_interest_days=@LawInterestDays,law_total_money=@LawTotalMoney,law_status_type=1 
                where law_id= @LawId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawDueMoney = model.LawDueMoney,
                LawInterestSdate = model.LawInterestSdate,
                LawInterestEdate = model.LawInterestEdate,
                LawInterestRatesId = model.LawInterestRatesId,
                LawInterestMoney = model.LawInterestMoney,
                LawInterestDays = model.LawInterestDays,
                LawTotalMoney = model.LawTotalMoney,
                LawId = model.LawId
            });
        }

        /// <summary>
        /// 更新最後編輯日
        /// </summary>
        public void UpdateLawContentLawContentLastchangeDate(LawContent model)
        {
            var sql = @"update law_content set law_content_lastchange_date=@LawContentLastchangeDate
                where law_id= @LawId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawContentLastchangeDate = model.LawContentLastchangeDate,
                LawId = model.LawId
            });
        }

        /// <summary>
        /// 更新律師服務費
        /// </summary>
        public void UpdateLawLawyerReward(LawLawyerReward model)
        {
            var sql = @"update law_lawyer_reward set law_reward_pay_year=@LawRewardPayYear,law_reward_pay_month=@LawRewardPayMonth,law_repayment_money=@LawRepaymentMoney,law_fees=@LawFees,law_rates=@LawRates,law_service_reward=@LawServiceReward where law_lawyer_pay_id=@LawLawyerPayId ";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                LawLawyerPayId = model.LawLawyerPayId,
                LawServiceReward = model.LawServiceReward,
                LawRewardPayYear = model.LawRewardPayYear,
                LawRewardPayMonth = model.LawRewardPayMonth,
                LawRepaymentMoney = model.LawRepaymentMoney,
                LawFees = model.LawFees,
                LawRates = model.LawRates,
                LawId = model.LawId
            });
        }

        /// <summary>
        /// 用ID刪除單筆資料
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public bool DeleteLawRepaymentListByID(string ID)
        {
            var result = false;
            if (!string.IsNullOrEmpty(ID))
            {
                using (TransactionScope ts = new TransactionScope())
                {
                    H2ORepository.Delete<LawRepaymentList>(new { LawRepaymentId = ID });
                    H2ORepository.Delete<LawOtherDesc>(new { LawRepaymentId = ID });
                    result = true;
                    ts.Complete();
                }
            }
            return result;
        }
        #endregion

        #region 存證信函
        /// <summary>
        /// 存證信函
        /// </summary>
        public LawEvidenceDetail GetLawEvidenceById(string EvidAgentId, string LawNoteNo)
        {
            LawEvidenceDetail result = new LawEvidenceDetail();
            string sql = @"select * from law_evidence where evid_agent_id=@EvidAgentId and law_note_no=@LawNoteNo";

            result = DbHelper.Query<LawEvidenceDetail>(H2ORepository.ConnectionStringName, sql, new { EvidAgentId = EvidAgentId, LawNoteNo = LawNoteNo }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 存證信函個人資料
        /// </summary>
        public LawEvidenceDetail GetLawAgentDetailById(string AgentID, string NoteNo)
        {
            LawEvidenceDetail result = new LawEvidenceDetail();
            string sql = @"select * from law_agent_content a left join law_agent_data b on left(a.agent_code,10)=b.agent_code where a.agent_code = @AgentID and a.law_note_no = @NoteNo";

            result = DbHelper.Query<LawEvidenceDetail>(H2ORepository.ConnectionStringName, sql, new { AgentID = AgentID, NoteNo = NoteNo }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 新增存證信函
        /// </summary>
        public void InsertLawEvidence(LawEvidence model)
        {
            string sql = @"INSERT INTO law_evidence(evid_no,law_id,law_note_no,evid_sender,evid_sender_add,evid_agent_id,evid_agent_name,evid_agent_add,evid_agent_add1,evid_reason,evid_money,evid_money_num,evid_account,evid_user,evid_phone,EvidenceCreatorID,evid_create_name,evid_create_date)
            VALUES (@EvidNo,@LawId, @LawNoteNo, @EvidSender,@EvidSenderAdd,@EvidAgentId,@EvidAgentName,@EvidAgentAdd,@EvidAgentAdd1,@EvidReason, @EvidMoney, @EvidMoneyNum,@EvidAccount,@EvidUser,@EvidPhone,@EvidenceCreatorID,@EvidCreateName,@EvidCreateDate);";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                EvidNo = model.EvidNo,
                EvidSender = model.EvidSender,
                EvidSenderAdd = model.EvidSenderAdd,
                EvidAgentId = model.EvidAgentId,
                EvidAgentName = model.EvidAgentName,
                EvidAgentAdd = model.EvidAgentAdd,
                EvidAgentAdd1 = model.EvidAgentAdd1,
                EvidReason = model.EvidReason,
                EvidMoney = model.EvidMoney,
                EvidMoneyNum = model.EvidMoneyNum,
                EvidAccount = model.EvidAccount,
                EvidUser = model.EvidUser,
                EvidPhone = model.EvidPhone,
                EvidenceCreatorID = model.EvidenceCreatorID,
                EvidCreateName = model.EvidCreateName,
                EvidCreateDate = model.EvidCreateDate,
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
            });
        }

        /// <summary>
        /// 新增存證信函Log
        /// </summary>
        public void InsertLawEvidenceLog(LawEvidenceLog model)
        {
            string sql = @"INSERT INTO law_evidence_log(evid_id,evid_no,law_id,law_note_no,evid_sender,evid_sender_add,evid_agent_id,evid_agent_name,evid_agent_add,evid_agent_add1,evid_reason,evid_money,evid_money_num,evid_account,evid_user,evid_phone,EvidenceLogCreatorID,evid_create_name,evid_create_date)
            VALUES (@EvidId,@EvidNo,@LawId, @LawNoteNo, @EvidSender,@EvidSenderAdd,@EvidAgentId,@EvidAgentName,@EvidAgentAdd,@EvidAgentAdd1,@EvidReason, @EvidMoney, @EvidMoneyNum,@EvidAccount,@EvidUser,@EvidPhone,@EvidenceLogCreatorID,@EvidCreateName,@EvidCreateDate);";

            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                EvidId = model.EvidId,
                EvidNo = model.EvidNo,
                EvidSender = model.EvidSender,
                EvidSenderAdd = model.EvidSenderAdd,
                EvidAgentId = model.EvidAgentId,
                EvidAgentName = model.EvidAgentName,
                EvidAgentAdd = model.EvidAgentAdd,
                EvidAgentAdd1 = model.EvidAgentAdd1,
                EvidReason = model.EvidReason,
                EvidMoney = model.EvidMoney,
                EvidMoneyNum = model.EvidMoneyNum,
                EvidAccount = model.EvidAccount,
                EvidUser = model.EvidUser,
                EvidPhone = model.EvidPhone,
                EvidenceLogCreatorID = model.EvidenceLogCreatorID,
                EvidCreateName = model.EvidCreateName,
                EvidCreateDate = model.EvidCreateDate,
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
            });
        }

        /// <summary>
        /// 更新存證信函
        /// </summary>
        public void UpdateLawEvidence(LawEvidence model)
        {
            var sql = @"update law_evidence set evid_no=@EvidNo,evid_sender=@EvidSender,evid_sender_add=@EvidSenderAdd,evid_agent_name=@EvidAgentName,evid_agent_add=@EvidAgentAdd
                ,evid_agent_add1=@EvidAgentAdd1,evid_reason=@EvidReason,evid_money=@EvidMoney,evid_money_num=@EvidMoneyNum,evid_account=@EvidAccount,evid_user=@EvidUser,evid_phone=@EvidPhone
                ,evid_create_name=@EvidCreateName,evid_create_date=@EvidCreateDate,EvidenceCreatorID = @EvidenceCreatorID where evid_id= @EvidId";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            {
                EvidId = model.EvidId,
                EvidNo = model.EvidNo,
                EvidSender = model.EvidSender,
                EvidSenderAdd = model.EvidSenderAdd,
                EvidAgentId = model.EvidAgentId,
                EvidAgentName = model.EvidAgentName,
                EvidAgentAdd = model.EvidAgentAdd,
                EvidAgentAdd1 = model.EvidAgentAdd1,
                EvidReason = model.EvidReason,
                EvidMoney = model.EvidMoney,
                EvidMoneyNum = model.EvidMoneyNum,
                EvidAccount = model.EvidAccount,
                EvidUser = model.EvidUser,
                EvidPhone = model.EvidPhone,
                EvidenceCreatorID = model.EvidenceCreatorID,
                EvidCreateName = model.EvidCreateName,
                EvidCreateDate = model.EvidCreateDate,
                LawId = model.LawId,
                LawNoteNo = model.LawNoteNo,
            });
        }
        #endregion

        #region 照會單
        /// <summary>
        /// 照會單
        /// </summary>
        public List<LawNote> GetLawNoteByLawNoteNo(string LawNoteNo)
        {
            List<LawNote> result = new List<LawNote>();
            string sql = @"select * from law_note where law_note_km in (select max(law_note_km) from law_note where law_note_no=@LawNoteNo)";

            result = DbHelper.Query<LawNote>(H2ORepository.ConnectionStringName, sql, new { LawNoteNo = LawNoteNo }).ToList();
            return result;
        }

        /// <summary>
        /// 照會單作業-已發送紀錄
        /// </summary>
        public List<MessageTo> GetMessageToByLawNoteNo(string LawNoteNo)
        {
            List<MessageTo> result = new List<MessageTo>();
            string sql = @"select MSGOBJCreateTime, MSGOBJName, MSGOBJReaderIP from MessageTo where MSGOBJNote=@LawNoteNo";
            sql += " union select mo_createtime, mo_orgname, mo_readip from k_message_to  where mg_no = @LawNoteNo";
            result = DbHelper.Query<MessageTo>(H2ORepository.ConnectionStringName, sql, new { LawNoteNo = LawNoteNo }).ToList();
            return result;
        }

        #endregion
        #endregion

        #region 查詢作業
        /// <summary>
        /// 查詢作業
        /// </summary>
        public List<LawSearchDetail> GetLawContentBySearch(LawSearchDetail model)
        {
            List<LawSearchDetail> result = new List<LawSearchDetail>();
            string sql = @"select 
            law_due_agentid,
            law_due_name,
            law_do_unit_name,
            law_due_money,
            law_close_type,
            law_status_type,
            law_year + '/' + law_month as production_ym,
            law_id,
            law_note_no,
            law_not_close_type_name 
            from law_content where 2>1";

            if (!string.IsNullOrWhiteSpace(model.LawDueAgentId))
            {
                sql += " and left(law_due_agentid,10)= @LawDueAgentId";
            }
            if (!string.IsNullOrWhiteSpace(model.LawDueName))
            {
                sql += " and law_due_name like @LawDueName";
            }
            if (!string.IsNullOrWhiteSpace(model.LawYear))
            {
                sql += " and law_year= @LawYear";
            }
            if (!string.IsNullOrWhiteSpace(model.LawMonth))
            {
                sql += " and law_month= @LawMonth";
            }
            if (model.LawDoUnitId != 0)
            {
                sql += " and law_do_unit_id= @LawDoUnitId";
            }
            if (model.LawDueMoney != 0)
            {
                sql += " and law_due_money= @LawDueMoney";
            }
            if (!string.IsNullOrWhiteSpace(model.LawCloseType))
            {
                sql += " and law_close_type like @LawCloseType";
            }
            if (model.LawStatusType != 0)
            {
                sql += " and law_status_type= @LawStatusType";
            }

            sql += " order by law_note_no";

            result = DbHelper.Query<LawSearchDetail>(H2ORepository.ConnectionStringName, sql, new
            {
                LawDueAgentId = model.LawDueAgentId,
                LawDueName = "%" + model.LawDueName + "%",
                LawYear = model.LawYear,
                LawMonth = model.LawMonth,
                LawDoUnitId = model.LawDoUnitId,
                LawDueMoney = model.LawDueMoney,
                LawCloseType = "%" + model.LawCloseType + "%",
                LawStatusType = model.LawStatusType,
            }).ToList();
            return result;
        }

        /// <summary>
        /// 取得啟用狀態結案狀態設定
        /// </summary>
        public List<LawCloseType> GetLawCloseTypeByLawCloseType(string LawCloseType)
        {
            List<LawCloseType> result = new List<LawCloseType>();
            string sql = @"select * from law_close_type where status_type=1 and close_type_id in (" + LawCloseType + ")";

            result = DbHelper.Query<LawCloseType>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 確認是否為管理人員
        /// </summary>
        public LawSys GetLawSysBySysMemberID(string SysMemberID)
        {
            LawSys result = new LawSys();
            string sql = @"select * from law_sys where SysMemberID = @SysMemberID";

            result = DbHelper.Query<LawSys>(H2ORepository.ConnectionStringName, sql, new { SysMemberID = SysMemberID }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 確認VMSM
        /// </summary>
        public OrgVm GetOrfVm(string AccID)
        {
            OrgVm result = new OrgVm();
            string sql = @"select distinct vm_code,vm_name,vm_leader_id,virtual_vm from org_vm where vm_leader_id= @AccID";
            result = DbHelper.Query<OrgVm>(MISRepository.ConnectionStringName, sql, new { AccID = AccID }).FirstOrDefault();
            if (result != null)
            {
                result.vm_flag = 1;
                switch (result.virtualvm)
                {
                    case "Y":
                        result.vsm_flag = 1;
                        string sql1 = @"select * from org_sm where vm_code in(select vm_code from org_vm where virtual_vm='Y') and vm_code=@vmcode";
                        List<OrgVm> orgsm = DbHelper.Query<OrgVm>(MISRepository.ConnectionStringName, sql1, new { vmcode = result.vmcode }).ToList();
                        if (orgsm.Count != 0)
                        {
                            for (int i = 0; i < orgsm.Count; i++)
                            {
                                result.smstr = i == 0 ? "'" + orgsm[i].smcode + "'" :  result.smstr + "," + "'" + orgsm[i].smcode + "'";
                            }
                        }

                        break;
                }
            }

            return result;
        }

        /// <summary>
        /// 未結案件明細列表
        /// </summary>
        public List<LawSearchDetail> GetLawContentByNotClose(LawSearchDetail model, OrgVm orgVm, LawSys lawSys)
        {
            List<LawSearchDetail> result = new List<LawSearchDetail>();
            string sql = string.Empty;
            if (lawSys != null)
            {
                sql = @"select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym ,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG
                ,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no	 where (law_content_cancel_type is null or law_content_cancel_type<>1) and law_status_type<>2";
            }
            else
            {
                if (orgVm.vm_flag == 1)
                {
                    if (orgVm.vsm_flag == 0)
                    {
                        sql = @"select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG
			            ,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where b.vm_code=@vmcode and (law_content_cancel_type is null or law_content_cancel_type<>1) and law_status_type<>2";
                    }
                    else
                    {
                        sql = @"select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG
			            ,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where b.sm_code in (" + orgVm.smstr + ") and (law_content_cancel_type is null or law_content_cancel_type<>1) and law_status_type<>2";
                    }
                }
                else
                {
                    sql = @"select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG
			        ,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where administrat_id=@vmleaderid and (law_content_cancel_type is null or law_content_cancel_type<>1) and law_status_type<>2";
                }
            }

            if (!string.IsNullOrWhiteSpace(model.LawDueAgentId))
            {
                sql += " and left(law_due_agentid,10)= @LawDueAgentId";
            }
            if (!string.IsNullOrWhiteSpace(model.LawDueName))
            {
                sql += " and law_due_name like @LawDueName";
            }
            if (model.LawDoUnitId != 0)
            {
                sql += " and law_do_unit_id= @LawDoUnitId";
            }
            if (model.LawDueMoney != 0)
            {
                sql += " and law_due_money= @LawDueMoney";
            }
            if (!string.IsNullOrWhiteSpace(model.LawCloseType))
            {
                sql += " and law_close_type like @LawCloseType";
            }
            if (!string.IsNullOrWhiteSpace(model.LawNoteNo))
            {
                sql += " and a.law_note_no like @LawNoteNo";
            }
            if (!string.IsNullOrWhiteSpace(model.VmName))
            {
                sql += " and vm_name like @VmName";
            }
            if (!string.IsNullOrWhiteSpace(model.SmName))
            {
                sql += " and sm_name like @SmName";
            }
            if (!string.IsNullOrWhiteSpace(model.CenterName))
            {
                sql += " and center_name like @CenterName";
            }
            if (!string.IsNullOrWhiteSpace(model.WcCenterName))
            {
                sql += " and wc_center_name like @WcCenterName";
            }

            sql += " order by a.law_note_no";

            result = DbHelper.Query<LawSearchDetail>(H2ORepository.ConnectionStringName, sql, new
            {
                vmleaderid = orgVm.vmleaderid.Substring(0, 10),
                vmcode = orgVm.vmcode,
                LawDueAgentId = model.LawDueAgentId,
                LawDueName = "%" + model.LawDueName + "%",
                LawNoteNo = "%" + model.LawNoteNo + "%",
                VmName = "%" + model.VmName + "%",
                LawDoUnitId = model.LawDoUnitId,
                LawDueMoney = model.LawDueMoney,
                LawCloseType = "%" + model.LawCloseType + "%",
                SmName = "%" + model.SmName + "%",
                CenterName = "%" + model.CenterName + "%",
                WcCenterName = "%" + model.WcCenterName + "%",
            }).ToList();
            return result;
        }

        /// <summary>
        /// 結案件明細列表
        /// </summary>
        public List<LawSearchDetail> GetLawContentByClose(LawSearchDetail model, OrgVm orgVm, LawSys lawSys)
        {
            List<LawSearchDetail> result = new List<LawSearchDetail>();
            string sql = string.Empty;
            if (lawSys != null)
            {
                sql = @"select *,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG
                ,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where (law_status_type=2 or law_content_cancel_type is not null or law_content_cancel_type=1)";
            }
            else
            {
                if (orgVm.vm_flag == 1)
                {
                    if (orgVm.vsm_flag == 0)
                    {
                        sql = @"select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG
			            ,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where b.vm_code=@vmcode and  (law_content_cancel_type is null or law_content_cancel_type<>1)";
                    }
                    else
                    {
                        sql = @"select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG
			            ,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where b.sm_code in (" + orgVm.smstr + ") and  (law_content_cancel_type is null or law_content_cancel_type<>1)";
                    }
                }
                else
                {
                    sql = @"select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG
			        ,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where administrat_id=@vmleaderid and (law_content_cancel_type is null or law_content_cancel_type<>1)";
                }
            }

            if (!string.IsNullOrWhiteSpace(model.LawDueAgentId))
            {
                sql += " and left(law_due_agentid,10)= @LawDueAgentId";
            }
            if (!string.IsNullOrWhiteSpace(model.LawDueName))
            {
                sql += " and law_due_name like @LawDueName";
            }
            if (model.LawDoUnitId != 0)
            {
                sql += " and law_do_unit_id= @LawDoUnitId";
            }
            if (model.LawDueMoney != 0)
            {
                sql += " and law_due_money= @LawDueMoney";
            }
            if (!string.IsNullOrWhiteSpace(model.LawCloseType))
            {
                sql += " and law_close_type like @LawCloseType";
            }
            if (!string.IsNullOrWhiteSpace(model.LawNoteNo))
            {
                sql += " and a.law_note_no like @LawNoteNo";
            }
            if (!string.IsNullOrWhiteSpace(model.VmName))
            {
                sql += " and vm_name like @VmName";
            }
            if (!string.IsNullOrWhiteSpace(model.SmName))
            {
                sql += " and sm_name like @SmName";
            }
            if (!string.IsNullOrWhiteSpace(model.CenterName))
            {
                sql += " and center_name like @CenterName";
            }
            if (!string.IsNullOrWhiteSpace(model.WcCenterName))
            {
                sql += " and wc_center_name like @WcCenterName";
            }

            sql += " order by a.law_note_no";

            result = DbHelper.Query<LawSearchDetail>(H2ORepository.ConnectionStringName, sql, new
            {
                vmleaderid = orgVm.vmleaderid.Substring(0, 10),
                vmcode = orgVm.vmcode,
                LawDueAgentId = model.LawDueAgentId,
                LawDueName = "%" + model.LawDueName + "%",
                LawNoteNo = "%" + model.LawNoteNo + "%",
                VmName = "%" + model.VmName + "%",
                LawDoUnitId = model.LawDoUnitId,
                LawDueMoney = model.LawDueMoney,
                LawCloseType = "%" + model.LawCloseType + "%",
                SmName = "%" + model.SmName + "%",
                CenterName = "%" + model.CenterName + "%",
                WcCenterName = "%" + model.WcCenterName + "%",
            }).ToList();
            return result;
        }

        /// <summary>
        /// 追回金額
        /// </summary>
        public LawRepaymentList GetLawRepaymentListByAgID(string LawDueAgentId, string LawId)
        {
            LawRepaymentList result = new LawRepaymentList();
            string sql = @"select sum(law_repayment_money)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where left(law_due_agentid,10)=@LawDueAgentId and law_id=@LawId";

            result = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql, new { LawDueAgentId = LawDueAgentId, LawId = LawId }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 存證信函
        /// </summary>
        public LawEvidenceDesc GetLawEvidenceDescByLawId(string LawId)
        {
            LawEvidenceDesc result = new LawEvidenceDesc();
            string sql = @"select * from law_evidence_desc where law_evidence_id in(select max(law_evidence_id) as law_evidence_id from law_evidence_desc where law_id = @LawId)";

            result = DbHelper.Query<LawEvidenceDesc>(H2ORepository.ConnectionStringName, sql, new { LawId = LawId }).FirstOrDefault();
            return result;
        }

        #region 報表

        public byte[] GetSearchReportList(LawSearchDetail model)
        {
            string sql = @"select a.*,b.law_agent_content_id,production_ym,sequence,vm_code,vm_name,sm_code,sm_name,wc_center,wc_center_name,center_code,center_name,administrat_id,admin_name,admin_level,agent_code,name,ag_status_code,ag_level,ag_level_name,record_date,register_date,ag_status_date,super_account,create_date,(select isnull(sum(law_repayment_capital),0) from law_repayment_list where law_note_no=a.law_note_no) as law_repayment_capital from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where 2>1 ";
            if (!string.IsNullOrWhiteSpace(model.LawDueAgentId))
            {
                sql += " and left(law_due_agentid,10)= @LawDueAgentId";
            }
            if (!string.IsNullOrWhiteSpace(model.LawDueName))
            {
                sql += " and law_due_name like @LawDueName";
            }
            if (!string.IsNullOrWhiteSpace(model.LawYear))
            {
                sql += " and law_year= @LawYear";
            }
            if (!string.IsNullOrWhiteSpace(model.LawMonth))
            {
                sql += " and law_month= @LawMonth";
            }
            if (model.LawDoUnitId != 0)
            {
                sql += " and law_do_unit_id= @LawDoUnitId";
            }
            if (model.LawDueMoney != 0)
            {
                sql += " and law_due_money= @LawDueMoney";
            }
            if (!string.IsNullOrWhiteSpace(model.LawCloseType))
            {
                sql += " and law_close_type like @LawCloseType";
            }
            if (model.LawStatusType != 0)
            {
                sql += " and law_status_type= @LawStatusType";
            }
            sql += " order by a.law_note_no";
            List<LawSearchDetail> lawSearchDetails = new List<LawSearchDetail>();
            List<LawSearchReportModel> Report = new List<LawSearchReportModel>();
            //取出資料
            lawSearchDetails = DbHelper.Query<LawSearchDetail>(H2ORepository.ConnectionStringName, sql, new
            {
                LawDueAgentId = model.LawDueAgentId,
                LawDueName = "%" + model.LawDueName + "%",
                LawYear = model.LawYear,
                LawMonth = model.LawMonth,
                LawDoUnitId = model.LawDoUnitId,
                LawDueMoney = model.LawDueMoney,
                LawCloseType = "%" + model.LawCloseType + "%",
                LawStatusType = model.LawStatusType,
            }).ToList();
            //資料排版
            int SN = 1;
            for (int i = 0; i < lawSearchDetails.Count; i++)
            {
                LawSearchReportModel ReportModel = new LawSearchReportModel();
                ReportModel.SN = SN.ToString();
                ReportModel.LawNoteNo = lawSearchDetails[i].LawNoteNo;
                ReportModel.VmName = lawSearchDetails[i].VmName;
                ReportModel.SmName = lawSearchDetails[i].SmName;
                ReportModel.CenterName = lawSearchDetails[i].CenterName;
                ReportModel.WcCenterName = lawSearchDetails[i].WcCenterName;
                ReportModel.LawDueName = lawSearchDetails[i].LawDueName;
                ReportModel.LawDueAgentId = lawSearchDetails[i].LawDueAgentId;
                ReportModel.ProductionYm = lawSearchDetails[i].LawYear + "/" + lawSearchDetails[i].LawMonth + "-" + lawSearchDetails[i].LawPaySequence + "薪";
                ReportModel.LawDueMoney = lawSearchDetails[i].LawDueMoney;
                ReportModel.LawRepaymentCapital = lawSearchDetails[i].LawRepaymentCapital;
                List<LawLitigationProgress> litigationProgresses = GetLawLitigationProgressByLawIdNotoNo(lawSearchDetails[i].LawId.ToString(), lawSearchDetails[i].LawNoteNo);
                if (litigationProgresses.Count != 0)
                {
                    for (int j = 0; j < litigationProgresses.Count; j++)
                    {
                        if (!String.IsNullOrEmpty(litigationProgresses[j].LawLitigationprogress))
                        {
                            ReportModel.LawLitigationProgress = j == 0 ? "●" + litigationProgresses[j].LawLitigationprogress + "\n" : ReportModel.LawLitigationProgress + "●" + litigationProgresses[j].LawLitigationprogress + "\n";
                        }
                    }
                }
                List<LawDoProgress> lawDoProgresses = GetLawDoProgressByLawIdNotoNo(lawSearchDetails[i].LawId.ToString(), lawSearchDetails[i].LawNoteNo);
                if (lawDoProgresses.Count != 0)
                {
                    for (int j = 0; j < lawDoProgresses.Count; j++)
                    {
                        if (!String.IsNullOrEmpty(lawDoProgresses[j].LawDoprogress))
                        {
                            ReportModel.LawDoProgress = j == 0 ? "●" + lawDoProgresses[j].LawDoprogress + "\n" : ReportModel.LawDoProgress + "●" + lawDoProgresses[j].LawDoprogress + "\n";
                        }
                    }
                }
                ReportModel.LawDoUnitName = lawSearchDetails[i].LawDoUnitName;
                if (lawSearchDetails[i].LawStatusType == 2)
                {
                    ReportModel.LawNotCloseTypeName = "結案";
                    List<LawCloseType> lawCloseTypes = GetLawCloseTypeByLawCloseType(lawSearchDetails[i].LawCloseType);
                    if (lawCloseTypes.Count != 0)
                    {
                        for (int j = 0; j < lawCloseTypes.Count; j++)
                        {
                            ReportModel.CloseTypeName = j == 0 ? "●" + lawCloseTypes[j].CloseTypeName + "\n" : ReportModel.CloseTypeName + "●" + lawCloseTypes[j].CloseTypeName + "\n";
                        }
                    }
                }
                else
                {
                    ReportModel.LawNotCloseTypeName = (!String.IsNullOrEmpty(ReportModel.LawNotCloseTypeName)) ? ReportModel.LawNotCloseTypeName : "未結案";
                }


                Report.Add(ReportModel);
                SN++;
            }
            //匯出excel格式
            MemoryStream ms = null;
            //string HeadName = "法追系統-查詢明細";
            List<CustHeader> headList = new List<CustHeader>();
            CustHeader c1 = new CustHeader();
            c1.CellsNum = new List<int> { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            c1.CellsObj = new ArrayList() { "序號", "照會單號", "團隊", "協理系", "處別", "單位", "業務員", "ID", "結欠月", "結欠金額", "追回本金金額", "訴訟程序", "執行程序", "承辦單位", "案件狀態", "結案狀態" };
            headList.AddRange(new CustHeader[] { c1 });
            ms = NPoiHelper.Export(Report, headList, "法追系統-查詢明細");
            return ms.ToArray();
        }

        public byte[] GetNotCloseReportList(LawSearchDetail model)
        {
            string sql = @"select a.*,b.law_agent_content_id,production_ym,sequence,vm_code,vm_name,sm_code,sm_name,wc_center,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as wc_center_name,center_code,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as center_name,administrat_id,admin_name,admin_level,agent_code,name,ag_status_code,ag_level,ag_level_name,record_date,register_date,ag_status_date,super_account,create_date,(select isnull(sum(law_repayment_capital),0) from law_repayment_list
            where law_note_no=a.law_note_no) as law_repayment_capital from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where(law_content_cancel_type is null or law_content_cancel_type<>1) and law_status_type<>2 order by a.law_note_no asc";

            List<LawSearchDetail> lawSearchDetails = new List<LawSearchDetail>();
            List<LawNotCloseReportModel> Report = new List<LawNotCloseReportModel>();
            //取出資料
            lawSearchDetails = DbHelper.Query<LawSearchDetail>(H2ORepository.ConnectionStringName, sql, new
            {
                LawDueAgentId = model.LawDueAgentId,
                LawDueName = "%" + model.LawDueName + "%",
                LawYear = model.LawYear,
                LawMonth = model.LawMonth,
                LawDoUnitId = model.LawDoUnitId,
                LawDueMoney = model.LawDueMoney,
                LawCloseType = "%" + model.LawCloseType + "%",
                LawStatusType = model.LawStatusType,
            }).ToList();
            int SN = 1;
            //資料排版
            for (int i = 0; i < lawSearchDetails.Count; i++)
            {
                LawNotCloseReportModel ReportModel = new LawNotCloseReportModel();
                ReportModel.SN = SN.ToString();
                ReportModel.LawNoteNo = lawSearchDetails[i].LawNoteNo;
                ReportModel.VmName = lawSearchDetails[i].VmName;
                ReportModel.SmName = lawSearchDetails[i].SmName;
                ReportModel.CenterName = lawSearchDetails[i].CenterName;
                ReportModel.WcCenterName = lawSearchDetails[i].WcCenterName;
                ReportModel.LawDueName = lawSearchDetails[i].LawDueName;
                ReportModel.LawDueAgentId = lawSearchDetails[i].LawDueAgentId;
                ReportModel.ProductionYm = lawSearchDetails[i].LawYear + "/" + lawSearchDetails[i].LawMonth + "-" + lawSearchDetails[i].LawPaySequence + "薪";
                ReportModel.LawDueMoney = lawSearchDetails[i].LawDueMoney;
                ReportModel.LawRepaymentCapital = lawSearchDetails[i].LawRepaymentCapital;
                ReportModel.LawPhoneCall1Desc = lawSearchDetails[i].LawPhoneCall1Desc;
                ReportModel.LawPhoneCall2Desc = lawSearchDetails[i].LawPhoneCall2Desc;
                LawEvidenceDesc lawEvidenceDesc = GetLawEvidenceDescByLawId(lawSearchDetails[i].LawId.ToString());
                ReportModel.LawEvidencedesc = lawEvidenceDesc != null ? lawEvidenceDesc.LawEvidencedesc : "";
                List<LawLitigationProgress> litigationProgresses = GetLawLitigationProgressByLawIdNotoNo(lawSearchDetails[i].LawId.ToString(), lawSearchDetails[i].LawNoteNo);
                if (litigationProgresses.Count != 0)
                {
                    for (int j = 0; j < litigationProgresses.Count; j++)
                    {
                        if (!String.IsNullOrEmpty(litigationProgresses[j].LawLitigationprogress))
                        {
                            ReportModel.LawLitigationProgress = j == 0 ? "●" + litigationProgresses[j].LawLitigationprogress + "\n" : ReportModel.LawLitigationProgress + "●" + litigationProgresses[j].LawLitigationprogress + "\n";
                        }
                    }
                }
                List<LawDoProgress> lawDoProgresses = GetLawDoProgressByLawIdNotoNo(lawSearchDetails[i].LawId.ToString(), lawSearchDetails[i].LawNoteNo);
                if (lawDoProgresses.Count != 0)
                {
                    for (int j = 0; j < lawDoProgresses.Count; j++)
                    {
                        if (!String.IsNullOrEmpty(lawDoProgresses[j].LawDoprogress))
                        {
                            ReportModel.LawDoProgress = j == 0 ? "●" + lawDoProgresses[j].LawDoprogress + "\n" : ReportModel.LawDoProgress + "●" + lawDoProgresses[j].LawDoprogress + "\n";
                        }
                    }
                }
                ReportModel.LawDoUnitName = lawSearchDetails[i].LawDoUnitName;

                Report.Add(ReportModel);
                SN++;
            }
            //匯出excel格式
            MemoryStream ms = null;
            //string HeadName = "法追系統-查詢明細";
            List<CustHeader> headList = new List<CustHeader>();
            CustHeader c1 = new CustHeader();
            c1.CellsNum = new List<int> { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            c1.CellsObj = new ArrayList() { "序號", "照會單號", "團隊", "協理系", "處別", "單位", "業務員", "ID", "結欠月", "結欠金額", "追回本金金額", "第一次電催", "第二次電催", "存證信函", "訴訟程序", "執行程序", "承辦單位" };
            headList.AddRange(new CustHeader[] { c1 });
            ms = NPoiHelper.Export(Report, headList, "法追系統-未結案件明細");
            return ms.ToArray();
        }

        public byte[] GetCloseReportList(LawSearchDetail model)
        {
            string sql = @"select *,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG
            ,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where law_status_type=2 or(law_content_cancel_type is not null or law_content_cancel_type=1)";

            List<LawSearchDetail> lawSearchDetails = new List<LawSearchDetail>();
            List<LawCloseReportModel> Report = new List<LawCloseReportModel>();
            //取出資料
            lawSearchDetails = DbHelper.Query<LawSearchDetail>(H2ORepository.ConnectionStringName, sql, new
            {
                LawDueAgentId = model.LawDueAgentId,
                LawDueName = "%" + model.LawDueName + "%",
                LawYear = model.LawYear,
                LawMonth = model.LawMonth,
                LawDoUnitId = model.LawDoUnitId,
                LawDueMoney = model.LawDueMoney,
                LawCloseType = "%" + model.LawCloseType + "%",
                LawStatusType = model.LawStatusType,
            }).ToList();
            int SN = 1;
            //資料排版
            for (int i = 0; i < lawSearchDetails.Count; i++)
            {
                LawCloseReportModel ReportModel = new LawCloseReportModel();
                ReportModel.SN = SN.ToString();
                ReportModel.LawNoteNo = lawSearchDetails[i].LawNoteNo;
                ReportModel.VmName = lawSearchDetails[i].VmName;
                ReportModel.SmName = lawSearchDetails[i].SmName;
                ReportModel.CenterName = lawSearchDetails[i].CenterNameCG;
                ReportModel.WcCenterName = lawSearchDetails[i].WcCenterNameCG;
                ReportModel.LawDueName = lawSearchDetails[i].LawDueName;
                ReportModel.LawDueAgentId = lawSearchDetails[i].LawDueAgentId;
                ReportModel.ProductionYm = lawSearchDetails[i].LawYear + "/" + lawSearchDetails[i].LawMonth + "-" + lawSearchDetails[i].LawPaySequence + "薪";
                ReportModel.LawDueMoney = lawSearchDetails[i].LawDueMoney;
                ReportModel.LawRepaymentCapital = lawSearchDetails[i].LawRepaymentCapital;
                ReportModel.LawPhoneCall1Desc = lawSearchDetails[i].LawPhoneCall1Desc;
                ReportModel.LawPhoneCall2Desc = lawSearchDetails[i].LawPhoneCall2Desc;
                LawEvidenceDesc lawEvidenceDesc = GetLawEvidenceDescByLawId(lawSearchDetails[i].LawId.ToString());
                ReportModel.LawEvidencedesc = lawEvidenceDesc != null ? lawEvidenceDesc.LawEvidencedesc : "";
                List<LawLitigationProgress> litigationProgresses = GetLawLitigationProgressByLawIdNotoNo(lawSearchDetails[i].LawId.ToString(), lawSearchDetails[i].LawNoteNo);
                if (litigationProgresses.Count != 0)
                {
                    for (int j = 0; j < litigationProgresses.Count; j++)
                    {
                        if (!String.IsNullOrEmpty(litigationProgresses[j].LawLitigationprogress))
                        {
                            ReportModel.LawLitigationProgress = j == 0 ? "●" + litigationProgresses[j].LawLitigationprogress + "\n" : ReportModel.LawLitigationProgress + "●" + litigationProgresses[j].LawLitigationprogress + "\n";
                        }
                    }
                }
                List<LawDoProgress> lawDoProgresses = GetLawDoProgressByLawIdNotoNo(lawSearchDetails[i].LawId.ToString(), lawSearchDetails[i].LawNoteNo);
                if (lawDoProgresses.Count != 0)
                {
                    for (int j = 0; j < lawDoProgresses.Count; j++)
                    {
                        if (!String.IsNullOrEmpty(lawDoProgresses[j].LawDoprogress))
                        {
                            ReportModel.LawDoProgress = j == 0 ? "●" + lawDoProgresses[j].LawDoprogress + "\n" : ReportModel.LawDoProgress + "●" + lawDoProgresses[j].LawDoprogress + "\n";
                        }
                    }
                }
                ReportModel.LawDoUnitName = lawSearchDetails[i].LawDoUnitName;

                Report.Add(ReportModel);
                SN++;
            }
            //匯出excel格式
            MemoryStream ms = null;
            //string HeadName = "法追系統-查詢明細";
            List<CustHeader> headList = new List<CustHeader>();
            CustHeader c1 = new CustHeader();
            c1.CellsNum = new List<int> { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            c1.CellsObj = new ArrayList() { "序號", "照會單號", "團隊", "協理系", "處別", "單位", "業務員", "ID", "結欠月", "結欠金額", "追回本金金額", "第一次電催", "第二次電催", "存證信函", "訴訟程序", "執行程序", "承辦單位" };
            headList.AddRange(new CustHeader[] { c1 });
            ms = NPoiHelper.Export(Report, headList, "法追系統-已結案件明細");
            return ms.ToArray();
        }
        #endregion

        #endregion

        #region 報表作業
        #region 法院統計報表作業
        /// <summary>
        /// 報表作業
        /// </summary>
        public List<LawMasterReportLog> GetLawMasterReportLog()
        {
            List<LawMasterReportLog> result = new List<LawMasterReportLog>();
            int year = DateTime.Now.Year - 2006 + 1;
            string sql = @"select * from law_master_report_log where law_master_id in(select top @year law_master_id from law_master_report_log where create_date=(select max(create_date) from law_master_report_log)) order by law_master_id";

            result = DbHelper.Query<LawMasterReportLog>(H2ORepository.ConnectionStringName, sql, new { year = year }).ToList();
            return result;
        }

        /// <summary>
        /// 法追統計
        /// </summary>
        public LawContent GetLawcontentByLawyear(int year)
        {
            LawContent result = new LawContent();
            //string year = DateTime.Now.ToString("yyyy");
            string sql = @"select isnull(sum(law_due_money),0) as law_due_money from law_content where law_year=" + year + " and law_content_cancel_type is null";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 法追統計
        /// </summary>
        public LawContent GetLawcontentByLawyears(string year_str)
        {
            LawContent result = new LawContent();
            //string year = DateTime.Now.ToString("yyyy");
            string sql = @"select isnull(sum(law_due_money),0) as law_due_money from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 節欠總金額
        /// </summary>
        public LawMasterReportLog GetLawMasterReportLogByLawyear(int y_str, string year_str)
        {
            LawMasterReportLog result = new LawMasterReportLog();
            int ystr = y_str - 95 + 1;
            string sql = @"select sum(law_total_due) as law_total_due,sum(law_total_repayment) as law_total_repayment,create_date from law_master_report_log where law_master_id in(select top " + ystr + " law_master_id from law_master_report_log where create_date=(select max(create_date) from law_master_report_log)) and law_year in(" + year_str + ") group by create_date";

            result = DbHelper.Query<LawMasterReportLog>(H2ORepository.ConnectionStringName, sql).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 節欠其他總金額
        /// </summary>
        public LawMasterReportLog GetLawMasterReportLogOtherByLawyear(int y_str, int Lawyear)
        {
            LawMasterReportLog result = new LawMasterReportLog();
            int ystr = y_str - 95 + 1;
            string sql = @"select * from law_master_report_log where law_master_id in(select top " + ystr + " law_master_id from law_master_report_log where create_date=(select max(create_date) from law_master_report_log)) and law_year=@Lawyear ";

            result = DbHelper.Query<LawMasterReportLog>(H2ORepository.ConnectionStringName, sql, new { Lawyear = Lawyear }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 清償金額合計
        /// </summary>
        public LawRepaymentList GetSumLawRepaymentList(string year_str)
        {
            LawRepaymentList result = new LawRepaymentList();

            string sql = @"select (aa.law_repayment_money+bb.law_repayment_money+cc.law_repayment_money) as law_repayment_money from (select isnull(a.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null and law_status_type<>2 ) ) as a) as aa,(select isnull(b.law_repayment_money,0) as law_repayment_money from (select sum(law_due_money)as law_repayment_money from law_report_reference where law_year in(" + year_str + ") and law_note_no in(select law_note_no from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) not in(select close_type_id from law_close_type where count_type=0 ))) as b ) as bb,(select isnull(c.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) in(select close_type_id from law_close_type where count_type=0 )) ) as c) as cc";

            result = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 清償金額
        /// </summary>
        public LawRepaymentList GetLawRepaymentList(int Lawyear)
        {
            LawRepaymentList result = new LawRepaymentList();

            string sql = @"select (aa.law_repayment_money+bb.law_repayment_money+cc.law_repayment_money) as law_repayment_money from (select isnull(a.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year=@Lawyear and law_content_cancel_type is null and law_status_type<>2) ) as a) as aa,(select isnull(b.law_repayment_money,0) as law_repayment_money from (select sum(law_due_money)as law_repayment_money from law_report_reference where law_year=@Lawyear and law_note_no in(select law_note_no from law_content where law_year=@Lawyear and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) not in(select close_type_id from law_close_type where count_type=0 ))) as b ) as bb,(select isnull(c.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year=@Lawyear and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) in(select close_type_id from law_close_type where count_type=0 )) ) as c) as cc";

            result = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql, new { Lawyear = Lawyear }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 系統紀錄log
        /// </summary>
        public bool InsertLawMasterReportLog(string MeberID)
        {
            bool result = false;
            decimal due_money1;
            string pstr = string.Empty;
            int repayment_money1;
            int yearnow = DateTime.Now.Year - 1911;
            for (int i = 95; i <= yearnow; i++)
            {
                string sql = @"select isnull(sum(law_due_money),0) as law_due_money from law_content where law_year=@year and law_content_cancel_type is null";
                LawContent lawcontent = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { year = i }).FirstOrDefault();
                due_money1 = lawcontent != null ? lawcontent.LawDueMoney : 0;

                string sql1 = @"select (aa.law_repayment_money+bb.law_repayment_money+cc.law_repayment_money) as law_repayment_money from (select isnull(a.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year=@year and law_content_cancel_type is null and law_status_type<>2) ) as a) as aa,(select isnull(b.law_repayment_money,0) as law_repayment_money from (select sum(law_due_money)as law_repayment_money from law_report_reference where law_year=@year and law_note_no in(select law_note_no from law_content where law_year=@year and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) not in(select close_type_id from law_close_type where count_type=0 ))) as b  ) as bb,(select isnull(c.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year=@year and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) in(select close_type_id from law_close_type where count_type=0 ) ) ) as c )as cc";
                LawRepaymentList lawRepaymentList = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql1, new { year = i }).FirstOrDefault();
                repayment_money1 = lawRepaymentList != null ? lawRepaymentList.LawRepaymentMoney : 0;
                if (due_money1 != 0 && repayment_money1 != 0)
                {
                    decimal decimalValue = (repayment_money1 / due_money1) * 100;
                    pstr = Math.Round(decimalValue, 2) + "%";
                }
                else
                {
                    if (due_money1 != 0 && repayment_money1 == 0)
                    {
                        pstr = "0%";
                    }
                    else
                    {
                        pstr = "";
                    }
                }

                string sql2 = @"INSERT INTO law_master_report_log(law_year,law_total_due,law_total_repayment,law_repay_percent,create_date,MasterReportLogCreatorID)
                           VALUES (@LawYear, @LawTotalDue, @LawTotalRepayment, @LawRepayPercent, @CreateDate,@MasterReportLogCreatorID);";
                DbHelper.Execute(H2ORepository.ConnectionStringName, sql2, new
                {
                    LawYear = i,
                    LawTotalDue = due_money1,
                    LawTotalRepayment = repayment_money1,
                    LawRepayPercent = pstr,
                    CreateDate = DateTime.Now.ToString("yyyy/MM/dd"),
                    MasterReportLogCreatorID = MeberID
                });

                result = true;
            }

            return result;
        }

        #region 報表

        //public byte[] GetSumStatisticReportList(StatisticsDetail model)
        //{
        //    int set_year, y_str;
        //    string year_str = string.Empty;

        //    set_year = Convert.ToInt32(model.LawYearType);
        //    List<StatisticsDetail> statistics = new List<StatisticsDetail>();

        //    for (int i = 95; i <= set_year; i++)
        //    {
        //        year_str = i == 95 ? "95" : year_str + "," + i;
        //    }
        //    y_str = DateTime.Now.Year - 1911;

        //    decimal due_money;
        //    int repayment_money;
        //    string datett = (DateTime.Now.Year - 1911) + DateTime.Now.ToString("MMdd");
        //    int datenow = DateTime.Now.Year - 2006 + 1;
        //    string p_str, repay_per_str;
        //    string sql, sql1, sql2;
        //    LawMasterReportLog reportLog = new LawMasterReportLog();
        //    LawContent content = new LawContent();
        //    LawRepaymentList lawRepaymentList = new LawRepaymentList();
        //    List<LawMasterReportModel> Report = new List<LawMasterReportModel>();

        //    //取出資料
        //    List<CustHeader> headList = new List<CustHeader>();
        //    //資料排版
        //    for (int i = set_year; i <= y_str; i++)
        //    {
        //        LawMasterReportModel lawMasterReport = new LawMasterReportModel();
        //        if (i == set_year)
        //        {
        //            sql = @"select sum(law_total_due) as law_total_due,sum(law_total_repayment) as law_total_repayment,create_date from law_master_report_log where law_master_id in(select top " + datenow + " law_master_id from law_master_report_log where create_date=(select max(create_date) from law_master_report_log)) and law_year in(" + year_str + ") group by create_date";
        //            reportLog = DbHelper.Query<LawMasterReportLog>(H2ORepository.ConnectionStringName, sql).FirstOrDefault();
        //        }
        //        else
        //        {
        //            sql = @"select * from law_master_report_log where law_master_id in(select top " + datenow + " law_master_id from law_master_report_log where create_date=(select max(create_date) from law_master_report_log)) and law_year=" + i + "";
        //            reportLog = DbHelper.Query<LawMasterReportLog>(H2ORepository.ConnectionStringName, sql).FirstOrDefault();
        //        }
        //        lawMasterReport.DateNow = (DateTime.Now.Year - 1911) + DateTime.Now.ToString("MMdd");

        //        CustHeader c1 = new CustHeader();
        //        CustHeader c2 = new CustHeader();
        //        CustHeader c3 = new CustHeader();
        //        CustHeader c4 = new CustHeader();
        //        CustHeader c5 = new CustHeader();
        //        CustHeader c6 = new CustHeader();
        //        if (reportLog != null)
        //        {
        //            lawMasterReport.DateHave = (Convert.ToInt32(reportLog.CreateDate.Substring(0, 4)) - 1911).ToString() + reportLog.CreateDate.Substring(5, 2) + reportLog.CreateDate.Substring(8, 2);
        //            if (i == set_year)
        //            {
        //                c2.CellsNum = new List<int> { 2, 2, 2 };
        //                c2.CellsObj = new ArrayList() { "", "95 ~ " + i + " 年度統計/" + lawMasterReport.DateNow, lawMasterReport.DateHave };
        //                decimal decimalValue = Convert.ToDecimal(reportLog.LawTotalRepayment) / Convert.ToDecimal(reportLog.LawTotalDue) * 100;
        //                repay_per_str = Math.Round(decimalValue, 2) + "%";
        //            }
        //            else
        //            {
        //                c2.CellsNum = new List<int> { 2, 2, 2 };
        //                c2.CellsObj = new ArrayList() { "", i + " 年度統計/" + lawMasterReport.DateNow, lawMasterReport.DateHave };
        //                repay_per_str = reportLog.LawRepayPercent;
        //            }

        //        }
        //        else
        //        {
        //            c2.CellsNum = new List<int> { 2, 2, 2 };
        //            c2.CellsObj = new ArrayList() { "", i + " 年度統計/" + lawMasterReport.DateNow, "" };
        //            repay_per_str = "";
        //        }

        //        if (i == set_year)
        //        {
        //            sql1 = @"select isnull(sum(law_due_money),0) as law_due_money from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null";
        //            content = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql1).FirstOrDefault();
        //        }
        //        else
        //        {
        //            sql1 = @"select isnull(sum(law_due_money),0) as law_due_money from law_content where law_year=" + i + " and law_content_cancel_type is null";
        //            content = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql1).FirstOrDefault();
        //        }

        //        if (content != null)
        //        {
        //            if (reportLog != null)
        //            {
        //                c3.CellsObj = new ArrayList() { "結欠總金額", content.LawDueMoney, reportLog.LawTotalDue };
        //            }
        //            else
        //            {
        //                c3.CellsObj = new ArrayList() { "結欠總金額", content.LawDueMoney, "" };
        //            }
        //            due_money = content.LawDueMoney;
        //        }
        //        else
        //        {
        //            c3.CellsObj = new ArrayList() { "結欠總金額", 0, 0 };
        //            due_money = 0;
        //        }

        //        if (i == set_year)
        //        {
        //            sql2 = @"select (aa.law_repayment_money+bb.law_repayment_money+cc.law_repayment_money) as law_repayment_money from (select isnull(a.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null and law_status_type<>2 ) ) as a) as aa,(select isnull(b.law_repayment_money,0) as law_repayment_money from (select sum(law_due_money)as law_repayment_money from law_report_reference where law_year in(" + year_str + ") and law_note_no in(select law_note_no from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) not in(select close_type_id from law_close_type where count_type=0 ))) as b ) as bb,(select isnull(c.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) in(select close_type_id from law_close_type where count_type=0 )) ) as c) as cc ";
        //            lawRepaymentList = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql2).FirstOrDefault();
        //        }
        //        else
        //        {
        //            sql2 = @"select (aa.law_repayment_money+bb.law_repayment_money+cc.law_repayment_money) as law_repayment_money from (select isnull(a.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year=" + i + " and law_content_cancel_type is null and law_status_type<>2) ) as a) as aa,(select isnull(b.law_repayment_money,0) as law_repayment_money from (select sum(law_due_money)as law_repayment_money from law_report_reference where law_year=" + i + " and law_note_no in(select law_note_no from law_content where law_year=" + i + " and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) not in(select close_type_id from law_close_type where count_type=0 ))) as b ) as bb,(select isnull(c.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year=" + i + " and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) in(select close_type_id from law_close_type where count_type=0 )) ) as c) as cc";
        //            lawRepaymentList = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql2).FirstOrDefault();
        //        }

        //        if (lawRepaymentList != null)
        //        {
        //            if (reportLog != null)
        //            {
        //                c4.CellsObj = new ArrayList() { "已清償本金累計金額", lawRepaymentList.LawRepaymentMoney, reportLog.LawTotalRepayment };
        //            }
        //            else
        //            {
        //                c4.CellsObj = new ArrayList() { "已清償本金累計金額", lawRepaymentList.LawRepaymentMoney, "" };
        //            }

        //            repayment_money = lawRepaymentList.LawRepaymentMoney;
        //        }
        //        else
        //        {
        //            c4.CellsObj = new ArrayList() { "已清償本金累計金額", 0, 0 };
        //            repayment_money = 0;
        //        }

        //        if (due_money != 0 && repayment_money != 0)
        //        {
        //            decimal decimalValue = repayment_money / due_money * 100;
        //            p_str = Math.Round(decimalValue, 2) + "%";
        //        }
        //        else
        //        {
        //            if (due_money != 0 && repayment_money == 0)
        //            {
        //                p_str = "0%";
        //            }
        //            else
        //            {
        //                p_str = "";
        //            }
        //        }

        //        //c1~c5 迴圈            
        //        c1.CellsNum = new List<int> { 2, 2, 2 };
        //        c1.CellsObj = new ArrayList() { "總表", lawMasterReport.DateNow, "(含續佣抵結欠)" };

        //        c3.CellsNum = new List<int> { 2, 2, 2 };
        //        c4.CellsNum = new List<int> { 2, 2, 2 };

        //        c5.CellsNum = new List<int> { 2, 2, 2 };
        //        c5.CellsObj = new ArrayList() { "達成率【金額】", p_str, repay_per_str };

        //        c6.CellsNum = new List<int> { 6 };
        //        c6.CellsObj = new ArrayList() { "" };

        //        headList.AddRange(new CustHeader[] { c1, c2, c3, c4, c5, c6 });
        //    }
        //    //匯出excel格式
        //    MemoryStream ms = null;
        //    ms = NPoiHelper.Export(Report, headList, "法追系統-當月(" + datett + ")總表");
        //    return ms.ToArray();
        //}

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="year">年度</param>
        /// <returns>Stream</returns>
        public Stream GetStatisticReportList(string LawYearType)
        {
            MemoryStream ms = new MemoryStream();
            ExcelPackage excel = new ExcelPackage();
            decimal due_money;
            int repayment_money;
            string datett = (DateTime.Now.Year - 1911) + DateTime.Now.ToString("MMdd");
            string year_str = string.Empty;
            int datenow = DateTime.Now.Year - 2006 + 1;
            int datet = DateTime.Now.Year - 1911;
            string p_str, repay_per_str;
            ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("法追系統 - 當月(" + datett + ")總表");
            int d = 3, g = 5, s = 2, z = 4, p = 6;
            LawMasterReportModel lawMasterReport = new LawMasterReportModel();
            lawMasterReport.DateNow = (DateTime.Now.Year - 1911) + DateTime.Now.ToString("MMdd");

            for (int i = 95; i <= datet; i++)
            {
                LawMasterReportLog reportLog = new LawMasterReportLog();
                reportLog = GetLawMasterReportLogOtherByLawyear(datet, i);
                if (reportLog != null)
                {
                    lawMasterReport.DateHave = (Convert.ToInt32(reportLog.CreateDate.Substring(0, 4)) - 1911).ToString() + reportLog.CreateDate.Substring(5, 2) + reportLog.CreateDate.Substring(8, 2);
                    repay_per_str = reportLog.LawRepayPercent;
                }
                else
                {
                    lawMasterReport.DateHave = DateTime.Now.ToString("yyyy/MM/dd");
                    repay_per_str = "";
                }
                string sql1 = @"select sum(law_due_money) as law_due_money from law_content where law_year=@year and law_content_cancel_type is null";
                LawContent lawContent = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql1, new { year = i }).FirstOrDefault();
                string sql2 = @"select (aa.law_repayment_money+bb.law_repayment_money+cc.law_repayment_money) as law_repayment_money from (select isnull(a.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year=@year and law_content_cancel_type is null and law_status_type<>2) ) as a) as aa,(select isnull(b.law_repayment_money,0) as law_repayment_money from (select sum(law_due_money)as law_repayment_money from law_report_reference where law_year=@year and law_note_no in(select law_note_no from law_content where law_year=@year and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) not in(select close_type_id from law_close_type where count_type=0 ))) as b  ) as bb,(select isnull(c.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year=@year and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) in(select close_type_id from law_close_type where count_type=0 ) ) ) as c )as cc";
                LawRepaymentList lawRepaymentList = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql2, new { year = i }).FirstOrDefault();

                ExcelSetCell(sheet, new string[] { "" }, s, 1);
                ExcelSetCell(sheet, new string[] { "" }, s, 2);
                sheet.Cells[s, 1, s, 2].Merge = true;
                sheet.Cells[s, 1, s, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { i + " 年度統計/" + lawMasterReport.DateNow }, s, 3);
                ExcelSetCell(sheet, new string[] { "" }, s, 4);
                sheet.Cells[s, 3, s, 4].Merge = true;
                sheet.Cells[s, 3, s, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { lawMasterReport.DateHave }, s, 5);
                ExcelSetCell(sheet, new string[] { "" }, s, 6);
                sheet.Cells[s, 5, s, 6].Merge = true;
                sheet.Cells[s, 5, s, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                s = s + 5;

                if (lawContent != null)
                {
                    if (reportLog != null)
                    {
                        //c3.CellsObj = new ArrayList() { "結欠總金額", lawContent.LawDueMoney, reportLog.LawTotalDue };
                        ExcelSetCell(sheet, new string[] { "結欠總金額" }, d, 1);
                        ExcelSetCell(sheet, new string[] { "" }, d, 2);
                        sheet.Cells[d, 1, d, 2].Merge = true;
                        sheet.Cells[d, 1, d, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { lawContent.LawDueMoney.ToString("###,###") }, d, 3);
                        ExcelSetCell(sheet, new string[] { "" }, d, 4);
                        sheet.Cells[d, 3, d, 4].Merge = true;
                        sheet.Cells[d, 3, d, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { reportLog.LawTotalDue.ToString("###,###") }, d, 5);
                        ExcelSetCell(sheet, new string[] { "" }, d, 6);
                        sheet.Cells[d, 5, d, 6].Merge = true;
                        sheet.Cells[d, 5, d, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    else
                    {
                        //c3.CellsObj = new ArrayList() { "結欠總金額", lawContent.LawDueMoney, "" };
                        ExcelSetCell(sheet, new string[] { "結欠總金額" }, d, 1);
                        ExcelSetCell(sheet, new string[] { "" }, d, 2);
                        sheet.Cells[d, 1, d, 2].Merge = true;
                        sheet.Cells[d, 1, d, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { lawContent.LawDueMoney.ToString("###,###") }, d, 3);
                        ExcelSetCell(sheet, new string[] { "" }, d, 4);
                        sheet.Cells[d, 3, d, 4].Merge = true;
                        sheet.Cells[d, 3, d, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { "" }, d, 5);
                        ExcelSetCell(sheet, new string[] { "" }, d, 6);
                        sheet.Cells[d, 5, d, 6].Merge = true;
                        sheet.Cells[d, 5, d, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    due_money = lawContent.LawDueMoney;
                }
                else
                {
                    //c3.CellsObj = new ArrayList() { "結欠總金額", 0, 0 };
                    ExcelSetCell(sheet, new string[] { "結欠總金額" }, d, 1);
                    ExcelSetCell(sheet, new string[] { "" }, d, 2);
                    sheet.Cells[d, 1, d, 2].Merge = true;
                    sheet.Cells[d, 1, d, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    ExcelSetCell(sheet, new string[] { "0" }, d, 3);
                    ExcelSetCell(sheet, new string[] { "" }, d, 4);
                    sheet.Cells[d, 3, d, 4].Merge = true;
                    sheet.Cells[d, 3, d, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    ExcelSetCell(sheet, new string[] { "0" }, d, 5);
                    ExcelSetCell(sheet, new string[] { "" }, d, 6);
                    sheet.Cells[d, 5, d, 6].Merge = true;
                    sheet.Cells[d, 5, d, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    due_money = 0;
                }
                d = d + 5;
                if (lawRepaymentList != null)
                {
                    //c4.CellsObj = new ArrayList() { "已清償本金累計金額", lawRepaymentList.LawRepaymentMoney, reportLog.LawTotalRepayment };
                    ExcelSetCell(sheet, new string[] { "已清償本金累計金額" }, z, 1);
                    ExcelSetCell(sheet, new string[] { "" }, z, 2);
                    sheet.Cells[z, 1, z, 2].Merge = true;
                    sheet.Cells[z, 1, z, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    ExcelSetCell(sheet, new string[] { lawRepaymentList.LawRepaymentMoney.ToString("###,###") }, z, 3);
                    ExcelSetCell(sheet, new string[] { "" }, z, 4);
                    sheet.Cells[z, 3, z, 4].Merge = true;
                    sheet.Cells[z, 3, z, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    if (reportLog != null)
                    {
                        ExcelSetCell(sheet, new string[] { reportLog.LawTotalRepayment.ToString("###,###") }, z, 5);
					}
					else
					{
                        ExcelSetCell(sheet, new string[] { "" }, z, 5);
                    }
                    ExcelSetCell(sheet, new string[] { "" }, z, 6);
                    sheet.Cells[z, 5, z, 6].Merge = true;
                    sheet.Cells[z, 5, z, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    repayment_money = lawRepaymentList.LawRepaymentMoney;
                }
                else
                {
                    //c4.CellsObj = new ArrayList() { "已清償本金累計金額", 0, 0 };
                    ExcelSetCell(sheet, new string[] { "已清償本金累計金額" }, z, 1);
                    ExcelSetCell(sheet, new string[] { "" }, z, 2);
                    sheet.Cells[z, 1, z, 2].Merge = true;
                    sheet.Cells[z, 1, z, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    ExcelSetCell(sheet, new string[] { "0" }, z, 3);
                    ExcelSetCell(sheet, new string[] { "" }, z, 4);
                    sheet.Cells[z, 3, z, 4].Merge = true;
                    sheet.Cells[z, 3, z, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    ExcelSetCell(sheet, new string[] { "0" }, z, 5);
                    ExcelSetCell(sheet, new string[] { "" }, z, 6);
                    sheet.Cells[z, 5, z, 6].Merge = true;
                    sheet.Cells[z, 5, z, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    repayment_money = 0;
                }
                z = z + 5;
                if (due_money != 0 && repayment_money != 0)
                {
                    //decimal decimalValue = Convert.ToDecimal(reportLog.LawTotalRepayment) / Convert.ToDecimal(reportLog.LawTotalDue) * 100;
                    decimal decimalValue = repayment_money / due_money * 100;
                    p_str = Math.Round(decimalValue, 2) + "%";
                }
                else
                {
                    if (due_money != 0 && repayment_money == 0)
                    {
                        p_str = "0%";
                    }
                    else
                    {
                        p_str = "";
                    }
                }
                //#FFFF66
                ExcelSetCell(sheet, new string[] { "達成率【金額】" }, g, 1);
                ExcelSetCell(sheet, new string[] { "" }, g, 2);
                sheet.Cells[g, 1, g, 2].Merge = true;
                sheet.Cells[g, 1, g, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { p_str }, g, 3);
                ExcelSetCell(sheet, new string[] { "" }, g, 4);
                sheet.Cells[g, 3, g, 4].Merge = true;
                sheet.Cells[g, 3, g, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[g, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[g, 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFF66"));
                sheet.Cells[g, 3].Style.Font.Color.SetColor(Color.Blue);
                //ExcelSetCell(sheet, new string[] { item.RepayMoney == "0" ? "0" : Convert.ToInt32(item.RepayMoney).ToString("###,###") }, g, w);

                ExcelSetCell(sheet, new string[] { repay_per_str }, g, 5);
                ExcelSetCell(sheet, new string[] { "" }, g, 6);
                sheet.Cells[g, 5, g, 6].Merge = true;
                sheet.Cells[g, 5, g, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[g, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[g, 5].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFF66"));
                sheet.Cells[g, 5].Style.Font.Color.SetColor(Color.Blue);
                g = g + 5;

                ExcelSetCell(sheet, new string[] { "" }, p, 1);
                ExcelSetCell(sheet, new string[] { "" }, p, 2);
                ExcelSetCell(sheet, new string[] { "" }, p, 3);
                ExcelSetCell(sheet, new string[] { "" }, p, 4);
                ExcelSetCell(sheet, new string[] { "" }, p, 5);
                ExcelSetCell(sheet, new string[] { "" }, p, 6);
                sheet.Cells[p, 1, p, 6].Merge = true;
                p = p + 5;

            }

            //sheet 標題 橫排 直排
            ExcelSetCell(sheet, new string[] { "總表" }, 1, 1);
            ExcelSetCell(sheet, new string[] { "" }, 1, 2);
            sheet.Cells[1, 1, 1, 2].Merge = true;
            sheet.Cells[1, 1, 1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //sheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 1].Style.Font.Color.SetColor(Color.Blue);

            ExcelSetCell(sheet, new string[] { lawMasterReport.DateNow }, 1, 3);
            ExcelSetCell(sheet, new string[] { "" }, 1, 4);
            sheet.Cells[1, 3, 1, 4].Merge = true;
            sheet.Cells[1, 3, 1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //sheet.Cells[1, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 3].Style.Font.Color.SetColor(Color.Blue);

            ExcelSetCell(sheet, new string[] { "(含續佣抵結欠)" }, 1, 5);
            ExcelSetCell(sheet, new string[] { "" }, 1, 6);
            sheet.Cells[1, 5, 1, 6].Merge = true;
            sheet.Cells[1, 5, 1, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //sheet.Cells[1, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 5].Style.Font.Color.SetColor(Color.Blue);

            sheet.Column(1).Width = 22;
            sheet.Cells.Style.ShrinkToFit = true;
            //字型
            sheet.Cells.Style.Font.Name = "新細明體";
            //文字大小
            sheet.Cells.Style.Font.Size = 12;
            excel.SaveAs(ms);
            excel.Dispose();
            ms.Position = 0;

            return ms;
        }

        #region 合計報表
        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="year">年度</param>
        /// <returns>Stream</returns>
        public Stream GetSumStatisticReportList(string LawYearType)
        {
            MemoryStream ms = new MemoryStream();
            ExcelPackage excel = new ExcelPackage();
            decimal due_money;
            int repayment_money;
            string datett = (DateTime.Now.Year - 1911) + DateTime.Now.ToString("MMdd");
            int datenow = DateTime.Now.Year - 2006 + 1;
            int datet = DateTime.Now.Year - 1911;
            string p_str, repay_per_str;
            ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("法追系統-當月(" + datett + ")總表_年度合併");
            int d = 3, g = 5, s = 2, z = 4, p = 6;
            int set_year, y_str;
            string year_str = string.Empty;

            set_year = Convert.ToInt32(LawYearType);
            List<StatisticsDetail> statistics = new List<StatisticsDetail>();

            for (int i = 95; i <= set_year; i++)
            {
                year_str = i == 95 ? "95" : year_str + "," + i;
            }
            y_str = DateTime.Now.Year - 1911;
            string sql, sql1, sql2;
            LawMasterReportLog reportLog = new LawMasterReportLog();
            LawContent content = new LawContent();
            LawRepaymentList lawRepaymentList = new LawRepaymentList();
            List<LawMasterReportModel> Report = new List<LawMasterReportModel>();
            LawMasterReportModel lawMasterReport = new LawMasterReportModel();
            //取出資料
            List<CustHeader> headList = new List<CustHeader>();
            //資料排版
            for (int i = set_year; i <= y_str; i++)
            {

                if (i == set_year)
                {
                    sql = @"select sum(law_total_due) as law_total_due,sum(law_total_repayment) as law_total_repayment,create_date from law_master_report_log where law_master_id in(select top " + datenow + " law_master_id from law_master_report_log where create_date=(select max(create_date) from law_master_report_log)) and law_year in(" + year_str + ") group by create_date";
                    reportLog = DbHelper.Query<LawMasterReportLog>(H2ORepository.ConnectionStringName, sql).FirstOrDefault();
                }
                else
                {
                    sql = @"select * from law_master_report_log where law_master_id in(select top " + datenow + " law_master_id from law_master_report_log where create_date=(select max(create_date) from law_master_report_log)) and law_year=" + i + "";
                    reportLog = DbHelper.Query<LawMasterReportLog>(H2ORepository.ConnectionStringName, sql).FirstOrDefault();
                }
                lawMasterReport.DateNow = (DateTime.Now.Year - 1911) + DateTime.Now.ToString("MMdd");

                if (reportLog != null)
                {
                    lawMasterReport.DateHave = (Convert.ToInt32(reportLog.CreateDate.Substring(0, 4)) - 1911).ToString() + reportLog.CreateDate.Substring(5, 2) + reportLog.CreateDate.Substring(8, 2);
                    if (i == set_year)
                    {
                        //c2.CellsNum = new List<int> { 2, 2, 2 };
                        //c2.CellsObj = new ArrayList() { "", "95 ~ " + i + " 年度統計/" + lawMasterReport.DateNow, lawMasterReport.DateHave };
                        ExcelSetCell(sheet, new string[] { "" }, s, 1);
                        ExcelSetCell(sheet, new string[] { "" }, s, 2);
                        sheet.Cells[s, 1, s, 2].Merge = true;
                        sheet.Cells[s, 1, s, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { "95 ~ " + i + " 年度統計/" + lawMasterReport.DateNow }, s, 3);
                        ExcelSetCell(sheet, new string[] { "" }, s, 4);
                        sheet.Cells[s, 3, s, 4].Merge = true;
                        sheet.Cells[s, 3, s, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { lawMasterReport.DateHave }, s, 5);
                        ExcelSetCell(sheet, new string[] { "" }, s, 6);
                        sheet.Cells[s, 5, s, 6].Merge = true;
                        sheet.Cells[s, 5, s, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        decimal decimalValue = Convert.ToDecimal(reportLog.LawTotalRepayment) / Convert.ToDecimal(reportLog.LawTotalDue) * 100;
                        repay_per_str = Math.Round(decimalValue, 2) + "%";
                    }
                    else
                    {
                        //c2.CellsNum = new List<int> { 2, 2, 2 };
                        //c2.CellsObj = new ArrayList() { "", i + " 年度統計/" + lawMasterReport.DateNow, lawMasterReport.DateHave };
                        ExcelSetCell(sheet, new string[] { "" }, s, 1);
                        ExcelSetCell(sheet, new string[] { "" }, s, 2);
                        sheet.Cells[s, 1, s, 2].Merge = true;
                        sheet.Cells[s, 1, s, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { i + " 年度統計/" + lawMasterReport.DateNow }, s, 3);
                        ExcelSetCell(sheet, new string[] { "" }, s, 4);
                        sheet.Cells[s, 3, s, 4].Merge = true;
                        sheet.Cells[s, 3, s, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { lawMasterReport.DateHave }, s, 5);
                        ExcelSetCell(sheet, new string[] { "" }, s, 6);
                        sheet.Cells[s, 5, s, 6].Merge = true;
                        sheet.Cells[s, 5, s, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        repay_per_str = reportLog.LawRepayPercent;
                    }

                }
                else
                {
                    //c2.CellsNum = new List<int> { 2, 2, 2 };
                    //c2.CellsObj = new ArrayList() { "", i + " 年度統計/" + lawMasterReport.DateNow, "" };
                    ExcelSetCell(sheet, new string[] { "" }, s, 1);
                    ExcelSetCell(sheet, new string[] { "" }, s, 2);
                    sheet.Cells[s, 1, s, 2].Merge = true;
                    sheet.Cells[s, 1, s, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    ExcelSetCell(sheet, new string[] { i + " 年度統計/" + lawMasterReport.DateNow }, s, 3);
                    ExcelSetCell(sheet, new string[] { "" }, s, 4);
                    sheet.Cells[s, 3, s, 4].Merge = true;
                    sheet.Cells[s, 3, s, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    ExcelSetCell(sheet, new string[] { lawMasterReport.DateHave }, s, 5);
                    ExcelSetCell(sheet, new string[] { "" }, s, 6);
                    sheet.Cells[s, 5, s, 6].Merge = true;
                    sheet.Cells[s, 5, s, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    repay_per_str = "";
                }
                s = s + 5;

                if (i == set_year)
                {
                    sql1 = @"select isnull(sum(law_due_money),0) as law_due_money from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null";
                    content = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql1).FirstOrDefault();
                }
                else
                {
                    sql1 = @"select isnull(sum(law_due_money),0) as law_due_money from law_content where law_year=" + i + " and law_content_cancel_type is null";
                    content = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql1).FirstOrDefault();
                }

                if (content != null)
                {
                    if (reportLog != null)
                    {
                        //c3.CellsObj = new ArrayList() { "結欠總金額", content.LawDueMoney, reportLog.LawTotalDue };
                        ExcelSetCell(sheet, new string[] { "結欠總金額" }, d, 1);
                        ExcelSetCell(sheet, new string[] { "" }, d, 2);
                        sheet.Cells[d, 1, d, 2].Merge = true;
                        sheet.Cells[d, 1, d, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { content.LawDueMoney.ToString("###,###") }, d, 3);
                        ExcelSetCell(sheet, new string[] { "" }, d, 4);
                        sheet.Cells[d, 3, d, 4].Merge = true;
                        sheet.Cells[d, 3, d, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { reportLog.LawTotalDue.ToString("###,###") }, d, 5);
                        ExcelSetCell(sheet, new string[] { "" }, d, 6);
                        sheet.Cells[d, 5, d, 6].Merge = true;
                        sheet.Cells[d, 5, d, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    else
                    {
                        //c3.CellsObj = new ArrayList() { "結欠總金額", content.LawDueMoney, "" };
                        ExcelSetCell(sheet, new string[] { "結欠總金額" }, d, 1);
                        ExcelSetCell(sheet, new string[] { "" }, d, 2);
                        sheet.Cells[d, 1, d, 2].Merge = true;
                        sheet.Cells[d, 1, d, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { content.LawDueMoney.ToString("###,###") }, d, 3);
                        ExcelSetCell(sheet, new string[] { "" }, d, 4);
                        sheet.Cells[d, 3, d, 4].Merge = true;
                        sheet.Cells[d, 3, d, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { "" }, d, 5);
                        ExcelSetCell(sheet, new string[] { "" }, d, 6);
                        sheet.Cells[d, 5, d, 6].Merge = true;
                        sheet.Cells[d, 5, d, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    due_money = content.LawDueMoney;
                }
                else
                {
                    //c3.CellsObj = new ArrayList() { "結欠總金額", 0, 0 };
                    ExcelSetCell(sheet, new string[] { "結欠總金額" }, d, 1);
                    ExcelSetCell(sheet, new string[] { "" }, d, 2);
                    sheet.Cells[d, 1, d, 2].Merge = true;
                    sheet.Cells[d, 1, d, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    ExcelSetCell(sheet, new string[] { "0" }, d, 3);
                    ExcelSetCell(sheet, new string[] { "" }, d, 4);
                    sheet.Cells[d, 3, d, 4].Merge = true;
                    sheet.Cells[d, 3, d, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    ExcelSetCell(sheet, new string[] { "0" }, d, 5);
                    ExcelSetCell(sheet, new string[] { "" }, d, 6);
                    sheet.Cells[d, 5, d, 6].Merge = true;
                    sheet.Cells[d, 5, d, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    due_money = 0;
                }
                d = d + 5;

                if (i == set_year)
                {
                    sql2 = @"select (aa.law_repayment_money+bb.law_repayment_money+cc.law_repayment_money) as law_repayment_money from (select isnull(a.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null and law_status_type<>2 ) ) as a) as aa,(select isnull(b.law_repayment_money,0) as law_repayment_money from (select sum(law_due_money)as law_repayment_money from law_report_reference where law_year in(" + year_str + ") and law_note_no in(select law_note_no from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) not in(select close_type_id from law_close_type where count_type=0 ))) as b ) as bb,(select isnull(c.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year in(" + year_str + ") and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) in(select close_type_id from law_close_type where count_type=0 )) ) as c) as cc ";
                    lawRepaymentList = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql2).FirstOrDefault();
                }
                else
                {
                    sql2 = @"select (aa.law_repayment_money+bb.law_repayment_money+cc.law_repayment_money) as law_repayment_money from (select isnull(a.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year=" + i + " and law_content_cancel_type is null and law_status_type<>2) ) as a) as aa,(select isnull(b.law_repayment_money,0) as law_repayment_money from (select sum(law_due_money)as law_repayment_money from law_report_reference where law_year=" + i + " and law_note_no in(select law_note_no from law_content where law_year=" + i + " and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) not in(select close_type_id from law_close_type where count_type=0 ))) as b ) as bb,(select isnull(c.law_repayment_money,0) as law_repayment_money from (select sum(law_repayment_capital)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in (select law_id from law_content where law_year=" + i + " and law_content_cancel_type is null and law_status_type=2 and left(law_close_type,1) in(select close_type_id from law_close_type where count_type=0 )) ) as c) as cc";
                    lawRepaymentList = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql2).FirstOrDefault();
                }

                if (lawRepaymentList != null)
                {
                    if (reportLog != null)
                    {
                        //c4.CellsObj = new ArrayList() { "已清償本金累計金額", lawRepaymentList.LawRepaymentMoney, reportLog.LawTotalRepayment };
                        ExcelSetCell(sheet, new string[] { "已清償本金累計金額" }, z, 1);
                        ExcelSetCell(sheet, new string[] { "" }, z, 2);
                        sheet.Cells[z, 1, z, 2].Merge = true;
                        sheet.Cells[z, 1, z, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { lawRepaymentList.LawRepaymentMoney.ToString("###,###") }, z, 3);
                        ExcelSetCell(sheet, new string[] { "" }, z, 4);
                        sheet.Cells[z, 3, z, 4].Merge = true;
                        sheet.Cells[z, 3, z, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { reportLog.LawTotalRepayment.ToString("###,###") }, z, 5);
                        ExcelSetCell(sheet, new string[] { "" }, z, 6);
                        sheet.Cells[z, 5, z, 6].Merge = true;
                        sheet.Cells[z, 5, z, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    else
                    {
                        //c4.CellsObj = new ArrayList() { "已清償本金累計金額", lawRepaymentList.LawRepaymentMoney, "" };
                        ExcelSetCell(sheet, new string[] { "已清償本金累計金額" }, z, 1);
                        ExcelSetCell(sheet, new string[] { "" }, z, 2);
                        sheet.Cells[z, 1, z, 2].Merge = true;
                        sheet.Cells[z, 1, z, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { lawRepaymentList.LawRepaymentMoney.ToString("###,###") }, z, 3);
                        ExcelSetCell(sheet, new string[] { "" }, z, 4);
                        sheet.Cells[z, 3, z, 4].Merge = true;
                        sheet.Cells[z, 3, z, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { "" }, z, 5);
                        ExcelSetCell(sheet, new string[] { "" }, z, 6);
                        sheet.Cells[z, 5, z, 6].Merge = true;
                        sheet.Cells[z, 5, z, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    repayment_money = lawRepaymentList.LawRepaymentMoney;
                }
                else
                {
                    //c4.CellsObj = new ArrayList() { "已清償本金累計金額", 0, 0 };
                    ExcelSetCell(sheet, new string[] { "已清償本金累計金額" }, z, 1);
                    ExcelSetCell(sheet, new string[] { "" }, z, 2);
                    sheet.Cells[z, 1, z, 2].Merge = true;
                    sheet.Cells[z, 1, z, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    ExcelSetCell(sheet, new string[] { "0" }, z, 3);
                    ExcelSetCell(sheet, new string[] { "" }, z, 4);
                    sheet.Cells[z, 3, z, 4].Merge = true;
                    sheet.Cells[z, 3, z, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    ExcelSetCell(sheet, new string[] { "0" }, z, 5);
                    ExcelSetCell(sheet, new string[] { "" }, z, 6);
                    sheet.Cells[z, 5, z, 6].Merge = true;
                    sheet.Cells[z, 5, z, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    repayment_money = 0;
                }
                z = z + 5;

                if (due_money != 0 && repayment_money != 0)
                {
                    decimal decimalValue = repayment_money / due_money * 100;
                    p_str = Math.Round(decimalValue, 2) + "%";
                }
                else
                {
                    if (due_money != 0 && repayment_money == 0)
                    {
                        p_str = "0%";
                    }
                    else
                    {
                        p_str = "";
                    }
                }

                ExcelSetCell(sheet, new string[] { "達成率【金額】" }, g, 1);
                ExcelSetCell(sheet, new string[] { "" }, g, 2);
                sheet.Cells[g, 1, g, 2].Merge = true;
                sheet.Cells[g, 1, g, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { p_str }, g, 3);
                ExcelSetCell(sheet, new string[] { "" }, g, 4);
                sheet.Cells[g, 3, g, 4].Merge = true;
                sheet.Cells[g, 3, g, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[g, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[g, 3].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFF66"));
                sheet.Cells[g, 3].Style.Font.Color.SetColor(Color.Blue);
                //ExcelSetCell(sheet, new string[] { item.RepayMoney == "0" ? "0" : Convert.ToInt32(item.RepayMoney).ToString("###,###") }, g, w);

                ExcelSetCell(sheet, new string[] { repay_per_str }, g, 5);
                ExcelSetCell(sheet, new string[] { "" }, g, 6);
                sheet.Cells[g, 5, g, 6].Merge = true;
                sheet.Cells[g, 5, g, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[g, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[g, 5].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFF66"));
                sheet.Cells[g, 5].Style.Font.Color.SetColor(Color.Blue);
                g = g + 5;

                ExcelSetCell(sheet, new string[] { "" }, p, 1);
                ExcelSetCell(sheet, new string[] { "" }, p, 2);
                ExcelSetCell(sheet, new string[] { "" }, p, 3);
                ExcelSetCell(sheet, new string[] { "" }, p, 4);
                ExcelSetCell(sheet, new string[] { "" }, p, 5);
                ExcelSetCell(sheet, new string[] { "" }, p, 6);
                sheet.Cells[p, 1, p, 6].Merge = true;
                p = p + 5;
            }


            //sheet 標題 橫排 直排
            ExcelSetCell(sheet, new string[] { "總表" }, 1, 1);
            ExcelSetCell(sheet, new string[] { "" }, 1, 2);
            sheet.Cells[1, 1, 1, 2].Merge = true;
            sheet.Cells[1, 1, 1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //sheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 1].Style.Font.Color.SetColor(Color.Blue);

            ExcelSetCell(sheet, new string[] { lawMasterReport.DateNow }, 1, 3);
            ExcelSetCell(sheet, new string[] { "" }, 1, 4);
            sheet.Cells[1, 3, 1, 4].Merge = true;
            sheet.Cells[1, 3, 1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //sheet.Cells[1, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 3].Style.Font.Color.SetColor(Color.Blue);

            ExcelSetCell(sheet, new string[] { "(含續佣抵結欠)" }, 1, 5);
            ExcelSetCell(sheet, new string[] { "" }, 1, 6);
            sheet.Cells[1, 5, 1, 6].Merge = true;
            sheet.Cells[1, 5, 1, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //sheet.Cells[1, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 5].Style.Font.Color.SetColor(Color.Blue);

            sheet.Column(1).Width = 22;
            sheet.Cells.Style.ShrinkToFit = true;
            //字型
            sheet.Cells.Style.Font.Name = "新細明體";
            //文字大小
            sheet.Cells.Style.Font.Size = 12;
            excel.SaveAs(ms);
            excel.Dispose();
            ms.Position = 0;

            return ms;
        }
        #endregion

        #endregion
        #endregion

        #region 當月還款明細
        /// <summary>
        /// 當月還款明細
        /// </summary>
        public List<LawMonthRepaymentReportDetail> GetLawMonthRepaymentReport(LawMonthRepaymentReportDetail model)
        {
            List<LawMonthRepaymentReportDetail> result = new List<LawMonthRepaymentReportDetail>();
            string sql = @"select a.law_id,a.law_due_agentid,a.law_note_no,a.law_repayment_money,b.name,d.vm_code,d.vm_name,d.sm_code,d.sm_name from (";
            sql += " select law_id,law_due_agentid,law_note_no,sum(law_repayment_money)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list where law_id in(";
            sql += " select law_id from law_content where (law_content_cancel_type<>1 or law_content_cancel_type is null)) ";
            if (model.chkm == "1")
            {
                sql += " and (left(convert(varchar(50), law_repayment_date, 111), 7) = @selyearmonth or left(convert(varchar(50), create_date, 111), 7) = @selyearmonth) ";
            }
            else
            {
                sql += " and (left(convert(varchar(50),law_repayment_date,111),7) = @selyearmonth)";
            }
            sql += " group by law_id,law_due_agentid,law_note_no) a ";
            sql += " left join law_agent_data b on left(a.law_due_agentid,10)=b.agent_code ";
            sql += " left join law_report_vm_reference d on a.law_due_agentid=d.agent_id and a.law_id=d.law_id ";

            result = DbHelper.Query<LawMonthRepaymentReportDetail>(H2ORepository.ConnectionStringName, sql, new { selyearmonth = model.selyear + "/" + model.selmonth }).ToList();
            return result;
        }

        /// <summary>
        /// 當月還款明細總計
        /// </summary>
        public LawRepaymentList GetSumLawRepaymentMoney(LawMonthRepaymentReportDetail model)
        {
            LawRepaymentList result = new LawRepaymentList();
            string sql;
            string ymstr = model.selyear + "/" + model.selmonth;
            if (model.chkm == "1")
            {
                sql = "select sum(law_repayment_money)+sum(law_comm_deduction) as law_repayment_money from law_repayment_list a where (left(convert(varchar(50),law_repayment_date,111),7)=@ymstr or left(convert(varchar(50),create_date,111),7)=@ymstr)";
            }
            else
            {
                sql = "select sum(law_repayment_money) + sum(law_comm_deduction) as law_repayment_money from law_repayment_list a where left(convert(varchar(50), law_repayment_date, 111), 7) = @ymstr";
            }
            result = DbHelper.Query<LawRepaymentList>(H2ORepository.ConnectionStringName, sql, new { ymstr = ymstr }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="year">年度</param>
        /// <param name="month">月份</param>
        /// <returns>Stream</returns>
        public Stream QueryLawMonthRepaymentReport(string year, string month, string chkm)
        {
            MemoryStream ms = new MemoryStream();
            ExcelPackage excel = new ExcelPackage();
            ExcelWorksheet sheet = excel.Workbook.Worksheets.Add(string.Format("法追系統-當月({0}年{1}月)還款明細(含續佣抵結欠)", Convert.ToInt32(year) - 1911, month));
            LawMonthRepaymentReportDetail detail = new LawMonthRepaymentReportDetail();
            detail.chkm = chkm;
            detail.selmonth = month;
            detail.selyear = year;
            List<LawMonthRepaymentReportDetail> lawMonthRepaymentReportDetails = GetLawMonthRepaymentReport(detail);
            //sheet.Cells.Style.ShrinkToFit = true;

            ExcelSetCell(sheet, new string[] { "序號" }, 1, 1);
            sheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "團隊" }, 1, 2);
            sheet.Cells[1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "協理系" }, 1, 3);
            sheet.Cells[1, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "業務員" }, 1, 4);
            sheet.Cells[1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "追回金額" }, 1, 5);
            sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "總計" }, 1, 6);
            sheet.Cells[1, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            int SN = 1;
            var colPos = 2;
            for (int i = 0; i < lawMonthRepaymentReportDetails.Count; i++)
            {
                ExcelSetCell(sheet, new int[] { SN }, colPos, 1);
                sheet.Cells[colPos, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { lawMonthRepaymentReportDetails[i].vmname }, colPos, 2);
                sheet.Cells[colPos, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { lawMonthRepaymentReportDetails[i].smname }, colPos, 3);
                sheet.Cells[colPos, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ExcelSetCell(sheet, new string[] { lawMonthRepaymentReportDetails[i].name }, colPos, 4);
                sheet.Cells[colPos, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { lawMonthRepaymentReportDetails[i].LawRepaymentMoney.ToString("###,###") }, colPos, 5);
                ExcelSetCell(sheet, new string[] { "" }, colPos, 6);
                sheet.Cells[colPos, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                colPos++;
                SN++;
            }

            LawRepaymentList lawRepayment = GetSumLawRepaymentMoney(detail);
            int LawSumRepaymentMoney = lawRepayment != null ? lawRepayment.LawRepaymentMoney : 0;

            ExcelSetCell(sheet, new string[] { LawSumRepaymentMoney.ToString("###,###") }, 2, 6);
            sheet.Cells[2, 6, SN, 6].Merge = true;
            sheet.Cells[2, 6, SN, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; //水平致中
            sheet.Cells[2, 6, SN, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center; //垂直致中                                                                                                    
                                                                                              //字型
            sheet.Cells.Style.Font.Name = "新細明體";
            //文字大小
            sheet.Cells.Style.Font.Size = 12;
            excel.SaveAs(ms);
            excel.Dispose();
            ms.Position = 0;
            return ms;

            return null;


        }
        #endregion

        #region 團隊明細
        /// <summary>
        /// 團隊明細
        /// </summary>
        public List<LawVmSmReport> GetLawVmSmReport(LawVmSmReport model)
        {
            List<LawVmSmReport> result = new List<LawVmSmReport>();
            string sql = @"select vm_name,sm_name from law_vm_sm_report where law_year=@lawyear";

            result = DbHelper.Query<LawVmSmReport>(H2ORepository.ConnectionStringName, sql, new { lawyear = model.LawYear }).ToList();
            return result;
        }

        /// <summary>
        /// 團隊明細年度金額加總
        /// </summary>
        public LawVmSmDetail GetSumLawVmSmReportMoney(LawVmSmReport model, string year_str)
        {
            LawVmSmDetail result = new LawVmSmDetail();
            string sql = @"select sum(due_money) as due_money,sum(repay_money) as repay_money from law_vm_sm_report where law_year in(" + year_str + ") and vm_name= @vmname and sm_name= @smname";

            result = DbHelper.Query<LawVmSmDetail>(H2ORepository.ConnectionStringName, sql, new { vmname = model.VmName, smname = model.SmName }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 團隊明細年度金額
        /// </summary>
        public LawVmSmDetail GetLawVmSmReportMoney(LawVmSmReport model)
        {
            LawVmSmDetail result = new LawVmSmDetail();
            string sql = @"select * from law_vm_sm_report where vm_name=@VmName and sm_name=@SmName and law_year=@LawYear order by law_year";

            result = DbHelper.Query<LawVmSmDetail>(H2ORepository.ConnectionStringName, sql, new { VmName = model.VmName, SmName = model.SmName, LawYear = model.LawYear }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="year">年度</param>
        /// <returns>Stream</returns>
        public Stream QueryTeamReport(string year)
        {
            MemoryStream ms = new MemoryStream();
            ExcelPackage excel = new ExcelPackage();
            ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("團隊明細報表");
            LawVmSmReport detail = new LawVmSmReport();
            detail.LawYear = year;
            //取得團隊
            List<LawVmSmReport> lawVmSmReports = GetLawVmSmReport(detail);
            List<LawVmSmDetail> lawVmSmDetails = new List<LawVmSmDetail>();
            List<LawVmSmDetail> lawVmSmDetail_temp = new List<LawVmSmDetail>();
            int set_year, y_str, unitExtends = 0, p = 4, c = 2, d = 3, flag = 0, g = 5, flag_temp = 0;
            string year_str = string.Empty, yeargo, vm_tmp = string.Empty, sm_tmp = string.Empty, DueMoney = string.Empty, RepayMoney = string.Empty, pstr = string.Empty;
            set_year = Convert.ToInt32(year);
            y_str = DateTime.Now.Year - 1911;
            for (int i = set_year; i <= y_str; i++)
            {
                unitExtends = unitExtends + 3;
                if (set_year == 95)
                {
                    yeargo = i + "年度";
                }
                else if (set_year > 95 && flag == 0)
                {
                    flag = flag + 1;
                    yeargo = "95 ~ " + i + "年度";
                }
                else
                {
                    yeargo = i + "年度";
                }

                ExcelSetCell(sheet, new string[] { yeargo }, 2, c);
                ExcelSetCell(sheet, new string[] { "" }, 2, d);
                ExcelSetCell(sheet, new string[] { "" }, 2, p);
                sheet.Cells[2, c, 2, p].Merge = true;
                sheet.Cells[2, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { "發函金額" }, 3, c);
                ExcelSetCell(sheet, new string[] { "" }, 4, c);
                sheet.Cells[3, c, 4, c].Merge = true;
                sheet.Cells[3, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { "追回本" + "\n" + "金金額" }, 3, d);
                ExcelSetCell(sheet, new string[] { "" }, 4, d);
                sheet.Cells[3, d, 4, d].Merge = true;
                sheet.Cells[3, d].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { "百分比" }, 3, p);
                ExcelSetCell(sheet, new string[] { "" }, 4, p);
                sheet.Cells[3, p, 4, p].Merge = true;
                sheet.Cells[3, p].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                d = d + 3;
                c = c + 3;
                p = p + 3;
            }
            for (int i = 95; i <= set_year; i++)
            {
                year_str = i == 95 ? "95" : year_str + "," + i;
            }
            LawVmSmReport last = lawVmSmReports.Last();

            foreach (LawVmSmReport lawVmSmReport in lawVmSmReports)
            {
                int y = 2, t = 3, r = 4;
                if (vm_tmp != lawVmSmReport.VmName && vm_tmp != "公司" && vm_tmp != "")
                {
                    ExcelSetCell(sheet, new string[] { sm_tmp }, g, 1);
                    sheet.Cells[g, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[g, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFCC99"));
                    for (int a = 0; a < lawVmSmDetail_temp.Count; a++)
                    {
                        //團隊小計資料
                        ExcelSetCell(sheet, new string[] { lawVmSmDetail_temp[a].DueMoney }, g, y);
                        sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { lawVmSmDetail_temp[a].RepayMoney }, g, t);
                        sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { lawVmSmDetail_temp[a].pstr }, g, r);
                        sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[g, y].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[g, t].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[g, r].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[g, y].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFCC99"));
                        sheet.Cells[g, t].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFCC99"));
                        sheet.Cells[g, r].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFCC99"));
                        if (!String.IsNullOrEmpty(lawVmSmDetail_temp[a].pstr))
                        {
                            sheet.Cells[g, r].Style.Font.Color.SetColor(lawVmSmDetail_temp[a].dv > 60 ? Color.Blue : Color.Red);
                        }
                        y = y + 3;
                        t = t + 3;
                        r = r + 3;

                    }
                    lawVmSmDetail_temp.Clear();
                }
                //團隊
                if (lawVmSmReport.SmName == lawVmSmReport.VmName && lawVmSmReport.SmName == "公司")
                {
                    ExcelSetCell(sheet, new string[] { lawVmSmReport.SmName + "總計" }, g, 1);
                    sheet.Cells[g, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[g, 1].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                }
                else
                {
                    if (lawVmSmReport.VmName == lawVmSmReport.SmName && (lawVmSmReport.VmName.Substring(lawVmSmReport.VmName.Length - 2, 2) == "團隊" && lawVmSmReport.SmName.Substring(lawVmSmReport.SmName.Length - 2, 2) == "團隊") || (lawVmSmReport.VmName != lawVmSmReport.SmName && (lawVmSmReport.VmName.Substring(lawVmSmReport.VmName.Length - 2, 2) == "團隊" && lawVmSmReport.SmName.Substring(lawVmSmReport.SmName.Length - 2, 2) == "團隊")))
                    {
                        vm_tmp = lawVmSmReport.VmName;
                        if (lawVmSmReport.VmName == lawVmSmReport.SmName && (lawVmSmReport.VmName.Substring(lawVmSmReport.VmName.Length - 2, 2) == "團隊" && lawVmSmReport.SmName.Substring(lawVmSmReport.SmName.Length - 2, 2) == "團隊"))
                        {
                            sm_tmp = lawVmSmReport.SmName + "小計";
                            //ExcelSetCell(sheet, new string[] { lawVmSmReport.SmName + "小計" }, g, 1);
                        }
                        else
                        {
                            ExcelSetCell(sheet, new string[] { lawVmSmReport.VmName + "小計" }, g, 1);
                            sheet.Cells[g, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[g, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFCC99"));
                        }
                    }
                    else
                    {
                        ExcelSetCell(sheet, new string[] { lawVmSmReport.SmName }, g, 1);
                    }
                }
                sheet.Cells[g, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //sheet.Cells[g, 1].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                int v = 2, w = 3, x = 4;
                //decimal dv = 0; 
                //依照年分取得團隊資料
                for (int i = set_year; i <= y_str; i++)
                {
                    LawVmSmDetail item = new LawVmSmDetail();
                    if (i == set_year)
                    {
                        item = GetSumLawVmSmReportMoney(lawVmSmReport, year_str);
                    }
                    else
                    {
                        lawVmSmReport.LawYear = i.ToString();
                        item = GetLawVmSmReportMoney(lawVmSmReport);
                    }

                    if (item.RepayMoney == "0")
                    {
                        item.pstr = string.Empty;
                    }
                    else if (lawVmSmReport.VmName == lawVmSmReport.SmName)
                    {
                        decimal decimalValue = Convert.ToDecimal(item.RepayMoney) / Convert.ToDecimal(item.DueMoney) * 100;
                        item.dv = decimalValue;
                        item.pstr = Math.Round(decimalValue, 2) + "%";
                    }
                    else
                    {
                        item.pstr = string.Empty;
                    }

                    if (lawVmSmReport.VmName == lawVmSmReport.SmName && (lawVmSmReport.VmName.Substring(lawVmSmReport.VmName.Length - 2, 2) == "團隊" && lawVmSmReport.SmName.Substring(lawVmSmReport.SmName.Length - 2, 2) == "團隊") || (lawVmSmReport.VmName != lawVmSmReport.SmName && (lawVmSmReport.VmName.Substring(lawVmSmReport.VmName.Length - 2, 2) == "團隊" && lawVmSmReport.SmName.Substring(lawVmSmReport.SmName.Length - 2, 2) == "團隊")))
                    {
                        //vm_tmp = lawVmSmReport.VmName;
                        if (lawVmSmReport.VmName == lawVmSmReport.SmName && (lawVmSmReport.VmName.Substring(lawVmSmReport.VmName.Length - 2, 2) == "團隊" && lawVmSmReport.SmName.Substring(lawVmSmReport.SmName.Length - 2, 2) == "團隊"))
                        {
                            LawVmSmDetail item_temp = new LawVmSmDetail();
                            item_temp.DueMoney = item.DueMoney == "0" ? "0" : Convert.ToInt32(item.DueMoney).ToString("###,###");
                            item_temp.RepayMoney = item.RepayMoney == "0" ? "0" : Convert.ToInt32(item.RepayMoney).ToString("###,###");
                            item_temp.pstr = item.pstr;
                            item_temp.dv = item.dv;
                            lawVmSmDetail_temp.Add(item_temp);
                        }
                        else
                        {
                            //ExcelSetCell(sheet, new string[] { lawVmSmReport.VmName + "小計" }, g, 1);
                        }
                    }
                    else
                    {
                        ExcelSetCell(sheet, new string[] { item.DueMoney == "0" ? "0" : Convert.ToInt32(item.DueMoney).ToString("###,###") }, g, v);
                        sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { item.RepayMoney == "0" ? "0" : Convert.ToInt32(item.RepayMoney).ToString("###,###") }, g, w);
                        sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { item.pstr }, g, x);
                        sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        if (lawVmSmReport.SmName == lawVmSmReport.VmName && lawVmSmReport.SmName == "公司")
                        {
                            sheet.Cells[g, v].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[g, w].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[g, x].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[g, v].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                            sheet.Cells[g, w].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                            sheet.Cells[g, x].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                            if (!String.IsNullOrEmpty(item.pstr))
                            {
                                sheet.Cells[g, x].Style.Font.Color.SetColor(item.dv > 60 ? Color.Blue : Color.Red);
                            }
                        }
                    }

                    //ExcelSetCell(sheet, new string[] { item.DueMoney }, g, v);
                    //sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //ExcelSetCell(sheet, new string[] { item.RepayMoney }, g, w);
                    //sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //ExcelSetCell(sheet, new string[] { item.pstr }, g, x);
                    //sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    v = v + 3;
                    w = w + 3;
                    x = x + 3;
                }
                g++;
                if (g == 7 && flag_temp == 0)
                {
                    g = 6;
                    flag_temp = 1;
                }

                if (lawVmSmReport.Equals(last))
                {
                    ExcelSetCell(sheet, new string[] { sm_tmp }, g, 1);
                    sheet.Cells[g, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[g, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[g, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFCC99"));
                    int h = 2, k = 3, l = 4;
                    for (int a = 0; a < lawVmSmDetail_temp.Count; a++)
                    {
                        ExcelSetCell(sheet, new string[] { lawVmSmDetail_temp[a].DueMoney }, g, h);
                        sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { lawVmSmDetail_temp[a].RepayMoney }, g, k);
                        sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { lawVmSmDetail_temp[a].pstr }, g, l);
                        sheet.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[g, h].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[g, k].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[g, l].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[g, h].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFCC99"));
                        sheet.Cells[g, k].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFCC99"));
                        sheet.Cells[g, l].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFCC99"));
                        if (!String.IsNullOrEmpty(lawVmSmDetail_temp[a].pstr))
                        {
                            sheet.Cells[g, l].Style.Font.Color.SetColor(lawVmSmDetail_temp[a].dv > 60 ? Color.Blue : Color.Red);
                        }
                        h = h + 3;
                        k = k + 3;
                        l = l + 3;
                    }
                }
            }
            // 報表標題
            var colEnd = unitExtends + 1;
            //sheet 標題 橫排 直排
            ExcelSetCell(sheet, new string[] { "團隊明細報表" }, 1, 1);
            ExcelSetCell(sheet, new string[] { "" }, 1, colEnd);
            sheet.Cells[1, 1, 1, colEnd].Merge = true;
            sheet.Cells[1, 1, 1, colEnd].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "單位" }, 2, 1);
            sheet.Cells[2, 1, 4, 1].Merge = true;
            sheet.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            sheet.Column(1).Width = 22;
            sheet.Cells.Style.ShrinkToFit = true;
            //字型
            sheet.Cells.Style.Font.Name = "新細明體";
            //文字大小
            sheet.Cells.Style.Font.Size = 12;
            excel.SaveAs(ms);
            excel.Dispose();
            ms.Position = 0;

            return ms;
        }
        #endregion

        #region 95-98案件統計
        /// <summary>
        /// 上月總累計未結案件
        /// </summary>
        public List<LawContent> GetLawTopYearReport()
        {
            List<LawContent> result = new List<LawContent>();
            int mstr = DateTime.Now.Month;
            string sql = @"select law_note_no from law_content where law_year between '95' and '97' and (law_content_cancel_type<>1 or law_content_cancel_type is null) and law_status_type<>2 
            union select law_note_no from law_content where law_year='98' and law_month not in(11,12) and (law_content_cancel_type<>1 or law_content_cancel_type is null) and law_status_type<>2
            union select law_note_no from law_content where right(left(convert(varchar(10),law_close_date,111),7),2)= @mstr";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mstr = mstr }).ToList();
            return result;
        }

        /// <summary>
        /// 本月結案
        /// </summary>
        public List<LawContent> GetLawNowYearReport(string mm)
        {
            List<LawContent> result = new List<LawContent>();

            string sql = @"select law_note_no from law_content where right(left(convert(varchar(10),law_close_date,111),7),2)= @mm  and law_status_type =2 and (law_content_cancel_type<>1 or law_content_cancel_type is null)";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mm = mm }).ToList();
            return result;
        }

        /// <summary>
        /// 合計未結總數
        /// </summary>
        public List<LawContent> GetLawSumNotTypeReport(string mm)
        {
            List<LawContent> result = new List<LawContent>();

            string sql = @"select law_note_no from law_content where law_year between '95' and '97' and (law_content_cancel_type<>1 or law_content_cancel_type is null) and law_status_type<>2 
            union select law_note_no from law_content where law_year='98' and law_month not in(11,12) and (law_content_cancel_type<>1 or law_content_cancel_type is null) and law_status_type<>2";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mm = mm }).ToList();
            return result;
        }

        /// <summary>
        /// 委外案件數
        /// </summary>
        public List<LawContent> GetLawOutSideReport(string mm)
        {
            List<LawContent> result = new List<LawContent>();
            string yyyymm = DateTime.Now.Year + "/" + mm;
            string sql = @"select law_note_no from law_content where (law_content_cancel_type<>1 or law_content_cancel_type is null) and law_do_unit_id<>(select law_do_unit_id from law_do_unit where unit_name='法務室') and law_year<='98' and law_id not in(select law_id from law_content where law_year=98 and left(law_content_create_date,4)>='2010')
            and law_id in (select law_id from law_do_unit_log where left(case_date,7) <= @yyyymm)";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { yyyymm = yyyymm }).ToList();
            return result;
        }

        /// <summary>
        /// 自辦案件數
        /// </summary>
        public List<LawContent> GetLawSelfReport(string mm)
        {
            List<LawContent> result = new List<LawContent>();
            string yyyymm = DateTime.Now.Year + "/" + mm;
            string sql = @"select law_note_no from law_content where (law_content_cancel_type<>1 or law_content_cancel_type is null) and law_do_unit_id=(select law_do_unit_id from law_do_unit where unit_name='法務室') and law_year<='98'
            and law_id not in(select law_id from law_content where law_year=98 and left(law_content_create_date,4)>='2010')
            union 
            select law_note_no from law_do_unit_log where left(case_date,7) > @yyyymm and law_do_unit_id<>(select law_do_unit_id from law_do_unit where unit_name='法務室')and law_id not in(select law_id from law_content where law_year=98 and left(law_content_create_date,4)>='2010') ";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { yyyymm = yyyymm }).ToList();
            return result;
        }

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="year">年度</param>
        /// <param name="month">月份</param>
        /// <returns>Stream</returns>
        public Stream QueryLawYearReport()
        {
            MemoryStream ms = new MemoryStream();
            ExcelPackage excel = new ExcelPackage();
            ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("95 - 98 年度案件統計表");
            string mm;
            int month = DateTime.Now.Month;
            int YearCol = 2;
            for (int i = 1; i <= 12; i++)
            {
                if (i.ToString().Length == 1)
                {
                    mm = "0" + i;
                }
                else
                {
                    mm = i.ToString();
                }

                ExcelSetCell(sheet, new string[] { i + "月" }, 2, YearCol);
                sheet.Cells[2, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if (i > month)
                {
                    List<LawContent> Top = GetLawTopYearReport();
                    ExcelSetCell(sheet, new string[] { "" }, 3, YearCol);
                    sheet.Cells[3, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    List<LawContent> Now = GetLawNowYearReport(mm);
                    ExcelSetCell(sheet, new string[] { "" }, 4, YearCol);
                    sheet.Cells[4, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    List<LawContent> Sum = GetLawSumNotTypeReport(mm);
                    ExcelSetCell(sheet, new string[] { "" }, 5, YearCol);
                    sheet.Cells[5, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    List<LawContent> Out = GetLawOutSideReport(mm);
                    ExcelSetCell(sheet, new string[] { "" }, 6, YearCol);
                    sheet.Cells[6, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    List<LawContent> Own = GetLawSelfReport(mm);
                    ExcelSetCell(sheet, new string[] { "" }, 7, YearCol);
                    sheet.Cells[7, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                else
                {
                    List<LawContent> Top = GetLawTopYearReport();
                    ExcelSetCell(sheet, new int[] { Top.Count }, 3, YearCol);
                    sheet.Cells[3, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    List<LawContent> Now = GetLawNowYearReport(mm);
                    ExcelSetCell(sheet, new int[] { Now.Count }, 4, YearCol);
                    sheet.Cells[4, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    List<LawContent> Sum = GetLawSumNotTypeReport(mm);
                    ExcelSetCell(sheet, new int[] { Sum.Count }, 5, YearCol);
                    sheet.Cells[5, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    List<LawContent> Out = GetLawOutSideReport(mm);
                    ExcelSetCell(sheet, new int[] { Out.Count }, 6, YearCol);
                    sheet.Cells[6, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    List<LawContent> Own = GetLawSelfReport(mm);
                    ExcelSetCell(sheet, new int[] { Own.Count }, 7, YearCol);
                    sheet.Cells[7, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                YearCol++;
            }

            ExcelSetCell(sheet, new string[] { "95 - 98 年度案件統計表" }, 1, 1);
            sheet.Cells[1, 1, 1, 13].Merge = true;
            sheet.Cells[1, 1, 1, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "月份" }, 2, 1);
            sheet.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "上月總累計未結案件" }, 3, 1);
            sheet.Cells[3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "本月結案" }, 4, 1);
            sheet.Cells[4, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "合計未結總數" }, 5, 1);
            sheet.Cells[5, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "委外案件數" }, 6, 1);
            sheet.Cells[6, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelSetCell(sheet, new string[] { "自辦案件數" }, 7, 1);
            sheet.Cells[7, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Column(1).Width = 22;
            excel.SaveAs(ms);
            excel.Dispose();
            ms.Position = 0;

            return ms;
        }

        #endregion

        #region 年度案件統計
        /// <summary>
        /// 受理案件
        /// </summary>
        public List<LawContent> GetLawAccpet(string year, string mm)
        {
            List<LawContent> result = new List<LawContent>();
            string mstr = year + "/" + mm;
            string sql = @"select law_note_no from law_content where left(convert(varchar(50),law_content_create_date,111),7)=@mstr and (law_content_cancel_type<>1 or law_content_cancel_type is null) 
                and law_note_no not in(select law_note_no from law_content where law_year between '95' and '97' union select law_note_no from law_content where law_year='98' and law_month not in(11,12))";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mstr = mstr }).ToList();
            return result;
        }

        /// <summary>
        /// 受理案件總計
        /// </summary>
        public List<LawContent> GetLawSumAccpet(string year)
        {
            List<LawContent> result = new List<LawContent>();
            string mstr = year;
            string sql = @"select law_note_no from law_content where left(convert(varchar(50),law_content_create_date,111),4)=@mstr and (law_content_cancel_type<>1 or law_content_cancel_type is null)  
                and law_note_no not in(select law_note_no from law_content where law_year between '95' and '97' union select law_note_no from law_content where law_year = '98' and law_month not in(11, 12)) ";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mstr = mstr }).ToList();
            return result;
        }

        /// <summary>
        /// 委外件數
        /// </summary>
        public List<LawContent> GetLawOutsource(string year, string mm)
        {
            List<LawContent> result = new List<LawContent>();
            string mstr = year + "/" + mm;
            string sql = @"select law_note_no from law_content where left(convert(varchar(50),law_content_create_date,111),7)=@mstr and (law_content_cancel_type<>1 or law_content_cancel_type is null) and law_do_unit_id<>(select law_do_unit_id from law_do_unit where unit_name='法務室') 
                and law_note_no not in(select law_note_no from law_content where law_year between '95' and '97' union select law_note_no from law_content where law_year = '98' and law_month not in(11, 12))";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mstr = mstr }).ToList();
            return result;
        }

        /// <summary>
        /// 委外件數總計
        /// </summary>
        public List<LawContent> GetLawSumOutsource(string year)
        {
            List<LawContent> result = new List<LawContent>();
            string mstr = year;
            string sql = @"select law_note_no from law_content where left(convert(varchar(50),law_content_create_date,111),4)=@mstr and (law_content_cancel_type<>1 or law_content_cancel_type is null) and law_do_unit_id<>(select law_do_unit_id from law_do_unit where unit_name='法務室')  
                and law_note_no not in(select law_note_no from law_content where law_year between '95' and '97' union select law_note_no from law_content where law_year = '98' and law_month not in(11, 12)) ";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mstr = mstr }).ToList();
            return result;
        }

        /// <summary>
        /// 自辦件數
        /// </summary>
        public List<LawContent> GetLawDoself(string year, string mm)
        {
            List<LawContent> result = new List<LawContent>();
            string mstr = year + "/" + mm;
            string sql = @"select law_note_no from law_content where left(convert(varchar(50),law_content_create_date,111),7)=@mstr and (law_content_cancel_type<>1 or law_content_cancel_type is null) and law_do_unit_id=(select law_do_unit_id from law_do_unit where unit_name='法務室') 
                and law_note_no not in(select law_note_no from law_content where law_year between '95' and '97' union select law_note_no from law_content where law_year = '98' and law_month not in(11, 12)) ";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mstr = mstr }).ToList();
            return result;
        }

        /// <summary>
        /// 自辦件數總計
        /// </summary>
        public List<LawContent> GetLawSumDoself(string year)
        {
            List<LawContent> result = new List<LawContent>();
            string mstr = year;
            string sql = @"select law_note_no from law_content where left(convert(varchar(50),law_content_create_date,111),4)=@mstr and (law_content_cancel_type<>1 or law_content_cancel_type is null) and law_do_unit_id=(select law_do_unit_id from law_do_unit where unit_name='法務室')
                and law_note_no not in(select law_note_no from law_content where law_year between '95' and '97' union select law_note_no from law_content where law_year = '98' and law_month not in(11, 12)) ";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mstr = mstr }).ToList();
            return result;
        }

        /// <summary>
        /// 未結
        /// </summary>
        public List<LawContent> GetLawNotOC(string year, string mm)
        {
            List<LawContent> result = new List<LawContent>();
            string mstr = year + "/" + mm;
            string sql = @"select law_note_no from law_content where left(convert(varchar(50),law_content_create_date,111),7)=@mstr and law_close_type is null and law_close_date is null and (law_content_cancel_type<>1 or law_content_cancel_type is null) 
                and law_note_no not in(select law_note_no from law_content where law_year between '95' and '97' union select law_note_no from law_content where law_year = '98' and law_month not in(11, 12)) ";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mstr = mstr }).ToList();
            return result;
        }

        /// <summary>
        /// 未結總計
        /// </summary>
        public List<LawContent> GetLawSumNotOC(string year)
        {
            List<LawContent> result = new List<LawContent>();
            string mstr = year;
            string sql = @"select law_note_no from law_content where left(convert(varchar(50),law_content_create_date,111),4)=@mstr and law_close_type is null and law_close_date is null and (law_content_cancel_type<>1 or law_content_cancel_type is null)
                and law_note_no not in(select law_note_no from law_content where law_year between '95' and '97' union select law_note_no from law_content where law_year = '98' and law_month not in(11, 12)) ";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mstr = mstr }).ToList();
            return result;
        }

        /// <summary>
        /// 結案
        /// </summary>
        public List<LawContent> GetLawOC(string year, string mm)
        {
            List<LawContent> result = new List<LawContent>();
            string mstr = year + "/" + mm;
            string sql = @"select law_note_no from law_content where left(convert(varchar(50),law_content_create_date,111),7)=@mstr and law_close_type is not null and law_close_date is not null and (law_content_cancel_type<>1 or law_content_cancel_type is null) 
                and law_note_no not in(select law_note_no from law_content where law_year between '95' and '97' union select law_note_no from law_content where law_year = '98' and law_month not in(11, 12)) ";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mstr = mstr }).ToList();
            return result;
        }

        /// <summary>
        /// 結案總計
        /// </summary>
        public List<LawContent> GetLawSumOC(string year)
        {
            List<LawContent> result = new List<LawContent>();
            string mstr = year;
            string sql = @"select law_note_no from law_content where left(convert(varchar(50),law_content_create_date,111),4)=@mstr and law_close_type is not null and law_close_date is not null and (law_content_cancel_type<>1 or law_content_cancel_type is null)
                and law_note_no not in(select law_note_no from law_content where law_year between '95' and '97' union select law_note_no from law_content where law_year = '98' and law_month not in(11, 12)) ";

            result = DbHelper.Query<LawContent>(H2ORepository.ConnectionStringName, sql, new { mstr = mstr }).ToList();
            return result;
        }

        /// <summary>
        /// 報表
        /// </summary>
        /// <param name="year">年度</param>
        /// <param name="month">月份</param>
        /// <returns>Stream</returns>
        public Stream QueryLawNowYearReport()
        {
            MemoryStream ms = new MemoryStream();
            ExcelPackage excel = new ExcelPackage();
            ExcelWorksheet sheet = excel.Workbook.Worksheets.Add("案件統計表");
            string mm;
            string md = DateTime.Now.ToString(".MM.dd");
            int month = DateTime.Now.Month;
            int year = DateTime.Now.Year;
            int a = 1, b = 2, c = 3, d = 4, e = 5, f = 6, g = 7;
            for (var xi = year; xi >= 2010; xi--)
            {
                int YearCol = 2;
                for (int i = 1; i <= 12; i++)
                {
                    if (i.ToString().Length == 1)
                    {
                        mm = "0" + i;
                    }
                    else
                    {
                        mm = i.ToString();
                    }
                    ExcelSetCell(sheet, new string[] { i + "月" }, b, YearCol);
                    sheet.Cells[b, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    if (i > month && xi == year)
                    {
                        ExcelSetCell(sheet, new string[] { "" }, c, YearCol);
                        sheet.Cells[c, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { "" }, d, YearCol);
                        sheet.Cells[d, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { "" }, e, YearCol);
                        sheet.Cells[e, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { "" }, f, YearCol);
                        sheet.Cells[f, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ExcelSetCell(sheet, new string[] { "" }, g, YearCol);
                        sheet.Cells[g, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    else
                    {
                        List<LawContent> Accpet = GetLawAccpet(xi.ToString(), mm);
                        ExcelSetCell(sheet, new int[] { Accpet.Count }, c, YearCol);
                        sheet.Cells[c, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        List<LawContent> Outsource = GetLawOutsource(xi.ToString(), mm);
                        ExcelSetCell(sheet, new int[] { Outsource.Count }, d, YearCol);
                        sheet.Cells[d, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        List<LawContent> Doself = GetLawDoself(xi.ToString(), mm);
                        ExcelSetCell(sheet, new int[] { Doself.Count }, e, YearCol);
                        sheet.Cells[e, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        List<LawContent> NotOC = GetLawNotOC(xi.ToString(), mm);
                        ExcelSetCell(sheet, new int[] { NotOC.Count }, f, YearCol);
                        sheet.Cells[f, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        List<LawContent> OC = GetLawOC(xi.ToString(), mm);
                        ExcelSetCell(sheet, new int[] { OC.Count }, g, YearCol);
                        sheet.Cells[g, YearCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    YearCol++;
                }

                ExcelSetCell(sheet, new string[] { (xi - 1911) + "年度案件統計表(計至" + (xi - 1911) + md + " 止)" }, a, 1);
                ExcelSetCell(sheet, new string[] { "" }, a, 2);
                ExcelSetCell(sheet, new string[] { "" }, a, 3);
                ExcelSetCell(sheet, new string[] { "" }, a, 4);
                ExcelSetCell(sheet, new string[] { "" }, a, 5);
                ExcelSetCell(sheet, new string[] { "" }, a, 6);
                ExcelSetCell(sheet, new string[] { "" }, a, 7);
                ExcelSetCell(sheet, new string[] { "" }, a, 8);
                ExcelSetCell(sheet, new string[] { "" }, a, 9);
                ExcelSetCell(sheet, new string[] { "" }, a, 10);
                ExcelSetCell(sheet, new string[] { "" }, a, 11);
                ExcelSetCell(sheet, new string[] { "" }, a, 12);
                ExcelSetCell(sheet, new string[] { "" }, a, 13);
                ExcelSetCell(sheet, new string[] { "" }, a, 14);
                sheet.Cells[a, 1, a, 14].Merge = true;
                sheet.Cells[a, 1, a, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { "月份" }, b, 1);
                sheet.Cells[b, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { "受理案件" }, c, 1);
                sheet.Cells[c, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { "委外件數" }, d, 1);
                sheet.Cells[d, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { "自辦件數" }, e, 1);
                sheet.Cells[e, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { "未結" }, f, 1);
                sheet.Cells[f, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { "結案" }, g, 1);
                sheet.Cells[g, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelSetCell(sheet, new string[] { "總計" }, b, 14);
                sheet.Cells[b, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                List<LawContent> SumAccpet = GetLawSumAccpet(xi.ToString());
                ExcelSetCell(sheet, new int[] { SumAccpet.Count }, c, 14);
                sheet.Cells[c, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                List<LawContent> SumOutsource = GetLawSumOutsource(xi.ToString());
                ExcelSetCell(sheet, new int[] { SumOutsource.Count }, d, 14);
                sheet.Cells[d, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                List<LawContent> SumDoself = GetLawSumDoself(xi.ToString());
                ExcelSetCell(sheet, new int[] { SumDoself.Count }, e, 14);
                sheet.Cells[e, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                List<LawContent> SumNotOC = GetLawSumNotOC(xi.ToString());
                ExcelSetCell(sheet, new int[] { SumNotOC.Count }, f, 14);
                sheet.Cells[f, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                List<LawContent> SumOC = GetLawSumOC(xi.ToString());
                ExcelSetCell(sheet, new int[] { SumOC.Count }, g, 14);
                sheet.Cells[g, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                a = a + 8;
                b = b + 8;
                c = c + 8;
                d = d + 8;
                e = e + 8;
                f = f + 8;
                g = g + 8;
            }
            sheet.Column(1).Width = 22;
            excel.SaveAs(ms);
            excel.Dispose();
            ms.Position = 0;

            return ms;
        }

        #endregion
        #endregion

        #region 電話催告通知
        /// <summary>
        /// 電話催告通知
        /// </summary>
        public List<LawPhoneCallLogDetail> GetLawPhoneCallLog()
        {
            List<LawPhoneCallLogDetail> result = new List<LawPhoneCallLogDetail>();

            string sql = @"select * from law_phone_call_log a left join law_content b on a.law_id=b.law_id where law_phone_call_read_log=0 and (b.law_content_cancel_type<>1 or b.law_content_cancel_type is null) and convert(varchar(50),getdate(),111)>=a.law_phone_call_limited_date  
                and law_call_log_id in (select max(law_call_log_id) from law_phone_call_log group by law_id)";

            result = DbHelper.Query<LawPhoneCallLogDetail>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 維護日期逾30日之案件列表通知
        /// </summary>
        public List<LawPhoneCallLogDetail> GetThirtyDayLawContent()
        {
            List<LawPhoneCallLogDetail> result = new List<LawPhoneCallLogDetail>();

            string sql = @"select law_id,law_note_no,law_content_lastchange_date,LawContentDouserID,law_douser_name from law_content where datediff(day,law_content_lastchange_date,convert(varchar(50),getdate(),111))>=30 and law_status_type<>2";

            result = DbHelper.Query<LawPhoneCallLogDetail>(H2ORepository.ConnectionStringName, sql).ToList();
            return result;
        }

        /// <summary>
        /// 更新催告通知
        /// </summary>
        public void UpdateLawPhoneCallLog(string PhoneCallReadID)
        {
            string ymd = DateTime.Now.ToString("yyyy/MM/dd");
            var sql = @"update law_phone_call_log 
            set law_phone_call_read_log=1,law_phone_call_readlog_date = @ymd,PhoneCallReadID = @PhoneCallReadID
            from law_phone_call_log a left join law_content b on a.law_id = b.law_id where law_phone_call_read_log = 0 and(b.law_content_cancel_type <> 1 or b.law_content_cancel_type is null) and convert(varchar(50),getdate(),111)>= a.law_phone_call_limited_date
            and law_call_log_id in (select max(law_call_log_id) from law_phone_call_log group by law_id)";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new
            { ymd = ymd, PhoneCallReadID = PhoneCallReadID });
        }
        #endregion

        #region 法追明細
        /// <summary>
        /// 未結案件列表
        /// </summary>
        public List<LawSearchDetail> GetLawSearchByNotClose(LawSearchDetail model, OrgVm orgVm)
        {
            List<LawSearchDetail> result = new List<LawSearchDetail>();
            org check = new org();
            string sql = string.Empty, sql1 = string.Empty;

            if (orgVm.vm_flag == 1)
            {
                if (orgVm.vsm_flag == 0)
                {
                    sql = @"select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where b.vm_code=@vmcode and law_status_type<>2 and (law_content_cancel_type is null or law_content_cancel_type<>1)";
                }
                else
                {
                    sql = @"select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where b.sm_code in (" + orgVm.smstr + ") and law_status_type<>2 and (law_content_cancel_type is null or law_content_cancel_type<>1)";
                }
            }
            else
            {
                sql1 = @"select * from org where manager=(select orgid from org where href>65 and login=@login and enabled='Y' and isdelete=0 )and right(orgname,2)='協理'";
                check = DbHelper.Query<org>(H2ORepository.ConnectionStringName, sql1, new
                {
                    login = orgVm.vmleaderid.Substring(0, 10),
                }).FirstOrDefault();

                if (check != null)
                {
                    sql = "select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where left(administrat_id,10)=@vmleaderid and law_status_type<>2 and (law_content_cancel_type is null or law_content_cancel_type<>1)";
				}
				else
				{
                    sql = "select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where left(administrat_id,10)=@vmleaderid and law_status_type<>2 and (law_content_cancel_type is null or law_content_cancel_type<>1)";
				}
            }

            if (!string.IsNullOrWhiteSpace(model.LawDueAgentId))
            {
                sql += " and left(law_due_agentid,10)= @LawDueAgentId";
            }
            if (!string.IsNullOrWhiteSpace(model.LawDueName))
            {
                sql += " and law_due_name like @LawDueName";
            }
            if (model.LawDueMoney != 0)
            {
                sql += " and law_due_money= @LawDueMoney";
            }
            if (!string.IsNullOrWhiteSpace(model.LawNoteNo))
            {
                sql += " and a.law_note_no like @LawNoteNo";
            }

            sql += " order by a.law_note_no";

            result = DbHelper.Query<LawSearchDetail>(H2ORepository.ConnectionStringName, sql, new
            {
                vmleaderid = orgVm.vmleaderid.Substring(0, 10),
                vmcode = orgVm.vmcode,
                LawDueAgentId = model.LawDueAgentId,
                LawDueName = "%" + model.LawDueName + "%",
                LawNoteNo = "%" + model.LawNoteNo + "%",
                LawDueMoney = model.LawDueMoney,
            }).ToList();
            return result;
        }

        /// <summary>
        /// 已結案件列表
        /// </summary>
        public List<LawSearchDetail> GetLawSearchByClose(LawSearchDetail model, OrgVm orgVm)
        {
            List<LawSearchDetail> result = new List<LawSearchDetail>();
            org check = new org();
            string sql = string.Empty, sql1 = string.Empty;

            if (orgVm.vm_flag == 1)
            {
                if (orgVm.vsm_flag == 0)
                {
                    sql = @"select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where b.vm_code=@vmcode and law_status_type = 2 and (law_content_cancel_type is null or law_content_cancel_type<>1)";
                }
                else
                {
                    sql = @"select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where b.sm_code in (" + orgVm.smstr + ") and law_status_type=2 and (law_content_cancel_type is null or law_content_cancel_type<>1)";
                }
            }
            else
            {
                sql1 = @"select * from org where manager=(select orgid from org where href>65 and login=@login and enabled='Y' and isdelete=0 )and right(orgname,2)='協理'";
                check = DbHelper.Query<org>(H2ORepository.ConnectionStringName, sql1, new
                {
                    login = orgVm.vmleaderid.Substring(0, 10),
                }).FirstOrDefault();

                if (check != null)
                {
                    sql = "select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where left(administrat_id,10)=@vmleaderid and law_status_type=2 and (law_content_cancel_type is null or law_content_cancel_type<>1)";
				}
				else
				{
                    sql = "select *,law_year + '/' + law_month + '-' + convert(nvarchar(255), law_pay_sequence) + '薪' as production_ym,(case when SUBSTRING(center_name,0,CHARINDEX('(',center_name)) = '' then center_name else  SUBSTRING(center_name,0,CHARINDEX('(',center_name)) END) as CenterNameCG,REPLACE(SUBSTRING(wc_center_name,CHARINDEX('(',wc_center_name)+1,len(wc_center_name)),')','') as WcCenterNameCG from law_content a left join law_agent_content b on a.law_note_no=b.law_note_no where left(administrat_id,10)=@vmleaderid and law_status_type=2 and (law_content_cancel_type is null or law_content_cancel_type<>1)";
                }

            }

            if (!string.IsNullOrWhiteSpace(model.LawDueAgentId))
            {
                sql += " and left(law_due_agentid,10)= @LawDueAgentId";
            }
            if (!string.IsNullOrWhiteSpace(model.LawDueName))
            {
                sql += " and law_due_name like @LawDueName";
            }
            if (model.LawDueMoney != 0)
            {
                sql += " and law_due_money= @LawDueMoney";
            }
            if (!string.IsNullOrWhiteSpace(model.LawNoteNo))
            {
                sql += " and a.law_note_no like @LawNoteNo";
            }

            sql += " order by a.law_note_no";

            result = DbHelper.Query<LawSearchDetail>(H2ORepository.ConnectionStringName, sql, new
            {
                vmleaderid = orgVm.vmleaderid.Substring(0, 10),
                vmcode = orgVm.vmcode,
                LawDueAgentId = model.LawDueAgentId,
                LawDueName = "%" + model.LawDueName + "%",
                LawNoteNo = "%" + model.LawNoteNo + "%",
                LawDueMoney = model.LawDueMoney,
            }).ToList();
            return result;
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
