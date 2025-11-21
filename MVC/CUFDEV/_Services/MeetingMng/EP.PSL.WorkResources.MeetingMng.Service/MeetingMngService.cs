using EP.H2OModels;
using EP.Platform.Service;
using Microsoft.CUF;
using Microsoft.CUF.Framework.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EP.PSL.WorkResources.MeetingMng.Service
{

    public class MeetingMngService : IMeetingMngService
    {
        #region 會議CRUD

        /// <summary>
        /// 修改召開會議時間狀態
        /// </summary>
        public void UpdateMeetingTime()
        {
            List<Meeting> m = new List<Meeting>();
            string sql = @"SELECT * FROM H2O.dbo.Meeting Where MTActive = 1";
            m = DbHelper.Query<Meeting>(H2ORepository.ConnectionStringName, sql, new { }).ToList();

            for (int i = 0; i < m.Count; i++)
            {
                if (DateTime.Now >= m[i].MTStartDate)
                {
                    var Upsql = @"Update Meeting Set MTActive = @MTActive FROM Meeting Where MTID = @MTID";
                    DbHelper.Execute(H2ORepository.ConnectionStringName, Upsql, new { MTID = m[i].MTID, MTActive = 2 });
                }
            }
        }

        /// <summary>
        /// 召開會議查詢條件
        /// </summary>
        public List<Meeting> GetMeetingList(QueryMeetingCondition cond)
        {
            List<Meeting> result = new List<Meeting>();
            string sql = @"SELECT MTID,MTName,MTStartDate,MTConvenerName FROM H2O.dbo.Meeting A WHERE ( A.MTID IN (SELECT Distinct(MTID) FROM H2O.dbo.MeetingMemberRelation WHERE imember = @imember";
            sql += " OR A.MTConvener = @imember OR A.MTChairman = @imember OR A.MTRecorder = @imember))";
            //會議名稱
            if (!string.IsNullOrWhiteSpace(cond.MTName))
            {
                sql += " AND MTName like @MTName ";
            }

            //會議說明
            if (!string.IsNullOrWhiteSpace(cond.MTDesc))
            {
                sql += " AND MTDesc like @MTDesc ";
            }
            //會議狀態
            if (!string.IsNullOrWhiteSpace(cond.MeetingReadType))
            {
                switch (cond.MeetingReadType)
                {
                    case "1":///未召開會議
                        sql += " AND MTActive = 1";
                        break;

                    case "2":///已召開會議
                        sql += " AND MTActive = 2";
                        break;

                    case "3":///我舉辦的會議
                        sql += " AND MTConvener = @imember";
                        break;

                    case "4":///歷史資料
                        sql += " AND MTActive = 4";
                        break;
                }

            }

            result = DbHelper.Query<Meeting>(H2ORepository.ConnectionStringName, sql,
                new
                {
                    imember = cond.imember,
                    MTName = "%" + cond.MTName + "%",
                    MTDesc = "%" + cond.MTDesc + "%",
                }).ToList();
            return result;
        }

        /// <summary>
        /// 取得會議的類型名稱列表
        /// </summary>
        /// <returns></returns>
        public List<ValueText> GetMeetingKindList()
        {
            List<ValueText> result = new List<ValueText>();

            string sql = @"SELECT TermDescCode as Value,TermDesc as Text FROM TermVal where DatabaseName=@DatabaseName  and TermId=@TermId  order by TermSN";
            result = DbHelper.Query<ValueText>(CUFRepository.ConnectionStringName, sql, new
            {
                DatabaseName = "H2O",
                TermId = "MeetingMngType"
            }).ToList();

            return result;
        }

        /// <summary>
        /// 取得會議列表
        /// </summary>
        public Meeting GetMeeting(int id)
        {
            Meeting result = new Meeting();
            string sql = @"select * from Meeting where MTID = @MTID";
            result = DbHelper.Query<Meeting>(H2ORepository.ConnectionStringName, sql, new { MTID = id }).FirstOrDefault();
            return result;
        }
        /// <summary>
        /// 取得會議地點列表
        /// </summary>
        public List<Meeting> GetMeetingPlace()
        {
            List<Meeting> result = new List<Meeting>();
            string sql = @"select distinct MTPlace from Meeting where MTPlace is not null group by MTPlace";
            result = DbHelper.Query<Meeting>(H2ORepository.ConnectionStringName, sql, new { }).ToList();
            return result;
        }

        /// <summary>
        /// 取得場地和設備列表
        /// </summary>
        public List<MeetingDevice> GetMeetingDevice()
        {
            List<MeetingDevice> result = new List<MeetingDevice>();
            string sql = @"select MDName,MDID from MeetingDevice where MDIsDelete = 0";
            result = DbHelper.Query<MeetingDevice>(H2ORepository.ConnectionStringName, sql, new { }).ToList();
            return result;
        }


        /// <summary>
        /// 用流水號取得編輯頁面主席部門和名字
        /// </summary>
        /// <param name="id">會議序號</param>
        public List<MeetingDetail> GetEditMeetingChairmanById(int id)
        {
            List<MeetingDetail> result = new List<MeetingDetail>();
            string sql = @" select * from h2o.dbo.Meeting A join cuf.dbo.sc_member B on A.MTChairman = B.imember join cuf.dbo.sc_unit  C on B.uunit = C.uunit join CUF.dbo.sc_account D on A.MTChairman = D.imember";
            sql += " where A.MTID = @MTID";
            result = DbHelper.Query<MeetingDetail>(H2ORepository.ConnectionStringName, sql,
                           new
                           {
                               MTID = id,
                           }).ToList();
            return result;
        }

        /// <summary>
        /// 用流水號取得編輯頁面紀錄者部門和名字
        /// </summary>
        /// <param name="id">會議序號</param>
        public List<MeetingDetail> GetEditMeetingRecorderById(int id)
        {
            List<MeetingDetail> result = new List<MeetingDetail>();
            string sql = @"select * from h2o.dbo.Meeting A join cuf.dbo.sc_member B on A.MTRecorder = B.imember join cuf.dbo.sc_unit  C on B.uunit = C.uunit join CUF.dbo.sc_account D on A.MTRecorder = D.imember";
            sql += " where A.MTID = @MTID";
            result = DbHelper.Query<MeetingDetail>(H2ORepository.ConnectionStringName, sql,
                           new
                           {
                               MTID = id,
                           }).ToList();
            return result;
        }

        /// <summary>
        /// 用流水號取得編輯頁面與會人員部門和名字
        /// </summary>
        /// <param name="id">會議序號</param>
        public List<MeetingDetail> GetEditMeetingParticipantsById(int id)
        {
            List<MeetingDetail> result = new List<MeetingDetail>();
            string sql = @"SELECT * FROM H2O.dbo.MeetingMemberRelation A join cuf.dbo.sc_member B on A.imember = B.imember join cuf.dbo.sc_unit  C on B.uunit = C.uunit join CUF.dbo.sc_account D on A.imember = D.imember";
            sql += " where A.MTID = @MTID";
            result = DbHelper.Query<MeetingDetail>(H2ORepository.ConnectionStringName, sql,
                           new
                           {
                               MTID = id,
                           }).ToList();
            return result;
        }

        /// <summary>
        /// 取得場地設備名稱
        /// </summary>
        /// <param name="id">會議序號</param>
        public string GetMDIDToShow(int MDID)
        {
            string sql = @" SELECT *FROM H2O.dbo.MeetingDevice where MDID = @MDID";
            MeetingDevice result = new MeetingDevice();

            result = DbHelper.Query<MeetingDevice>(H2ORepository.ConnectionStringName, sql,
                 new
                 {

                     MDID = MDID,

                 }).FirstOrDefault();

            return result.MDName;
        }

        /// <summary>
        /// 檢查會議時間有無重複
        /// </summary>
        public MeetingDetail ChkMeetingTime(int MDID, DateTime SDate, DateTime EDate)
        {
            MeetingDetail result = new MeetingDetail();

            string sql = "select * from MeetingOrder t1,MeetingDevice D ";
            sql += "  where  t1.MDID = D.MDID and ((t1.MOStartDate <= @MOStartDate AND t1.MOEndDate >=  @MOStartDate) or (t1.MOStartDate <=  @MOEndDate AND t1.MOEndDate >= @MOEndDate)";
            sql += " or (t1.MOStartDate >= @MOStartDate AND t1.MOEndDate <= @MOEndDate))";
            //for (int i = 0; i < MDIDList.Count; i++)
            //{
            //    MDID = MDIDList[i];
            //    sql += " and t1.MDID = @MDID";
            //    result = DbHelper.Query<MeetingDetail>(H2ORepository.ConnectionStringName, sql,
            //         new
            //         {
            //             MDID = MDID,
            //             MOStartDate = SDate,
            //             MOEndDate = EDate
            //         }).FirstOrDefault();
            //    if (result != null)
            //    {
            //        return result;
            //    }
            //}           
            sql += " and t1.MDID = @MDID";
            result = DbHelper.Query<MeetingDetail>(OldH2ORepository.ConnectionStringName, sql,
                 new
                 {
                     MDID = MDID,
                     MOStartDate = SDate.ToString("yyyy/MM/dd HH:mm"),
                     MOEndDate = EDate.ToString("yyyy/MM/dd HH:mm")
                 }).FirstOrDefault();
            if (result != null)
            {
                return result;
            }
            return result;
        }

        /// <summary>
        /// 檢查修改會議時間有無重複
        /// </summary>
        public MeetingDetail ChkEditMeetingTime(int MDID, DateTime SDate, DateTime EDate, int MTID)
        {
            MeetingDetail result = new MeetingDetail();
            string sql = @"with t1 as(select * from H2O.dbo.MeetingOrder)";
            sql += " select * from t1,MeetingDevice D,Meeting A";
            sql += " WHERE t1.MDID = D.MDID and t1.MTID = A.MTID and ((t1.MOStartDate <= @MOStartDate AND t1.MOEndDate >=  @MOStartDate) or (t1.MOStartDate <=  @MOEndDate AND t1.MOEndDate >= @MOEndDate)";
            sql += " or (t1.MOStartDate >= @MOStartDate AND t1.MOEndDate <= @MOEndDate))";


            sql += " and t1.MDID = @MDID and t1.MTID <> @MTID";
            result = DbHelper.Query<MeetingDetail>(H2ORepository.ConnectionStringName, sql,
                 new
                 {
                     MTID = MTID,
                     MDID = MDID,
                     MOStartDate = SDate,
                     MOEndDate = EDate
                 }).FirstOrDefault();
            if (result != null)
            {
                return result;
            }

            return result;
        }

        /// <summary>
        /// 檢查設備與場地有無
        /// </summary>
        public bool ChkMeetingDevice(int MTID)
        {
            var result = false;
            MeetingDeviceRelation m = new MeetingDeviceRelation();
            string sql = @"select * from H2O.dbo.MeetingDeviceRelation";
            sql += " where MTID = @MTID";
            m = DbHelper.Query<MeetingDeviceRelation>(H2ORepository.ConnectionStringName, sql, new { MTID = MTID }).FirstOrDefault();
            if (m == null)
            {
                result = true;
            }
            return result;
        }

        /// <summary>
        /// 新增設備與場地關聯
        /// </summary>
        public void CreateMeetingDeviceRelation(MeetingDeviceRelation model)
        {
            model.Insert(H2ORepository.ConnectionStringName);
        }

        /// <summary>
        /// 新增會議
        /// </summary>
        public int CreateMeeting(Meeting model)
        {
            model.Insert(H2ORepository.ConnectionStringName);
            return model.MTID;
        }

        /// <summary>
        /// 新增會議管理&設備中心預定的會議室和設備
        /// </summary>
        public void CreateMeetingOrder(List<MeetingOrder> modeltoList)
        {
            for (int i = 0; i < modeltoList.Count; i++)
            {
                modeltoList[i].Insert(new[] { H2ORepository.ConnectionStringName });
            }
        }

        /// <summary>
        /// 新增會議檔案&相關檔案
        /// </summary>
        /// <param name="modelfileList">檔案附件物件</param>
        /// <param name="tabUniqueId">前端傳來唯一碼</param>
        /// <param name="RetMsg">回傳訊息</param>
        public void CreateMeetingFile(Meeting model, List<MeetingFile> modelfileList, string tabUniqueId, out string RetMsg)
        {
            RetMsg = "";

            try
            {
                IDataSettingService dsService = new DataSettingService();
                var tempDirRoot = dsService.GetConfigValueByName("TempDir");
                string currRootDir = Path.Combine(tempDirRoot, model.MTConvener + @"\" + tabUniqueId);
                string newRootDir = dsService.GetConfigValueByName("MeetingMngDir");

                #region 執行資料
                string currefile = string.Empty;
                string extname = string.Empty;
                string newfilename = string.Empty;

                using (TransactionScope ts = new TransactionScope())
                {
                    for (int i = 0; i < modelfileList.Count; i++)
                    {
                        currefile = Path.Combine(currRootDir, modelfileList[i].MFMd5Name);
                        modelfileList[i].MFMd5Name = DateTime.Now.ToString("yyyyMM") + @"\" + modelfileList[i].MFMd5Name;
                        System.IO.FileInfo fi = new System.IO.FileInfo(currefile);
                        newfilename = Path.Combine(newRootDir, modelfileList[i].MFMd5Name);
                        //string uploaddir = Path.GetDirectoryName(newfilename);
                        //if (!Directory.Exists(uploaddir))
                        //    Directory.CreateDirectory(uploaddir);

                        modelfileList[i].Insert(new[] { H2ORepository.ConnectionStringName });
                        fi.CopyTo(newfilename);
                    }
                    ts.Complete();
                }
                #endregion
            }
            catch (IOException ioe)
            {
                Throw.SQL("檔案儲存錯誤:" + ioe.Message);
                RetMsg = ioe.Message;
            }
            catch (Exception e)
            {
                Throw.SQL(e.Message);
                RetMsg = e.Message;
            }

        }

        /// <summary>
        /// 新增會議資訊&人員資訊關聯&回覆確認
        /// </summary>
        public void CreateMeetingMemberRelation(List<MeetingMemberRelation> modeltoList)
        {
            for (int i = 0; i < modeltoList.Count; i++)
            {
                modeltoList[i].Insert(new[] { H2ORepository.ConnectionStringName });
            }
        }

        /// <summary>
        /// 用流水號取得編輯頁面場地與設備值
        /// </summary>
        /// <param name="id">會議序號</param>
        public MeetingDevice GetEditMeetingDevice(int id)
        {
            MeetingDevice result = new MeetingDevice();
            string sql = @" select MDName,A.MDID from h2o.dbo.MeetingDevice A join h2o.dbo.MeetingDeviceRelation B on A.MDID = B.MDID";
            sql += " where B.MTID = @MTID";
            result = DbHelper.Query<MeetingDevice>(H2ORepository.ConnectionStringName, sql,
                           new
                           {
                               MTID = id,
                           }).FirstOrDefault();
            return result;

        }

        /// <summary>
        /// 用流水號取得編輯頁面場地與設備關聯表值
        /// </summary>
        /// <param name="id">會議序號</param>
        public List<MeetingDeviceRelation> GetEditMeetingDeviceRelation(int id)
        {
            List<MeetingDeviceRelation> result = new List<MeetingDeviceRelation>();
            string sql = @" select * from h2o.dbo.MeetingDevice A join h2o.dbo.MeetingDeviceRelation B on A.MDID = B.MDID";
            sql += " where MDIsDelete = 0 and B.MTID = @MTID";
            result = DbHelper.Query<MeetingDeviceRelation>(H2ORepository.ConnectionStringName, sql,
                           new
                           {
                               MTID = id,
                           }).ToList();
            return result;

        }

        /// <summary>
        /// 修改設備與場地關聯
        /// </summary>
        public void EditMeetingDeviceRelation(List<MeetingDeviceRelation> modeltoList)
        {
            H2ORepository.Delete<MeetingDeviceRelation>(new { MTID = modeltoList[0].MTID });
            for (int i = 0; i < modeltoList.Count; i++)
            {
                modeltoList[i].Insert(new[] { H2ORepository.ConnectionStringName });
            }
        }

        /// <summary>
        /// 修改會議
        /// </summary>
        public void EditMeeting(Meeting model)
        {
            model.Update(new { MTID = model.MTID }, H2ORepository.ConnectionStringName);
        }

        /// <summary>
        /// 修改會議管理&設備中心預定的會議室和設備
        /// </summary>
        public void EditMeetingOrder(List<MeetingOrder> modeltoList)
        {
            H2ORepository.Delete<MeetingOrder>(new { MTID = modeltoList[0].MTID });
            for (int i = 0; i < modeltoList.Count; i++)
            {
                modeltoList[i].Insert(new[] { H2ORepository.ConnectionStringName });
            }
        }

        /// <summary>
        /// 修改會議檔案
        /// </summary>
        /// <param name="modelfileList">檔案附件物件</param>
        /// <param name="tabUniqueId">前端傳來唯一碼</param>
        /// <param name="RetMsg">回傳訊息</param>
        public void EditMeetingFile(Meeting model, List<MeetingFile> modelfileList, string tabUniqueId, out string RetMsg)
        {
            RetMsg = "";
            string sql = @"SELECT * FROM H2O.dbo.MeetingFile WHERE MFType = 1 and MTID = @MTID";
            List<MeetingFile> result = DbHelper.Query<MeetingFile>(H2ORepository.ConnectionStringName, sql, new
            {
                MTID = modelfileList[0].MTID
            }).ToList();

            if (result.Count != 0)
            {
                for (int j = 0; j < modelfileList.Count; j++)
                {
                    string s = result[j].MFMd5Name;
                    s = s.Substring(7);
                    //判斷有沒有更動舊檔案
                    if (modelfileList[j].MFMd5Name != s)
                    {
                        //H2ORepository.Delete<MeetingFile>(new { MFID = result[j].MFID , MTID = modelfileList[0].MTID, MFType = 1 });
                        try
                        {
                            IDataSettingService dsService = new DataSettingService();
                            var tempDirRoot = dsService.GetConfigValueByName("TempDir");
                            string currRootDir = Path.Combine(tempDirRoot, model.MTConvener + @"\" + tabUniqueId);
                            string newRootDir = dsService.GetConfigValueByName("MeetingMngDir");

                            #region 執行資料
                            string currefile = string.Empty;
                            string extname = string.Empty;
                            string newfilename = string.Empty;

                            using (TransactionScope ts = new TransactionScope())
                            {
                                for (int i = 0; i < modelfileList.Count; i++)
                                {
                                    currefile = Path.Combine(currRootDir, modelfileList[i].MFMd5Name);
                                    modelfileList[i].MFMd5Name = DateTime.Now.ToString("yyyyMM") + @"\" + modelfileList[i].MFMd5Name;
                                    System.IO.FileInfo fi = new System.IO.FileInfo(currefile);
                                    newfilename = Path.Combine(newRootDir, modelfileList[i].MFMd5Name);
                                    //string uploaddir = Path.GetDirectoryName(newfilename);
                                    //if (!Directory.Exists(uploaddir))
                                    //    Directory.CreateDirectory(uploaddir);

                                    modelfileList[i].Insert(new[] { H2ORepository.ConnectionStringName });
                                    //用copy的 因為message用move
                                    fi.CopyTo(newfilename);
                                }
                                ts.Complete();
                            }
                            #endregion
                        }
                        catch (IOException ioe)
                        {
                            Throw.SQL("檔案儲存錯誤:" + ioe.Message);
                            RetMsg = ioe.Message;
                        }
                        catch (Exception e)
                        {
                            Throw.SQL(e.Message);
                            RetMsg = e.Message;
                        }
                    }
                }
            }
            else
            {
                try
                {
                    IDataSettingService dsService = new DataSettingService();
                    var tempDirRoot = dsService.GetConfigValueByName("TempDir");
                    string currRootDir = Path.Combine(tempDirRoot, model.MTConvener + @"\" + tabUniqueId);
                    string newRootDir = dsService.GetConfigValueByName("MeetingMngDir");

                    #region 執行資料
                    string currefile = string.Empty;
                    string extname = string.Empty;
                    string newfilename = string.Empty;

                    using (TransactionScope ts = new TransactionScope())
                    {
                        for (int i = 0; i < modelfileList.Count; i++)
                        {
                            currefile = Path.Combine(currRootDir, modelfileList[i].MFMd5Name);
                            modelfileList[i].MFMd5Name = DateTime.Now.ToString("yyyyMM") + @"\" + modelfileList[i].MFMd5Name;
                            System.IO.FileInfo fi = new System.IO.FileInfo(currefile);
                            newfilename = Path.Combine(newRootDir, modelfileList[i].MFMd5Name);
                            //string uploaddir = Path.GetDirectoryName(newfilename);
                            //if (!Directory.Exists(uploaddir))
                            //    Directory.CreateDirectory(uploaddir);

                            modelfileList[i].Insert(new[] { H2ORepository.ConnectionStringName });
                            //用copy的 因為message用move
                            fi.CopyTo(newfilename);
                        }
                        ts.Complete();
                    }
                    #endregion
                }
                catch (IOException ioe)
                {
                    Throw.SQL("檔案儲存錯誤:" + ioe.Message);
                    RetMsg = ioe.Message;
                }
                catch (Exception e)
                {
                    Throw.SQL(e.Message);
                    RetMsg = e.Message;
                }
            }

        }

        /// <summary>
        /// 修改會議資訊&人員資訊關聯&回覆確認
        /// </summary>
        public void EditMeetingMemberRelation(List<MeetingMemberRelation> modeltoList)
        {
            H2ORepository.Delete<MeetingMemberRelation>(new { MTID = modeltoList[0].MTID });
            for (int i = 0; i < modeltoList.Count; i++)
            {
                modeltoList[i].Insert(new[] { H2ORepository.ConnectionStringName });
            }
        }

        /// <summary>
        /// 會議刪除功能
        /// </summary>
        /// <param name="MTID"></param>
        public bool DeleteMeeting(int MTID)
        {
            var result = false;
            using (TransactionScope ts = new TransactionScope())
            {
                //刪除會議主檔
                H2ORepository.Delete<Meeting>(new { MTID = MTID });
                //刪除會議檔案
                H2ORepository.Delete<MeetingFile>(new { MTID = MTID });
                //刪除會議管理&設備中心預定的會議室和設備
                H2ORepository.Delete<MeetingOrder>(new { MTID = MTID });
                //刪除會議成員表
                H2ORepository.Delete<MeetingMemberRelation>(new { MTID = MTID });

                //刪除決議事項

                string sql = @"SELECT * FROM H2O.dbo.MeetingJobRelation where MTID = @MTID";
                MeetingJobRelation MJR = DbHelper.Query<MeetingJobRelation>(H2ORepository.ConnectionStringName, sql, new { MTID = MTID }).FirstOrDefault();
                if (MJR != null)
                {
                    H2ORepository.Delete<Job>(new { JBID = MJR.JBID });
                    H2ORepository.Delete<JobPercentage>(new { JBID = MJR.JBID });
                }
                H2ORepository.Delete<MeetingJobRelation>(new { MTID = MTID });
                ts.Complete();
                result = true;
            }
            return result;
        }
        /// <summary>
        /// 會議刪除設備中心連動
        /// </summary>
        /// <param name="MTID"></param>
        public bool DeleteMeetingOrder(int MTID)
        {
            var result = false;
            using (TransactionScope ts = new TransactionScope())
            {
                //刪除會議管理 & 設備中心預定連動
                H2ORepository.Delete<MeetingDeviceRelation>(new { MTID = MTID });
                //刪除會議管理&設備中心預定的會議室和設備
                H2ORepository.Delete<MeetingOrder>(new { MTID = MTID });

                ts.Complete();
                result = true;
            }
            return result;
        }

        /// <summary>
        /// 編輯會議時會議檔案刪除功能
        /// </summary>
        /// <param name="mf">會議檔案</param>
        public void DeleteEditMeetingFile(MeetingFile mf)
        {
            MeetingFile result = new MeetingFile();
            string sql = "SELECT * FROM H2O.dbo.MeetingFile WHERE MFMd5Name like @MFMd5Name";
            result = DbHelper.Query<MeetingFile>(H2ORepository.ConnectionStringName, sql,
               new
               {
                   MFMd5Name = "%" + mf.MFMd5Name + "%"
               }).FirstOrDefault();
            if (result != null)
            {
                H2ORepository.Delete<MeetingFile>(new { MFID = result.MFID });
            }

        }
        #endregion

        #region 會議詳細頁面相關
        /// <summary>
        /// 詳細頁面顯示
        /// </summary>
        /// <param name="id">會議序號</param>
        /// <param name="imember">員編</param>
        public Meeting GetMeetingDetailById(int id, string imember)
        {
            string sql = "";
            sql = @"SELECT * FROM H2O.dbo.Meeting A WHERE";
            sql += " ( A.MTID IN (SELECT Distinct(MTID) FROM H2O.dbo.MeetingMemberRelation WHERE imember = @imember";
            sql += " OR A.MTConvener = @imember OR A.MTChairman = @imember OR A.MTRecorder = @imember)) and MTID = @MTID";

            Meeting result = DbHelper.Query<Meeting>(H2ORepository.ConnectionStringName, sql, new { MTID = id, imember = imember }).FirstOrDefault();
            result = DbHelper.Query<Meeting>(H2ORepository.ConnectionStringName, sql, new { MTID = id, imember = imember }).FirstOrDefault();

            return result;
        }

        /// <summary>
        /// 詳細頁面檔案顯示
        /// </summary>
        /// <param name="id">會議序號</param>
        public List<MeetingFile> GetMeetingDetailFilesById(int id)
        {
            List<MeetingFile> result = new List<MeetingFile>();
            string sql = @"SELECT * FROM H2O.dbo.MeetingFile where MTID = @MTID";
            result = DbHelper.Query<MeetingFile>(H2ORepository.ConnectionStringName, sql,
                           new
                           {
                               MTID = id,
                           }).ToList();
            return result;
        }

        /// <summary>
        /// 詳細頁面與會人員顯示
        /// </summary>
        /// <param name="id">會議序號</param>
        public List<MeetingMemberRelation> GetMeetingDetailParticipantsById(int id)
        {
            List<MeetingMemberRelation> result = new List<MeetingMemberRelation>();
            string sql = @"SELECT * FROM H2O.dbo.MeetingMemberRelation";
            sql += " where MTID = @MTID";
            result = DbHelper.Query<MeetingMemberRelation>(H2ORepository.ConnectionStringName, sql,
                           new
                           {
                               MTID = id,
                           }).ToList();
            return result;
        }

        /// <summary>
        /// 詳細頁面參加狀態顯示
        /// </summary>
        /// <param name="id">會議序號</param>
        public MeetingMemberRelation GetMeetingDetailMTReplyById(int id, string imember)
        {
            MeetingMemberRelation result = new MeetingMemberRelation();
            string sql = @"SELECT MTReply FROM H2O.dbo.MeetingMemberRelation";
            sql += " where MTID = @MTID and imember = @imember";
            result = DbHelper.Query<MeetingMemberRelation>(H2ORepository.ConnectionStringName, sql,
                           new
                           {
                               MTID = id,
                               imember = imember
                           }).FirstOrDefault();
            return result;
        }

        /// <summary>
        /// 詳細頁面已回覆參加人員顯示
        /// </summary>
        /// <param name="id">會議序號</param>
        //public List<MeetingMemberRelation> GetMeetingDetailParticipateById(int id)
        //{
        //    List<MeetingMemberRelation> result = new List<MeetingMemberRelation>();
        //    string sql = @"SELECT * FROM H2O.dbo.MeetingMemberRelation";
        //    sql += "  where MTReply = 1 and MTID = @MTID";
        //    result = DbHelper.Query<MeetingMemberRelation>(H2ORepository.ConnectionStringName, sql,
        //                   new
        //                   {
        //                       MTID = id,
        //                   }).ToList();
        //    return result;
        //}

        /// <summary>
        /// 詳細頁面已回覆不參加人員顯示
        /// </summary>
        /// <param name="id">會議序號</param>
        //public List<MeetingMemberRelation> GetMeetingDetailNoParticipateById(int id)
        //{
        //    List<MeetingMemberRelation> result = new List<MeetingMemberRelation>();
        //    string sql = @"SELECT * FROM H2O.dbo.MeetingMemberRelation";
        //    sql += "  where MTReply = 2 and MTID = @MTID";
        //    result = DbHelper.Query<MeetingMemberRelation>(H2ORepository.ConnectionStringName, sql,
        //                   new
        //                   {
        //                       MTID = id,
        //                   }).ToList();
        //    return result;
        //}

        /// <summary>
        /// 詳細頁面未回覆人員顯示
        /// </summary>
        /// <param name="id">會議序號</param>
        //public List<MeetingMemberRelation> GetMeetingDetailNoReplyById(int id)
        //{
        //    List<MeetingMemberRelation> result = new List<MeetingMemberRelation>();
        //    string sql = @"SELECT * FROM H2O.dbo.MeetingMemberRelation";
        //    sql += "  where MTReply = 0 and MTID = @MTID";
        //    result = DbHelper.Query<MeetingMemberRelation>(H2ORepository.ConnectionStringName, sql,
        //                   new
        //                   {
        //                       MTID = id,
        //                   }).ToList();
        //    return result;
        //}

        /// <summary>
        /// 詳細頁面人員參加狀態修改
        /// </summary>
        /// <param name="id">會議序號</param>
        /// <param name="mtreply">會議人員參加狀態</param>
        /// <param name="imember">成員</param>
        public void UpdateMeetingReplyById(int mtid, int mtreply, string imember)
        {
            var sql = @"Update MeetingMemberRelation Set MTReply = @MTReply FROM Meeting A JOIN MeetingMemberRelation B ON A.MTID=B.MTID Where A.MTID = @MTID and imember = @imember";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new { MTID = mtid, MTReply = mtreply, imember = imember });
        }

        /// <summary>
        /// 詳細頁面會議狀態修改加入至歷史資料
        /// </summary>
        /// <param name="id">會議序號</param>
        /// <param name="mtactive">會議狀態</param>
        public void UpdateMeetingActiveById(int mtid, int mtactive)
        {
            var sql = @"Update Meeting Set MTActive = @MTActive FROM Meeting Where MTID = @MTID";
            DbHelper.Execute(H2ORepository.ConnectionStringName, sql, new { MTID = mtid, MTActive = mtactive });
        }
        /// <summary>
        /// 用會議檔案流水號取得單一檔案並下載
        /// </summary>
        /// <param name="fid">會議檔案流水號</param>
        public MeetingFile GetMeetingFileByMfid(int fid)
        {
            string sql = @"SELECT * FROM H2O.dbo.MeetingFile WHERE MFID = @MFID";
            MeetingFile result = DbHelper.Query<MeetingFile>(H2ORepository.ConnectionStringName, sql, new
            {
                MFID = fid
            }).SingleOrDefault();

            return result;
        }

        /// <summary>
        /// 用會議流水號取得會議附件檔
        /// </summary>
        /// <param name="id">會議流水號</param>
        public List<MeetingFile> GetMeetingFileByMTId(int id)
        {
            string sql = @"SELECT * FROM H2O.dbo.MeetingFile WHERE MFType = 1 and MTID = @MTID";
            List<MeetingFile> result = DbHelper.Query<MeetingFile>(H2ORepository.ConnectionStringName, sql, new
            {
                MTID = id
            }).ToList();
            return result;
        }

        #endregion

        #region 相關檔案
        /// <summary>
        /// 相關檔案資訊顯示
        /// </summary>
        /// <param name="cond">會議相關檔案查詢條件</param>
        public List<MeetingFileTo> GetMeetingRelatedFilesById(QueryMeetingFilesCondition cond)
        {
            List<MeetingFileTo> result = new List<MeetingFileTo>();
            string sql = @"select * from h2o.dbo.MeetingFile A";
            sql += " join cuf.dbo.sc_member B on A.MFCreater = B.imember join cuf.dbo.sc_unit C on B.uunit = C.uunit";
            sql += " where MFType in (2,3) and MTID = @MTID";
            //檔案名稱
            if (!string.IsNullOrWhiteSpace(cond.MFFileName))
            {
                sql += " AND MFFileName like @MFFileName ";
            }
            //檔案說明
            if (!string.IsNullOrWhiteSpace(cond.MFDesc))
            {
                sql += " AND MFDesc like @MFDesc ";
            }

            result = DbHelper.Query<MeetingFileTo>(H2ORepository.ConnectionStringName, sql,
                new
                {
                    MTID = cond.MTID,
                    MFFileName = "%" + cond.MFFileName + "%",
                    MFDesc = "%" + cond.MFDesc + "%",
                }).ToList();
            return result;
        }

        /// <summary>
        /// 相關檔案刪除功能
        /// </summary>
        /// <param name="MFID">會議檔案流水號</param>
        public bool DeleteMeetingFile(int MFID)
        {
            var result = false;
            using (TransactionScope ts = new TransactionScope())
            {
                var count = 0;

                //刪除主檔
                count = H2ORepository.Delete<MeetingFile>(new { MFID = MFID });
                ts.Complete();
                result = true;
            }
            return result;
        }
        #endregion

        #region 決議事項
        /// <summary>
        /// 決議事項資訊顯示
        /// </summary>
        /// <param name="cond">決議事項查詢條件</param>
        public List<JobDetail> GetMeetingJobById(QueryMeetingJobCondition cond)
        {
            List<JobDetail> result = new List<JobDetail>();
            string sql = @"SELECT A.JBID,A.JBDesc,A.JBSubject,A.JBCreater,A.JBStartdate,A.JBEnddate,B.*,C.MTID";
            sql += " FROM h2o.dbo.Job A,h2o.dbo.JobPercentage B, h2o.dbo.MeetingJobRelation C";
            sql += " WHERE A.JBID = B.JBID AND A.JBID = C.JBID AND C.MTID = @MTID";


            //說明
            if (!string.IsNullOrWhiteSpace(cond.JBDesc))
            {
                sql += " AND A.JBDesc like @JBDesc ";
            }
            //標題
            if (!string.IsNullOrWhiteSpace(cond.JBSubject))
            {
                sql += " AND A.JBSubject like @JBSubject ";
            }

            result = DbHelper.Query<JobDetail>(H2ORepository.ConnectionStringName, sql,
                new
                {
                    MTID = cond.MTID,
                    JBDesc = "%" + cond.JBDesc + "%",
                    JBSubject = "%" + cond.JBSubject + "%",
                }).ToList();
            return result;
        }

        /// <summary>
        /// 新增決議事項
        /// </summary>
        public int CreateJob(Job model)
        {
            model.Insert(H2ORepository.ConnectionStringName);
            return model.JBID;
        }

        /// <summary>
        /// 新增會議資訊&決議事項關聯表
        /// </summary>
        public void CreateMeetingJobRelation(MeetingJobRelation model)
        {
            model.Insert(H2ORepository.ConnectionStringName);
        }

        /// <summary>
        /// 新增決議事項進度表
        /// </summary>
        public void CreateJobPercentage(List<JobPercentage> modeltoList)
        {
            for (int i = 0; i < modeltoList.Count; i++)
            {
                modeltoList[i].Insert(new[] { H2ORepository.ConnectionStringName });
            }
        }

        /// <summary>
        /// 決議事項詳細頁面
        /// </summary>
        public JobDetail GetJobDetailById(int JPID)
        {
            //string sql = @"SELECT A.JBID,A.JBDesc,A.JBSubject,A.JBStartdate,A.JBEnddate,A.JBCreateIP,A.JBCreateDate,A.JBCreater,B.*,C.MTID,D.nmember,E.nunit,F.nmember as Creater,G.nunit as CreaterNunit";
            //sql += " FROM h2o.dbo.Job A,h2o.dbo.JobPercentage B, h2o.dbo.MeetingJobRelation C,CUF.dbo.sc_member D,CUF.dbo.sc_unit E,CUF.dbo.sc_member F,CUF.dbo.sc_unit G";
            //sql += " WHERE A.JBID = B.JBID AND A.JBID = C.JBID AND B.imember = D.imember AND D.uunit = E.uunit AND A.JBCreater = F.imember AND F.uunit = G.uunit";
            //sql += " AND C.MTID = @MTID AND A.JBID = @JBID AND B.imember = @imember";
            string sql = @"SELECT *,F.nmember as Creater,G.nunit as CreaterNunit FROM H2O.dbo.JobPercentage A join H2O.dbo.Job B on A.JBID = B.JBID join  H2O.dbo.MeetingJobRelation C on A.JBID = C.JBID join H2O.dbo.Meeting D on C.MTID = D.MTID 
            join CUF.dbo.sc_member F on B.JBCreater = F.imember join CUF.dbo.sc_unit G on F.uunit = G.uunit";
            sql += " Where JPID = @JPID";


            JobDetail result = DbHelper.Query<JobDetail>(H2ORepository.ConnectionStringName, sql, new { JPID = JPID }).FirstOrDefault();

            return result;
        }

        /// <summary>
        /// 會議追蹤事項查詢條件
        /// </summary>
        public List<JobDetail> GetMeetingJobList(QueryMeetingJobCondition cond)
        {
            List<JobDetail> result = new List<JobDetail>();
            string sql = @"SELECT distinct JBSubject, MTName, JBStartDate,JBEndDate ,B.MTID FROM H2O.dbo.Job A join H2O.dbo.MeetingJobRelation B on A.JBID = B.JBID ";
            sql += " join H2O.dbo.Meeting C on B.MTID = C.MTID join H2O.dbo.JobPercentage D on B.JBID = D.JBID WHERE ( C.MTID IN (SELECT Distinct(MTID) FROM H2O.dbo.MeetingMemberRelation";
            sql += " WHERE imember = @imember OR C.MTConvener = @imember OR C.MTChairman = @imember OR C.MTRecorder = @imember))";
            //追蹤事項
            if (!string.IsNullOrWhiteSpace(cond.JBSubject))
            {
                sql += " AND JBSubject like @JBSubject ";
            }
            //追蹤事項狀態
            if (!string.IsNullOrWhiteSpace(cond.MeetingJobReadType))
            {
                switch (cond.MeetingJobReadType)
                {
                    case "1":///未完成追蹤事項
                        sql += " AND JPPercentage = 0";
                        break;

                    case "2":///已完成追蹤事項
                        sql += " AND JPPercentage = 100";
                        break;

                    case "3":///被拒絕追蹤事項
                        sql += " AND JPPercentage = -1";
                        break;
                }
            }
            result = DbHelper.Query<JobDetail>(H2ORepository.ConnectionStringName, sql,
                new
                {
                    imember = cond.imember,
                    JBSubject = "%" + cond.JBSubject + "%",
                }).ToList();
            return result;
        }

        /// <summary>
        /// 決議事項刪除功能
        /// </summary>
        /// <param name="JBID">決議追蹤事項流水號</param>
        /// <param name="JPID">決議追蹤事項進度表流水號</param>
        public bool DeleteJob(int JBID, int JPID)
        {
            var result = false;
            using (TransactionScope ts = new TransactionScope())
            {
                var count = 0;
                var count1 = 0;
                var count2 = 0;
                //先刪除JobPercentage，因為是一個JPID對應一個成員
                count = H2ORepository.Delete<JobPercentage>(new { JPID = JPID });
                //在搜尋還有沒有這筆(Job)資料
                string sql = @"select * from H2O.dbo.JobPercentage where JBID = @JBID";
                JobPercentage JP = DbHelper.Query<JobPercentage>(H2ORepository.ConnectionStringName, sql, new { JBID = JBID }).FirstOrDefault();
                //有就不刪關聯表
                if (JP == null)
                {
                    count1 = H2ORepository.Delete<Job>(new { JBID = JBID });
                    count2 = H2ORepository.Delete<MeetingJobRelation>(new { JBID = JBID });
                }

                ts.Complete();
                result = true;
            }
            return result;
        }
        #endregion
    }
}
