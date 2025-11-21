using Dapper;
using Newtonsoft.Json;
using SACTAPI.Models;
using SACTAPI.Utilities;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Http;

namespace SACTAPI.Controllers
{
    [RoutePrefix("EPWebAPI")]
    public class SACTAPIQU002Controller : ApiController
    {
        private readonly string _connectionString;
        private readonly string _notNextSign;
        private readonly string _isTest;

        public SACTAPIQU002Controller()
        {
            _connectionString = ConfigurationManager.ConnectionStrings["SACT"].ConnectionString;
            _notNextSign = "下一位簽核者不存在。";
            _isTest = ConfigurationManager.AppSettings["IsTest"];
        }

        private IDbConnection Connection
        {
            get { return new SqlConnection(_connectionString); }
        }

        // GET: SACTQU002
        [HttpPost]
        [Route("GetSignOffAGIDTOKEN")]
        public RespEncData GetSignOffAGIDTOKEN(ReqEncData reqmodel)
        {
            var response = new RespEncData();
            var responseObj = new ONextSignOff();
            var reqObj = new INextSign();
            string resultNowRole = string.Empty;

            try
            {
                // 檢查資料來源是否為空
                if (reqmodel == null || string.IsNullOrWhiteSpace(reqmodel.reqEncData))
                    return BuildErrorResponse("99", "請求資料為空", response, responseObj, null);

                // 解密與反序列化
                string decryptedJson = CryHelper.DecryptAES(reqmodel.reqEncData);
                reqObj = JsonConvert.DeserializeObject<INextSign>(decryptedJson);

                // Token 驗證
                bool tokenValid = GetTokenHelper.CheckToken(reqObj.Token);
                if (!tokenValid)
                    return BuildErrorResponse("99", "Token 驗證失敗", response, responseObj, reqObj);

                if (_isTest == "true")
                {
                    // 先存原本傳入資料 記錄 Log
                    LogHelper.SaveLog(new SACTAPILog
                    {
                        AGID = reqObj.AGID,
                        ApiName = "GetSignOffAGIDTOKEN",
                        Introducer = reqObj.IntroducerID,
                        Director = reqObj.DirectorID,
                        ResponseCode = "t",
                        LogMsg = "測試階段，先存原本傳入資料",
                        LogRequestData = JsonConvert.SerializeObject(reqObj),
                        LogResponseData = ""
                    });

                    // 傳入假資料要換成真的
                    var sacttest = new sacttest();

                    // 更新推介
                    sacttest = Getreqsacttest(reqObj.IntroducerID, "Introducer");
                    reqObj.IntroducerID = sacttest.o_agent_code;

                    // 更新輔導
                    sacttest = Getreqsacttest(reqObj.DirectorID, "Director");
                    reqObj.DirectorID = sacttest.o_agent_code;
                }

                // 主要邏輯
                var nextSignOff = GetNext(reqObj, out resultNowRole);

                // 回傳成功結果
                response.respEncData = EncryptToJson(nextSignOff);

                // 記錄 Log
                LogHelper.SaveLog(new SACTAPILog
                {
                    AGID = reqObj.AGID,
                    ApiName = "GetSignOffAGIDTOKEN",
                    Introducer = reqObj.IntroducerID,
                    Director = reqObj.DirectorID,
                    ResponseCode = nextSignOff.responseCode,
                    LogMsg = nextSignOff.responseObj1.NextCode == "Y" ? "此為最後一關" : $"下一關 : {resultNowRole} 下一位簽核者 : {nextSignOff.responseObj1?.NextSignOffID} {nextSignOff.responseObj1?.NextSignOffName}",
                    LogRequestData = JsonConvert.SerializeObject(reqObj),
                    LogResponseData = JsonConvert.SerializeObject(nextSignOff)
                });


                if (_isTest == "true")
                {
                    // 回傳假資料
                    var testmodel = Getresponosacttest(nextSignOff.responseObj1.NextSignOffID, resultNowRole);
                    nextSignOff.responseObj1.NextSignOffID = testmodel?.agent_code;
                    nextSignOff.responseObj1.NextSignOffName = testmodel?.agent_name;

                    // 重新加密
                    string fakeJson = JsonConvert.SerializeObject(nextSignOff, Formatting.Indented);
                    response.respEncData = CryHelper.EncryptAES(fakeJson);
                    return response;
                }
                else
                {
                    return response;
                }

            }
            catch (Exception ex)
            {
                //string msg = ex.Message == _notNextSign ? _notNextSign : "失敗";
                return BuildErrorResponse("99", ex.Message, response, responseObj, reqObj, ex);
            }
        }

        #region 私用函式

        /// <summary>
        /// 取得下一位簽核者資訊
        /// </summary>
        /// <param name="reqObj"></param>
        /// <returns></returns>
        private ONextSignOff GetNext(INextSign reqObj, out string resultNowRole)
        {
            // 判斷下一關
            var nextRole = NextRole(reqObj);

            // 建立簽核步驟清單
            var steps = nextRole.Item1;

            // 下一關 
            SignStep next = nextRole.Item2;

            // 現在關卡
            SignStep now = nextRole.Item3;

            resultNowRole = next?.Role;

            // 取得輔導人相關資訊
            var orginDto = GetOrgin(reqObj.DirectorID);

            // 確認輔導人資料是否正確
            if (orginDto.status_code != "0")
            {
                switch (orginDto.status_code)
                {
                    case "1":
                        throw new Exception($"usp_orginForSACTAPI.status_code = {orginDto.status_code} ，輔導人ID {orginDto.client_id} 已終止");
                    case "2":
                        throw new Exception($"usp_orginForSACTAPI.status_code = {orginDto.status_code} ，輔導人ID {orginDto.client_id} 不存在");
                    default:
                        break;
                }
            }

            // 決定簽核到哪個階級
            string finalRole = FinalSignoffRole(reqObj.LevelCode);
            string nextCode = "N"; // Y:最後一關 / N:還有下一關
            var signOffInfo = new SignOffInfo();

            // 判斷是否為最後一關          
            if (next == null)
            {
                // 沒有下一關
                nextCode = "Y";
            }
            else
            {
                // 找到簽約人對應的最終簽核主管
                var finalStep = steps.FirstOrDefault(s => s.Role == finalRole);
                if (finalStep == null)
                {
                    // 假如是VM 就是最後一關
                    var last = steps.Last(); // VM
                    if (now.Role == last.Role)
                    {
                        nextCode = "Y";
                    }
                }
                else if (now.Role == finalStep.Role)
                {
                    // 假如現在對應到finalRole 就是最後一關
                    nextCode = "Y";
                }
            }

            if (nextCode == "N")
            {
                signOffInfo = NextSignOffInfo(next, orginDto); // 取得下一位簽核業務員資訊

                //假如沒有找到下一位
                if (string.IsNullOrEmpty(signOffInfo.ID))
                {
                    throw new Exception(_notNextSign);
                }
            }

            // 準備回傳
            ONextSignOff respObj = new ONextSignOff();
            respObj.responseCode = "00";
            respObj.responseMsg = "成功";
            respObj.responseObj1 = new SignoffResultObj();
            respObj.responseObj1.AGID = reqObj.AGID;  // 簽約人ID        
            respObj.responseObj1.NextSignOffID = signOffInfo?.ID ?? string.Empty;
            respObj.responseObj1.NextSignOffName = signOffInfo?.Name ?? string.Empty;
            respObj.responseObj1.NextSignOffLevelName = signOffInfo?.LevelName ?? string.Empty;
            respObj.responseObj1.NextCode = nextCode;

            // 若輔導人簽核已為 Y，才回傳簽約人相關資料
            if (string.Equals(reqObj.DirectorSign, "Y", StringComparison.OrdinalIgnoreCase))
            {
                respObj.responseObj2 = new SignerInfoObj();
                respObj.responseObj2.CenterCode = orginDto?.CenterCode ?? string.Empty;
                respObj.responseObj2.CenterName = orginDto?.CenterName ?? string.Empty;
                respObj.responseObj2.WcCenterCode = orginDto?.WcCenterCode ?? string.Empty;
                respObj.responseObj2.WcCenterName = orginDto?.WcCenterName ?? string.Empty;
                respObj.responseObj2.UmCode = orginDto?.UmCode ?? string.Empty;
                respObj.responseObj2.UmName = orginDto?.UmName ?? string.Empty;
            }

            return respObj;
        }

        /// <summary>
        /// 下一位關卡
        /// </summary>
        /// <param name="reqObj"></param>
        /// <returns>簽核步驟簽單,下一關卡</returns>
        private Tuple<List<SignStep>, SignStep, SignStep> NextRole(INextSign reqObj)
        {
            // 建立簽核步驟清單，推介>=輔導>OM>SM>VM
            var steps = new List<SignStep>();

            // 推介人
            steps.Add(new SignStep { Role = "Introducer", SignFlag = reqObj.IntroducerSign, ID = reqObj.IntroducerID });

            // 輔導人
            steps.Add(new SignStep { Role = "Director", SignFlag = reqObj.DirectorSign, ID = reqObj.DirectorID });

            // OM/SM/VM
            steps.Add(new SignStep { Role = "OM", SignFlag = reqObj.OMSign });
            steps.Add(new SignStep { Role = "SM", SignFlag = reqObj.SMSign });
            steps.Add(new SignStep { Role = "VM", SignFlag = reqObj.VMSign });

            // 決定下一關
            SignStep next = null;
            foreach (var s in steps)
            {
                // 判斷現在輪到誰簽
                if (!string.Equals(s.SignFlag, "Y", StringComparison.OrdinalIgnoreCase))
                {
                    next = s;

                    break;
                }
            }

            // 現在關卡
            SignStep now = steps.LastOrDefault(s => string.Equals(s.SignFlag, "Y", StringComparison.OrdinalIgnoreCase));

            var result = Tuple.Create(steps, next, now);

            return result;
        }

        /// <summary>
        /// 將物件轉成 JSON 並加密
        /// </summary>
        private string EncryptToJson(object obj)
        {
            string json = JsonConvert.SerializeObject(obj, Formatting.Indented);
            return CryHelper.EncryptAES(json);
        }

        /// <summary>
        /// 簽約人員的稱謂代碼 決定最終簽核給誰
        /// TS -> 簽至 OM
        /// AM~UM (2~5) -> 簽至 SM
        /// OM (6) -> 簽至 VM
        /// </summary>
        /// <param name="LevelCode"></param>
        /// <returns></returns>
        private string FinalSignoffRole(string LevelCode)
        {
            string finalRole;

            switch (LevelCode)
            {
                case "1": // TS
                    finalRole = "OM";
                    break;
                case "2": // AM
                case "3": // DM
                case "4": // UM
                case "5": // UM(籌備協理B)
                    finalRole = "SM";
                    break;
                case "6": // OM
                    finalRole = "VM";
                    break;
                default:
                    finalRole = string.Empty;
                    break;
            }
            return finalRole;
        }

        /// <summary>
        /// 下一位簽核業務員 ID
        /// </summary>
        /// <param name="next"></param>
        /// <param name="orginDto">輔導人相關資訊</param>
        /// <returns></returns>
        private SignOffInfo NextSignOffInfo(SignStep next, Orgin orginDto)
        {
            var signOffInfo = new SignOffInfo();

            switch (next.Role)
            {
                case "Introducer":
                case "Director":
                    signOffInfo.ID = next.ID;
                    signOffInfo.Name = orginDto.agent_name;
                    signOffInfo.LevelName = orginDto.Ag_levelName;
                    break;
                case "OM":
                    signOffInfo.ID = string.IsNullOrEmpty(orginDto.Om_leader_proxy) ? orginDto.Om_leader : orginDto.Om_leader_proxy;
                    signOffInfo.Name = string.IsNullOrEmpty(orginDto.Om_leaderName_proxy) ? orginDto.Om_leaderName : orginDto.Om_leaderName_proxy;
                    signOffInfo.LevelName = string.IsNullOrEmpty(orginDto.Om_levelName_proxy) ? orginDto.Om_levelName : orginDto.Om_levelName_proxy;
                    break;
                case "SM":
                    signOffInfo.ID = string.IsNullOrEmpty(orginDto.Sm_leader_proxy) ? orginDto.Sm_leader : orginDto.Sm_leader_proxy;
                    signOffInfo.Name = string.IsNullOrEmpty(orginDto.Sm_leaderName_proxy) ? orginDto.Sm_leaderName : orginDto.Sm_leaderName_proxy;
                    signOffInfo.LevelName = string.IsNullOrEmpty(orginDto.Sm_levelName_proxy) ? orginDto.Sm_levelName : orginDto.Sm_levelName_proxy;
                    break;
                case "VM":
                    signOffInfo.ID = orginDto.Vm_leader;
                    signOffInfo.Name = orginDto.Vm_leaderName;
                    signOffInfo.LevelName = orginDto.Vm_levelName;
                    break;
                default:
                    signOffInfo.ID = string.Empty;
                    signOffInfo.Name = string.Empty;
                    signOffInfo.LevelName = string.Empty;
                    break;
            }

            return signOffInfo;
        }

        /// <summary>
        /// 建立錯誤回傳並記錄 Log
        /// </summary>
        private RespEncData BuildErrorResponse(string code, string message, RespEncData response, ONextSignOff responseObj, INextSign reqObj, Exception ex = null)
        {
            responseObj.responseCode = code;
            responseObj.responseMsg = message == _notNextSign ? _notNextSign : "失敗"; // 錯誤訊息等於下一位簽核者不存在才能給對方
            response.respEncData = EncryptToJson(responseObj);

            LogHelper.SaveLog(new SACTAPILog
            {
                AGID = reqObj?.AGID,
                ApiName = "GetSignOffAGIDTOKEN",
                Introducer = reqObj?.IntroducerID,
                Director = reqObj?.DirectorID,
                ResponseCode = code,
                LogMsg = ex != null ? $"錯誤訊息 : {ex}" : message,
                LogRequestData = reqObj != null ? JsonConvert.SerializeObject(reqObj) : string.Empty,
                LogResponseData = JsonConvert.SerializeObject(responseObj)
            });

            return response;
        }
        #endregion

        #region SQL
        /// <summary>
        /// 取得業務員相關資訊
        /// </summary>
        /// <param name="AgentCode"></param>
        /// <returns></returns>
        private Orgin GetOrgin(string AgentCode)
        {
            using (IDbConnection dbConnection = Connection)
            {
                dbConnection.Open();
                var parameters = new DynamicParameters();
                parameters.Add("@cliend_id", AgentCode);
                var orgin = dbConnection.QuerySingle<Orgin>("usp_orginForSACTAPI", parameters, commandType: CommandType.StoredProcedure);

                return orgin;
            }
        }

        /// <summary>
        /// 取得假資料相關資訊(回傳)
        /// </summary>
        /// <param name="AgentCode"></param>
        /// <returns></returns>
        private sacttest Getresponosacttest(string AgentCode, string Otype)
        {
            using (IDbConnection dbConnection = Connection)
            {
                dbConnection.Open();
                var parameters = new DynamicParameters();
                parameters.Add("@o_agent_code", AgentCode);
                parameters.Add("@o_type", Otype);

                var sql = @"SELECT * FROM sacttest WHERE o_agent_code = @o_agent_code AND o_type = @o_type";

                var sacttest = dbConnection.QuerySingleOrDefault<sacttest>(sql, parameters);

                return sacttest;
            }
        }

        /// <summary>
        /// 取得假資料相關資訊(傳入)
        /// </summary>
        /// <param name="AgentCode"></param>
        /// <returns></returns>
        private sacttest Getreqsacttest(string AgentCode, string Otype)
        {
            using (IDbConnection dbConnection = Connection)
            {
                dbConnection.Open();
                var parameters = new DynamicParameters();
                parameters.Add("@agent_code", AgentCode);
                parameters.Add("@o_type", Otype);

                var sql = @"SELECT * FROM sacttest WHERE agent_code = @agent_code AND o_type = @o_type";

                var sacttest = dbConnection.QuerySingleOrDefault<sacttest>(sql, parameters);

                return sacttest;
            }
        }
        #endregion

        #region 私用類別
        private class Orgin
        {
            /// <summary>
            /// 流水號
            /// </summary>
            public string ID { get; set; }

            /// <summary>
            /// 狀態代碼
            /// 0(在職):有資料;1(終止)、2(不存在):沒資料
            /// </summary>
            public string status_code { get; set; }

            /// <summary>
            /// 業務員代碼10碼
            /// </summary>
            public string client_id { get; set; }

            /// <summary>
            /// 業務員代碼12碼
            /// </summary>
            public string agent_code { get; set; }

            /// <summary>
            /// 業務員名稱
            /// </summary>
            public string agent_name { get; set; }

            /// <summary>
            /// 業務員職級
            /// </summary>
            public string Ag_level { get; set; }

            /// <summary>
            /// 業務員稱謂
            /// </summary>
            public string Ag_levelName { get; set; }

            /// <summary>
            /// 處代碼
            /// </summary>
            public string CenterCode { get; set; }

            /// <summary>
            /// 處名稱
            /// </summary>
            public string CenterName { get; set; }

            /// <summary>
            /// 通訊處代碼
            /// </summary>
            public string WcCenterCode { get; set; }

            /// <summary>
            /// 通訊處名稱
            /// </summary>
            public string WcCenterName { get; set; }

            /// <summary>
            /// 籌處代碼 (UM)
            /// </summary>
            public string UmCode { get; set; }

            /// <summary>
            /// 籌處名稱 (UM)
            /// </summary>
            public string UmName { get; set; }

            /// <summary>
            /// 處經理代碼10碼
            /// </summary>
            public string Om_leader { get; set; }

            /// <summary>
            /// 處經理名稱
            /// </summary>
            public string Om_leaderName { get; set; }

            /// <summary>
            /// 處經理職級
            /// </summary>
            public string Om_level { get; set; }

            /// <summary>
            /// 處經理稱謂
            /// </summary>
            public string Om_levelName { get; set; }

            /// <summary>
            /// 處代理人代碼10碼
            /// </summary>
            public string Om_leader_proxy { get; set; }

            /// <summary>
            /// 處代理人名稱
            /// </summary>
            public string Om_leaderName_proxy { get; set; }

            /// <summary>
            /// 處代理人職級
            /// </summary>
            public string Om_level_proxy { get; set; }

            /// <summary>
            /// 處代理人稱謂
            /// </summary>
            public string Om_levelName_proxy { get; set; }

            /// <summary>
            /// 協理代碼10碼 (SM)
            /// </summary>
            public string Sm_leader { get; set; }

            /// <summary>
            /// 協理名稱 (SM)
            /// </summary>
            public string Sm_leaderName { get; set; }

            /// <summary>
            /// 協理職級代碼
            /// </summary>
            public string Sm_level { get; set; }

            /// <summary>
            /// 協理稱謂
            /// </summary>
            public string Sm_levelName { get; set; }

            /// <summary>
            /// 協理代理人代碼10碼 (SM)
            /// </summary>
            public string Sm_leader_proxy { get; set; }

            /// <summary>
            /// 協理代理人名稱 (SM)
            /// </summary>
            public string Sm_leaderName_proxy { get; set; }

            /// <summary>
            /// 協理代理人職級代碼
            /// </summary>
            public string Sm_level_proxy { get; set; }

            /// <summary>
            /// 協理代理人稱謂
            /// </summary>
            public string Sm_levelName_proxy { get; set; }

            /// <summary>
            /// 副總代碼10碼 (VM)
            /// </summary>
            public string Vm_leader { get; set; }

            /// <summary>
            /// 副總名稱 (VM)
            /// </summary>
            public string Vm_leaderName { get; set; }

            /// <summary>
            /// 副總職級代碼
            /// </summary>
            public string Vm_level { get; set; }

            /// <summary>
            /// 副總稱謂
            /// </summary>
            public string Vm_levelName { get; set; }
        }
        private class SignStep
        {
            public string Role { get; set; }
            public string SignFlag { get; set; }
            public string ID { get; set; }
        }
        private class SignOffInfo
        {
            public string ID { get; set; }
            public string Name { get; set; }
            public string LevelName { get; set; }
        }
        private class sacttest
        {
            public string o_agent_code { get; set; }
            public string o_agent_name { get; set; }
            public string o_type { get; set; }
            public string agent_code { get; set; }
            public string agent_name { get; set; }
            public string n_type { get; set; }
        }
        #endregion
    }
}