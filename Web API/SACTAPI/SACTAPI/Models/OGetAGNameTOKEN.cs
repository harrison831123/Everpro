namespace SACTAPI.Models
{
    /// <summary>
    /// 需求單號：
    /// 傳送輔導人姓名與ID給廠商
    /// </summary>
    public class OGetAGNameTOKEN
    {
        public OGetAGNameTOKEN()
        {
            responseCode = string.Empty;
            responseMsg = string.Empty;
            responseObj = null;
        }

        public string responseCode { get; set; }//回傳代碼00成功、99失敗
        public string responseMsg { get; set; }//回傳訊息
        public RegisterData responseObj;//成功物件：輔導人資料
    }
}