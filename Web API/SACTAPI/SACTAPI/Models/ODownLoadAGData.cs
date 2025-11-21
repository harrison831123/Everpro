namespace SACTAPI.Models
{
    /// <summary>
    /// 需求單號：
    /// 回覆下載簽約資料是否成功
    /// </summary>
    public class ODownLoadAGData
    {
        public ODownLoadAGData()
        {
            responseCode = string.Empty;
            responseMsg = string.Empty;
        }
        public string responseCode { get; set; }//回傳代碼
        public string responseMsg { get; set; }//回傳訊息       
    }
}