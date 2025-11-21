namespace SACTAPI.Models
{
    /// <summary>
    /// 需求單號：
    /// 輔導人資料
    /// </summary>
    public class RegisterData
    {
        public RegisterData()
        {
            RegisterNo = string.Empty;
            DirectorID = string.Empty;
            AgName = string.Empty;
            directorLevel = string.Empty;
            introducerName = string.Empty;
            introducerLevel = string.Empty;
        }
        public string RegisterNo { get; set; }//壽險登錄證字號 
        public string DirectorID { get; set; }//輔導人ID
        public string AgName { get; set; }//輔導人姓名
        public string directorLevel { get; set; }//輔導人稱謂
        public string introducerName { get; set; }//推介人姓名
        public string introducerLevel { get; set; }//推介人稱謂
    }
}