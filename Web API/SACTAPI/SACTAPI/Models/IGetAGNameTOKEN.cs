namespace SACTAPI.Models
{
    public class IGetAGNameTOKEN
    {
        public IGetAGNameTOKEN()
        {
            broker = string.Empty;
            token = string.Empty;
            insurer = string.Empty;
            RegisterNo = string.Empty;
            introducerID = string.Empty;
        }

        public string broker { get; set; }//保經公司代碼
        public string token { get; set; }//TOKEN
        public string insurer { get; set; }//廠商代碼
        public string RegisterNo { get; set; }//壽險登錄證字號
        public string introducerID { get; set; }//推介人ID
    }
}