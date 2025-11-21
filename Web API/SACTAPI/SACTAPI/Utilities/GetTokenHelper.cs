using System;

namespace SACTAPI.Utilities
{
    public static class GetTokenHelper
    {
        public static Boolean CheckToken(string strToken)
        {
            //Token 組成方式 固定8碼 + 西元年月日
            //EX : 12684149 + 當天日期(YYYYMMDD)
            DateTime Logtime = DateTime.Now;
            String token = "12684149" + DateTime.Now.ToString("yyyyMMdd");
            Boolean result = strToken == token ? true : false;
            return result;
        }        
    }
}