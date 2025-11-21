using System.Web.Mvc;

namespace EP.SD.SalesSupport.CUSCRM.Web
{
    public class CUSCRMAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "CUSCRM";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "CUSCRM_default",
                "CUSCRM/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}