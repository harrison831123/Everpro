using System.Web.Mvc;

namespace EP.SD.SalesZone.AGUPG.Web.Areas.AGUPG
{
    public class AGUPGAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "AGUPG";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "AGUPG_default",
                "AGUPG/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}