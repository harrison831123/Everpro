using System.Web.Mvc;

namespace EP.SD.SalesSupport.LAW.Web
{
    public class LAWAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "LAW";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "LAW_default",
                "LAW/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}