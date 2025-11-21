using System.Web.Mvc;
 
namespace EP.PSL.WorkResources.MeetingMng.Web
{
    public class MeetingMngAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "MeetingMng";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "MeetingMng_default",
                "MeetingMng/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}