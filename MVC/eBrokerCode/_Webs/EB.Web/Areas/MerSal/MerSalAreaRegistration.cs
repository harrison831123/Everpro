using System.Web.Mvc;

namespace EB.SL.MerSal.Web.Areas.MerSal
{
    public class MerSalAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "MerSal";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "MerSal_default",
                "MerSal/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}