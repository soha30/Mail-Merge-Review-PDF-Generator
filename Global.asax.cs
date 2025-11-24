using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Optimization;
using System.Web.Routing;
using System.Web.Security;
using System.Web.SessionState;

namespace ali1982ReviewMailmerge
{
    public class Global : HttpApplication
    {
        void Application_Start(object sender, EventArgs e)
        {
            //// register syncfusion : here you put your license key in the xxxxx
            //Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("xxxxx");
            
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("xxxxxxxx");

            // Code that runs on application startup
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

     
        }
    }
}