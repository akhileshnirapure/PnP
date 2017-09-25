using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core;

namespace Provisioner.HostWebLists
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
             * 1.   Create Two Lists (Customer, Settings) based on Generic List Template.
             * 2.   Add additional Single line column called 'ConfigValue' in Settings list
             * 3.   Add single item in Settings List.
             * 4.   Associate Integrated Workflow App with Customer list.
             * 
             * */

            var targetSiteUrl = ConfigurationManager.AppSettings["SITE_URL"];
            var userName = ConfigurationManager.AppSettings["USER_NAME"];
            var password = ConfigurationManager.AppSettings["PASSWORD"];



            if (Validate(targetSiteUrl, userName, password) == false)
                throw new Exception("Ensure valid SiteUrl,Username and Password is provided.");


            //  get client context
            var context = new AuthenticationManager()
                             .GetSharePointOnlineAuthenticatedContextTenant(targetSiteUrl,userName,password);

            //  keeping things simple for demo
            var provisioner = new SimpleProvisioner(context);

            provisioner.Create_Artefacts();
            provisioner.Add_Default_Setting_ListItem();
            provisioner.Associate_Integrated_Workflow_To_Customer_List();


        }

        private static bool Validate(string targetSiteUrl, string userName, string password)
        {
            return !string.IsNullOrEmpty(targetSiteUrl) && !string.IsNullOrEmpty(userName) &&
                   !string.IsNullOrEmpty(password);
        }
    }
}
