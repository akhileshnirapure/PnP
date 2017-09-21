using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
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

    internal class SimpleProvisioner
    {
        private readonly ClientContext _context;
        private Web _web;

        readonly Guid ConfigurationValueFieldId = new Guid();
        string CONFIGURATION_VALUE_FIELD = @"";
        private List _settingsList;

        public SimpleProvisioner(ClientContext context)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            _context = context;
        }

        public void Create_Artefacts()
        {
            _web = _context.Web;

            //  create Customer list
            _web.CreateList(ListTemplateType.GenericList, "Customer", false, true, "Lists/Customer", false);
            //  create settings list
            _settingsList = _web.CreateList(ListTemplateType.GenericList, "Settings", false, true, "Lists/Settings", false);
            _context.ExecuteQueryRetry();

            //  ensure ConfigruationValue field
            if (!_settingsList.FieldExistsByName("ConfigurationValue"))
                _settingsList.Fields.AddFieldAsXml(CONFIGURATION_VALUE_FIELD, true, AddFieldOptions.AddFieldToDefaultView);
            
        }

        public void Add_Default_Setting_ListItem()
        {
            ListItemCreationInformation newItem = new ListItemCreationInformation();
            var defaultItem = _settingsList.AddItem(newItem);

            defaultItem["Title"] = "Greeting";
            defaultItem["ConfigurationValue"] = "Good Morning, ";
            defaultItem.Update();
            _context.ExecuteQueryRetry();
        }

        public void Associate_Integrated_Workflow_To_Customer_List()
        {
           
            /*
             * 
             * 
             * */


        }
    }
}
