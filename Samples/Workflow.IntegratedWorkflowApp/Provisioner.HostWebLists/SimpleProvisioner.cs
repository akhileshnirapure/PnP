using System;
using System.Linq;
using Microsoft.SharePoint.Client;

namespace Provisioner.HostWebLists
{
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
            if (!_web.ListExists("Customer"))
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
           
            /*  Facts
             *  
             *  1.  In Integrated Workflow App, the Workflow definition is published in App Web.
             *  2.  Tasks and History List needs to be in App Web.
             *  3.  For creating Subscription 'AppWeb' instance should be used (which can be fetched from HostWeb by means of Workflow AppName)
             *  
             * */

            // Target List workflow association is created as Extension method (refer that for details)

            //  Take reference of target list to which the workflow needs to be associated

            var customerList = _web.GetListByTitle("Customer");





        }
    }


    public static class ListExtensions
    {
        public static void AssociateIntegratedWorkflow(this List currentList, string appWebLeafUrl, string historyListName, string taskListName, string workflowDefinitionName)
        {
            //  fetch app-web by 'App name'

            var hostWeb = currentList.ParentWeb;
            var context = hostWeb.Context;

            //  validate App-Web exists by leafname

            if(string.IsNullOrEmpty(appWebLeafUrl))
                throw new ArgumentNullException(appWebLeafUrl,"Please provide appweb leaf url.");
            if(!hostWeb.WebExists(appWebLeafUrl))
                throw new ArgumentNullException(appWebLeafUrl,"Invalid app-web leaf url provided. Hint: provide Add-In Name from AppMainfest.xml");


            
            var appWeb = hostWeb.GetWeb(appWebLeafUrl);
            
            List historyList, tasksList;

            if (!appWeb.ListExists(historyListName))
            {
                historyList = appWeb.CreateList(ListTemplateType.WorkflowHistory, historyListName, false);
                context.ExecuteQueryRetry();
            }
            else
            {
                historyList = appWeb.GetListByTitle(historyListName);
            }

            if (!appWeb.ListExists(taskListName))
            {
                tasksList = appWeb.CreateList(ListTemplateType.Tasks, taskListName, false);
                context.ExecuteQueryRetry();
            }
            else
            {
                tasksList = appWeb.GetListByTitle(taskListName);
            }

            //  validate if workflow definition exists
            var publishedDefinitions = appWeb.GetWorkflowDefinitions(true);

            if(publishedDefinitions.Any(p=> string.Compare(p.DisplayName,workflowDefinitionName,StringComparison.InvariantCultureIgnoreCase) != 0))
                throw new ArgumentException("Invalid workflow definition name.",workflowDefinitionName);
            

            var wfDefinition = appWeb.GetWorkflowDefinition(workflowDefinitionName);


        }
    }
}