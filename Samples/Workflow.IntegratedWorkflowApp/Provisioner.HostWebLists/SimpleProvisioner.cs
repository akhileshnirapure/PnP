using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Workflow;

namespace Provisioner.HostWebLists
{
    internal class SimpleProvisioner
    {
        private readonly ClientContext _context;
        private Web _web;

        string CONFIGURATION_VALUE_FIELD = @"<Field Type='Text' DisplayName='ConfigurationValue' Required='FALSE' MaxLength='255' ID='{870192fd-663b-4ea3-a972-efea5d2ea5b8}' StaticName='ConfigurationValue' Name='ConfigurationValue' />";
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
            if (!_web.ListExists("Settings"))
                _settingsList = _web.CreateList(ListTemplateType.GenericList, "Settings", false, true, "Lists/Settings", false);

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

            var settingsList = _web.GetListByTitle("Settings");

            var customerList = _web.GetListByTitle("Customer");

            string appWebLeafName = "Configurations";           //  leaf name of Add-In
            string appWebHistoryListName = "CustomerHistory";   //  history List name
            string appWebTasksListName = "CustomerTasks";       //  tasks list name
            string appWebIntegratedWorkflowName = "Greeter";    //  Workflow Name
            string listWorkflowAssociationName = "Awesome Greeting";    //  List & Workflow Association Name
            string coolStatusFieldName = "Greeting";            //  Status Column name

            var additionalConfigurations = new Dictionary<string, string>
            {
                {"Config_SettingsListId", settingsList.Id.ToString()}
            };

            customerList.AssociateIntegratedWorkflow(appWebLeafName,appWebHistoryListName,appWebTasksListName,appWebIntegratedWorkflowName,listWorkflowAssociationName,false,true,true,coolStatusFieldName,additionalConfigurations);




        }
    }
}