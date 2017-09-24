using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;

namespace Provisioner.HostWebLists
{
    public static class ListIntegratedWorkflowExtensions
    {
        public static void AssociateIntegratedWorkflow(this List currentList, 
                            string appWebLeafName, string historyListName, string taskListName, 
                            string workflowDefinitionName, string subscriptionName, bool startManually, 
                            bool startOnCreate, bool startOnChange, string statusFieldName, 
                            Dictionary<string, string> associationValues = null)

        {
            var hostWeb = currentList.ParentWeb;
            var context = hostWeb.Context;

            //  validate App-Web exists by leafname
            //  the app-web url leaf name e.g. https://yourtenant.sharepoint.com/appwebname
            //  hint: this referes to 'Name' property in AppManifest.xml

            if(string.IsNullOrEmpty(appWebLeafName))
                throw new ArgumentNullException(appWebLeafName,"Please provide appweb leaf name.");
            if(!hostWeb.WebExists(appWebLeafName))
                throw new ArgumentNullException(appWebLeafName,"Invalid app-web leaf name provided. Hint: provide Add-In Name from AppMainfest.xml");


            //  get app-web reference as we need 
            //  1.  fetch the workflow definition
            //  2.  associate with history and tasks list in App-Web if exists else create
            var appWeb = hostWeb.GetWeb(appWebLeafName);
            
            List historyList, tasksList;

            if (!appWeb.ListExists(historyListName))
            {
                historyList = appWeb.CreateList(ListTemplateType.WorkflowHistory, historyListName, false);
                historyList.EnsureProperty(p => p.Id);
                context.ExecuteQueryRetry();
            }
            else
            {
                historyList = appWeb.GetListByTitle(historyListName);
                historyList.EnsureProperty(p => p.Id);

            }

            if (!appWeb.ListExists(taskListName))
            {
                tasksList = appWeb.CreateList(ListTemplateType.Tasks, taskListName, false);
                tasksList.EnsureProperty(p => p.Id);
            }
            else
            {
                tasksList = appWeb.GetListByTitle(taskListName);
                tasksList.EnsureProperty(p => p.Id);
            }

            //  get published workflow defintions from app-web
            var publishedDefinitions = appWeb.GetWorkflowDefinitions(true);
            
            //  validate workflow definition exists in app-web
            if (publishedDefinitions.Any(p=> string.Compare(p.DisplayName,workflowDefinitionName,StringComparison.InvariantCultureIgnoreCase) != 0))
                throw new ArgumentException("Invalid workflow definition name.",workflowDefinitionName);
            
            //  todo:validate existing subscription with host-web list

            var existingSubscriptions = currentList.GetWorkflowSubscription(subscriptionName);


            //  fetch the workflow definition for association
            var wfDefinition = appWeb.GetWorkflowDefinition(workflowDefinitionName);


            var sub = new WorkflowSubscription(currentList.Context)
            {
                DefinitionId = wfDefinition.Id,
                Enabled = true,
                Name = subscriptionName,
                StatusFieldName = statusFieldName
            };


            //  add workflow events
            var eventTypes = new List<string>();
            if (startManually) eventTypes.Add("WorkflowStart");
            if (startOnCreate) eventTypes.Add("ItemAdded");
            if (startOnChange) eventTypes.Add("ItemUpdated");

            sub.EventTypes = eventTypes;

            //  set mandatory properties
            //  the history and tasks list should be referred from AppWeb
            sub.SetProperty("HistoryListId", historyList.Id.ToString());
            sub.SetProperty("TaskListId", tasksList.Id.ToString());

            //  set custom values
            //  these values (key/value pair) can be accessed in workflow using GetConfiguration Workflow activity
            if (associationValues != null)
            {
                foreach (var key in associationValues.Keys)
                {
                    sub.SetProperty(key, associationValues[key]);
                }
            }

            var servicesManager = new WorkflowServicesManager(currentList.Context, appWeb);

            var subscriptionService = servicesManager.GetWorkflowSubscriptionService();

            var subscriptionResult = subscriptionService.PublishSubscriptionForList(sub, currentList.Id);

            currentList.Context.ExecuteQueryRetry();
            

        }
    }
}