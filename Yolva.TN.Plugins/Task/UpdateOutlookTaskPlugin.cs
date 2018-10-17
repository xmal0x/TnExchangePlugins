using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Yolva.TN.Plugins.BaseClasses;
using Yolva.TN.Plugins.Common;

namespace Yolva.TN.Plugins.Task
{
    public class UpdateOutlookTaskPlugin : PluginBase
    {
        /// <summary>
        /// post operation
        /// </summary>
        private string serviceUrl = "https://azuretaskwebapp.azurewebsites.net/api/tasks/UpdateTaskInOutlook";
        public UpdateOutlookTaskPlugin() : base("", "")
        {
        }
        public UpdateOutlookTaskPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }
        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            var task = pluginContext.TargetImageEntity;
            if (task == null)
                return;

            if (task.Contains("description") || task.Contains("scheduledend") || task.Contains("ownerid"))
            {
                //throw new InvalidPluginExecutionException("" + task.Contains("description") + "\n" + task.Contains("scheduledend") + "\n" + task.Contains("ownerid"));
                if (pluginContext.PluginExecutionContext.Depth > 1)
                    throw new InvalidPluginExecutionException("Depth");

                TaskEntity updateTask = new TaskEntity();
                Entity oldTask = pluginContext.OrganizationService.Retrieve("task", task.Id, new ColumnSet(true));
                if (oldTask == null)
                {
                    throw new InvalidPluginExecutionException("Error then try get old task data");
                }

                updateTask.CrmId = task.Id;

                updateTask.Subject = oldTask["subject"].ToString();
                updateTask.Body = oldTask["description"].ToString();
                updateTask.DuoDate = DateTime.Parse(oldTask["scheduledend"].ToString());
                updateTask.NewTaskOwnerId = Guid.Empty;


                if (task.Contains("subject"))
                {
                    updateTask.Subject = task["subject"].ToString();

                }
                if (task.Contains("description"))
                {
                    updateTask.Body = task["description"].ToString();

                }
                if (task.Contains("scheduledend"))
                {
                    updateTask.DuoDate = DateTime.Parse(task["scheduledend"].ToString());

                }
                if (task.Contains("ownerid"))
                {
                    updateTask.NewTaskOwnerId = ((EntityReference)task["ownerid"]).Id;

                }

                var data = UpdateTaskInOutlook(updateTask, serviceUrl);
            }
        }

        public string UpdateTaskInOutlook(TaskEntity task, string serviceUrl)
        {
            var ApiServiceUrl = serviceUrl;
            using (WebClient client = new WebClient())
            {
                client.Headers[HttpRequestHeader.ContentType] = "application/json";
                client.Encoding = Encoding.UTF8;
                var jsonObj = JsonConvert.SerializeObject(task);
                var dataString = client.UploadString(ApiServiceUrl, jsonObj);
                var data = JsonConvert.DeserializeObject(dataString);
                return data.ToString();
            }
        }
    }
}
