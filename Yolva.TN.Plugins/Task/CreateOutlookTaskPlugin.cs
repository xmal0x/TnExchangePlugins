using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Net;
using System.Text;
using Yolva.TN.Plugins.BaseClasses;
using Yolva.TN.Plugins.Common;

namespace Yolva.TN.Plugins.Task
{
    public class CreateOutlookTaskPlugin : PluginBase
    {
        /// <summary>
        /// post operat
        /// </summary>
        private string serviceUrl = "https://azuretaskwebapp.azurewebsites.net/api/tasks/CreateTaskInOutlook";
        public CreateOutlookTaskPlugin() : base("", "")
        {
        }
        public CreateOutlookTaskPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }
        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            //throw new InvalidPluginExecutionException("bolt");
            var task = pluginContext.TargetImageEntity;
            if (!string.IsNullOrEmpty(task.GetAttributeValue<string>("ylv_outlookid")))
                return;
            if (task == null)
                return;
            TaskEntity newTask = new TaskEntity()
            {
                Body = string.Empty,
                CrmId = task.Id,
                DuoDate = null,
                OwnerId = pluginContext.UserId,
                NewTaskOwnerId = Guid.Empty,
                OutlookId = string.Empty,
                Subject = string.Empty,
                TaskStatus = TaskStatusCode.Open
            };

            if (task.Contains("subject"))
            {
                newTask.Subject = "CRM task: " + (string)task["subject"];
            }
            if (task.Contains("description"))
            {
                newTask.Body = task["description"].ToString();
            }
            if (task.Contains("scheduledend"))
            {
                newTask.DuoDate = DateTime.Parse(task["scheduledend"].ToString());
            }
            else
            {
                newTask.DuoDate = DateTime.Now.AddDays(1);
            }
            if (task.Contains("ownerid"))
            {
                newTask.NewTaskOwnerId = ((EntityReference)task["ownerid"]).Id;
            }

            var data = CreateTaskInOutlook(newTask, serviceUrl);
            //todo j
            Entity crmTask = new Entity("task");
            crmTask.Id = task.Id;
            crmTask["ylv_outlookid"] = data.Trim('"');
            pluginContext.OrganizationService.Update(crmTask);
            
        }
        public static string CreateTaskInOutlook(TaskEntity task, string serviceUrl)
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
