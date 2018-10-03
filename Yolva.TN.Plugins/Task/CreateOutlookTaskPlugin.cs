﻿using Microsoft.Xrm.Sdk;
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
    public class CreateOutlookTaskPlugin : PluginBase
    {
        private string serviceUrl = "http://azuretaskappweb.azurewebsites.net/api/tasks/CreateTaskInOutlook";
        public CreateOutlookTaskPlugin() : base("", "")
        {
        }
        public CreateOutlookTaskPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }
        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            
            var task = pluginContext.TargetImageEntity;
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
            if (task.Contains("ownerid"))
            {
                newTask.NewTaskOwnerId = ((EntityReference)task["ownerid"]).Id;
            }

            var data = CreateTaskInOutlook(newTask, serviceUrl);
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
