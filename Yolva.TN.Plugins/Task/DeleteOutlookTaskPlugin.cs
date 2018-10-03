using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Yolva.TN.Plugins.BaseClasses;
using Yolva.TN.Plugins.Common;

namespace Yolva.TN.Plugins.Task
{
    public class DeleteOutlookTaskPlugin : PluginBase
    {
        private string serviceUrl = "http://azuretaskappweb.azurewebsites.net/api/tasks/DeleteTaskInOutlook";
        public DeleteOutlookTaskPlugin() : base("", "")
        {
        }
        public DeleteOutlookTaskPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }

        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            var taskRef = pluginContext.TargetImageEntityReference;
            if (taskRef == null)
                return;
            var taskEntity = pluginContext.OrganizationService.Retrieve(taskRef.LogicalName, taskRef.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("ylv_outlookid"));
            if (taskEntity.Contains("ylv_outlookid"))
                DeleteTaskInOutlook(new TaskEntity { OutlookId = taskEntity["ylv_outlookid"].ToString() }, serviceUrl);

        }

        public static string DeleteTaskInOutlook(TaskEntity outlookId, string serviceUrl)
        {
            var ApiServiceUrl = serviceUrl;
            using (WebClient client = new WebClient())
            {
                client.Headers[HttpRequestHeader.ContentType] = "application/json";
                client.Encoding = Encoding.UTF8;
                var jsonObj = JsonConvert.SerializeObject(outlookId);
                var dataString = client.UploadString(ApiServiceUrl, jsonObj);
                var data = JsonConvert.DeserializeObject(dataString);
                return data.ToString();
            }
        }
    }
}
