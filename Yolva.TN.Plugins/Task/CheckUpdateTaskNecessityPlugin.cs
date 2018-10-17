using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yolva.TN.Plugins.BaseClasses;
using Microsoft.Xrm.Sdk;

namespace Yolva.TN.Plugins.Task
{
    public class CheckUpdateTaskNecessityPlugin : PluginBase
    {
        /// <summary>
        /// pre operation
        /// </summary>
        public CheckUpdateTaskNecessityPlugin() : base("", "")
        {
        }
        public CheckUpdateTaskNecessityPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }
        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            var task = pluginContext.TargetImageEntity;
            if (!task.Contains("description") && !task.Contains("scheduledend") && !task.Contains("ownerid"))
                throw new InvalidPluginExecutionException("task not need update");
        }
    }
}
