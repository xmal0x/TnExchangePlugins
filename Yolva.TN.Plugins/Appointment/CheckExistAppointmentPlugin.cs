using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yolva.TN.Plugins.BaseClasses;

namespace Yolva.TN.Plugins.Appointment
{
    /// <summary>
    /// pre create
    /// </summary>
    public class CheckExistAppointmentPlugin : PluginBase
    {
        public CheckExistAppointmentPlugin() : base("", "")
        {
        }
        public CheckExistAppointmentPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }
        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            var appointment = pluginContext.TargetImageEntity;
            //throw new InvalidPluginExecutionException((appointment == null).ToString());
            if (!string.IsNullOrEmpty(appointment.GetAttributeValue<string>("ylv_outlookid")))
                throw new InvalidPluginExecutionException("appointment exist");
        }
    }
}
