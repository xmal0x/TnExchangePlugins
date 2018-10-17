using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yolva.TN.Plugins.BaseClasses;

namespace Yolva.TN.Plugins.Appointment
{
    public class CheckUpdateAppointmentNecessityPlugin : PluginBase
    {
        public CheckUpdateAppointmentNecessityPlugin() : base("", "")
        {
        }
        public CheckUpdateAppointmentNecessityPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }
        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            
            //var preAppointment = pluginContext.PreImageEntity;
            var appointment = pluginContext.TargetImageEntity;
            if(string.IsNullOrEmpty(appointment.GetAttributeValue<string>("location")))
                throw new InvalidPluginExecutionException("Location not update");

            // throw new InvalidPluginExecutionException("Location Pre: " + preAppointment.GetAttributeValue<string>("location") + "\n: " + appointment.GetAttributeValue<string>("location"));

            //f (preAppointment.GetAttributeValue<string>("location") == appointment.GetAttributeValue<string>("location"))
            //throw new InvalidPluginExecutionException("Location Pre: ");
                
        }
    }
}
