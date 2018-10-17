using System;
using System.Net;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using Yolva.TN.Plugins.BaseClasses;
using Yolva.TN.Plugins.Common;

namespace Yolva.TN.Plugins.Appointment
{
    public class UpdateOutlookAppointmentPlugin : PluginBase
    {

        private string serviceUrl = "https://azuretaskwebapp.azurewebsites.net/api/Appointments/UpdateAppointmentInOutlook";
        public UpdateOutlookAppointmentPlugin() : base("", "")
        {
        }
        public UpdateOutlookAppointmentPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }
        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            var appointment = pluginContext.TargetImageEntity;
            if (appointment == null)
                return;

            if (appointment.Contains("description") || appointment.Contains("scheduledstart") || appointment.Contains("scheduledend"))
            {
                if (pluginContext.PluginExecutionContext.Depth > 1)
                    throw new InvalidPluginExecutionException("Depth");

                AppointmentEntity updateAppointment = new AppointmentEntity();
                Entity oldAppointment = pluginContext.OrganizationService.Retrieve("appointment", appointment.Id, new ColumnSet(true));
                if (oldAppointment == null)
                {
                    throw new InvalidPluginExecutionException("Error then try get old appointment data");
                }

                updateAppointment.CrmId = appointment.Id;

                updateAppointment.Body = oldAppointment["description"].ToString();
                updateAppointment.Start = DateTime.Parse(oldAppointment["scheduledstart"].ToString());
                updateAppointment.End = DateTime.Parse(oldAppointment["scheduledend"].ToString());

                if (appointment.Contains("description"))
                {
                    updateAppointment.Body = appointment["description"].ToString();
                }
                if (appointment.Contains("scheduledstart"))
                {
                    updateAppointment.Start = DateTime.Parse(appointment["scheduledstart"].ToString());
                }
                if (appointment.Contains("scheduledend"))
                {
                    updateAppointment.End = DateTime.Parse(appointment["scheduledend"].ToString());
                }
                var data = UpdateTaskInOutlook(updateAppointment, serviceUrl);
            }
        }

        public string UpdateTaskInOutlook(AppointmentEntity appointment, string serviceUrl)
        {
            var ApiServiceUrl = serviceUrl;
            using (WebClient client = new WebClient())
            {
                client.Headers[HttpRequestHeader.ContentType] = "application/json";
                client.Encoding = Encoding.UTF8;
                var jsonObj = JsonConvert.SerializeObject(appointment);
                var dataString = client.UploadString(ApiServiceUrl, jsonObj);
                var data = JsonConvert.DeserializeObject(dataString);
                return data.ToString();
            }
        }
    }
}
