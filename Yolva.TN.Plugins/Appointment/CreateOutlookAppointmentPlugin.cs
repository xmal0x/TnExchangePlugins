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

namespace Yolva.TN.Plugins.Appointment
{
    /// <summary>
    /// Post create
    /// </summary>
    public class CreateOutlookAppointmentPlugin : PluginBase
    {
        private string serviceUrl = "https://azuretaskwebapp.azurewebsites.net/api/Appointments/CreateAppointmentInOutlook";
        public CreateOutlookAppointmentPlugin() : base("", "")
        {
        }
        public CreateOutlookAppointmentPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }

        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            var appointment = pluginContext.TargetImageEntity;

            if (appointment == null)
                return;

            AppointmentEntity newAppointment = new AppointmentEntity()
            {
                Body = string.Empty,
                End = null,
                Location = string.Empty,
                OutlookId = string.Empty,
                Start = null,
                Subject = string.Empty,
                RequiredAttendeesEmails = new List<string>()
            };

            newAppointment.CrmId = appointment.Id;
            if (appointment.Contains("subject"))
                newAppointment.Subject = appointment["subject"].ToString();
            if (appointment.Contains("location"))
                newAppointment.Location = appointment["location"].ToString();
            if (appointment.Contains("description"))
                newAppointment.Body = appointment["description"].ToString();
            if (appointment.Contains("scheduledstart"))
                newAppointment.Start = DateTime.Parse(appointment["scheduledstart"].ToString());
            if (appointment.Contains("scheduledend"))
                newAppointment.End = DateTime.Parse(appointment["scheduledend"].ToString());
            if (appointment.Contains("ylv_outlookid"))
                newAppointment.OutlookId = appointment["ylv_outlookid"].ToString();

            if (appointment.Contains("ylv_emailness"))
            {
                foreach (var attendees in appointment["ylv_emailness"].ToString().Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                    newAppointment.RequiredAttendeesEmails.Add(attendees.Trim());
            }
            if (string.IsNullOrEmpty(newAppointment.OutlookId))
            {
                var attendeesCollection = appointment.GetAttributeValue<EntityCollection>("requiredattendees");
                foreach (var att in attendeesCollection.Entities)
                {
                    var responserRef = (EntityReference)att["partyid"];
                    if (responserRef.LogicalName == "contact")
                    {
                        var response = new Entity("ylv_response");
                        response["ylv_name"] = appointment["subject"].ToString();
                        response["ylv_contact"] = responserRef;
                        response["ylv_responsetype"] = new OptionSetValue((int)ResponseType.Unknown);
                        response["ylv_responses"] = appointment.ToEntityReference();
                        response.Id = pluginContext.OrganizationService.Create(response);
                    }
                    else if (responserRef.LogicalName == "systemuser")
                    {
                        var response = new Entity("ylv_response");
                        response["ylv_name"] = appointment["subject"].ToString();
                        response["ylv_sysuser"] = responserRef;
                        response["ylv_responsetype"] = new OptionSetValue((int)ResponseType.Unknown);
                        response["ylv_responses"] = appointment.ToEntityReference();
                        response.Id = pluginContext.OrganizationService.Create(response);
                    }
                }
            }

            var data = CreateAppointmentInOutlook(newAppointment, serviceUrl);
            
            Entity crmAppointment = new Entity("appointment");
            crmAppointment.Id = appointment.Id;
            crmAppointment["ylv_outlookid"] = data.Trim('"');
            pluginContext.OrganizationService.Update(crmAppointment);
            
        }
        public static string CreateAppointmentInOutlook(AppointmentEntity appointment, string serviceUrl)
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
