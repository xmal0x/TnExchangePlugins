using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Yolva.TN.Plugins.Common
{
    public class TaskEntity : IntegrationEntity
    {
        public TaskEntity()
        {
            CrmEntityLogicalName = "task";
        }
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime? DuoDate { get; set; }
        public Guid OwnerId { get; set; }
        public Guid NewTaskOwnerId { get; set; }
        public TaskStatusCode? TaskStatus { get; set; }
    }
}