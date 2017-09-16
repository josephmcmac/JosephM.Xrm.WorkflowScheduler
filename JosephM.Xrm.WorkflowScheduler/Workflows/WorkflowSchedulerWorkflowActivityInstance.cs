using JosephM.Xrm.WorkflowScheduler.Core;
using JosephM.Xrm.WorkflowScheduler.Extentions;
using JosephM.Xrm.WorkflowScheduler.Services;
using Microsoft.Xrm.Sdk;
using Schema;
using System;
using System.Linq;

namespace JosephM.Xrm.WorkflowScheduler.Workflows
{
    public abstract class WorkflowSchedulerWorkflowActivityInstance<T> : XrmWorkflowActivityInstance<T>
        where T : XrmWorkflowActivityRegistration
    {
        private WorkflowSchedulerService _workflowSchedulerService;
        public WorkflowSchedulerService WorkflowSchedulerService
        {
            get
            {
                if (_workflowSchedulerService == null)
                    _workflowSchedulerService = new WorkflowSchedulerService(XrmService);
                return _workflowSchedulerService;
            }
        }

        private Entity _target;
        public Entity Target
        {
            get
            {
                if (_target == null)
                {
                    _target = XrmService.Retrieve(TargetType, TargetId);
                }
                return _target;
            }
        }

        public void SendNotificationEmail(string recipientType, Guid recipientId, string subject, string content)
        {
            var email = new Entity(Entities.email);
            email.AddFromParty(Entities.queue, GetNotificationSendingQueue());
            email.AddToParty(recipientType, recipientId);

            email.SetField(Fields.email_.subject, subject.Left(XrmService.GetMaxLength(Fields.email_.subject, Entities.email)));
            email.SetField(Fields.email_.description, content);
            email.SetLookupField(Fields.email_.regardingobjectid, TargetId, TargetType);

            //create and send the email
            var emailId = XrmService.Create(email);
            XrmService.SendEmail(emailId);
        }

        public string GetCrmURL()
        {
            var baseUrl = Target.GetStringField(Fields.jmcg_workflowtask_.jmcg_crmbaseurl);
            if (baseUrl != null && !baseUrl.StartsWith("http") && baseUrl.Contains("."))
            {
                var split = baseUrl.Split('.');
                var type = split.First();
                var field = split.ElementAt(1);
                var query = XrmService.BuildSourceQuery(type, new[] { field });
                var firstRecord = XrmService.RetrieveFirst(query);
                baseUrl = firstRecord.GetStringField(field);
            }

            return baseUrl;
        }

        private Guid GetNotificationSendingQueue()
        {
            var queueId = Target.GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom);
            if (!queueId.HasValue)
                throw new NullReferenceException(string.Format("Error required field {0} is empty on the target {1}", XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom, TargetType), XrmService.GetEntityLabel(TargetType)));
            return queueId.Value;
        }
    }
}