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

        private string _crmUrl;
        public string GetCrmURL()
        {
            if (_crmUrl == null)
            {
                _crmUrl = Target.GetStringField(Fields.jmcg_workflowtask_.jmcg_crmbaseurl);
                if (_crmUrl != null && !_crmUrl.StartsWith("http") && _crmUrl.Contains("."))
                {
                    var split = _crmUrl.Split('.');
                    var type = split.First();
                    var field = split.ElementAt(1);
                    var query = XrmService.BuildSourceQuery(type, new[] { field });
                    var firstRecord = XrmService.RetrieveFirst(query);
                    _crmUrl = firstRecord.GetStringField(field);
                }
            }
            return _crmUrl;
        }

        private Guid GetNotificationSendingQueue()
        {
            var queueId = Target.GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom);
            if (!queueId.HasValue)
                throw new NullReferenceException(string.Format("Error required field {0} is empty on the target {1}", XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom, TargetType), XrmService.GetEntityLabel(TargetType)));
            return queueId.Value;
        }

        private int? _userTimeZoneCode;
        private int UserTimeZoneCode
        {
            get
            {
                if (!_userTimeZoneCode.HasValue)
                {
                    var userSettings = XrmService.GetFirst(Entities.usersettings, Fields.usersettings_.systemuserid, CurrentUserId, new[] { Fields.usersettings_.timezonecode });
                    if (userSettings == null)
                        throw new NullReferenceException(string.Format("Error getting {0} for user ", XrmService.GetEntityLabel(Entities.usersettings)));
                    if (userSettings.GetField(Fields.usersettings_.timezonecode) == null)
                        throw new NullReferenceException(string.Format("Error {0} is empty in the {1} record", XrmService.GetFieldLabel(Fields.usersettings_.timezonecode, Entities.usersettings), XrmService.GetEntityLabel(Entities.usersettings)));


                    _userTimeZoneCode = userSettings.GetInt(Fields.usersettings_.timezonecode);
                }
                return _userTimeZoneCode.Value;
            }
        }

        private Entity _timeZone;
        private Entity TimeZone
        {
            get
            {
                if (_timeZone == null)
                {
                    _timeZone = XrmService.GetFirst(Entities.timezonedefinition, Fields.timezonedefinition_.timezonecode, UserTimeZoneCode, new[] { Fields.timezonedefinition_.standardname });
                }
                return _timeZone;
            }
        }

        private LocalisationService _localisationService;
        public LocalisationService LocalisationService
        {
            get
            {
                if (_localisationService == null)
                {
                    _localisationService = new LocalisationService(new LocalisationSettings(TimeZone.GetStringField(Fields.timezonedefinition_.standardname)));
                }
                return _localisationService;
            }
        }
    }
}