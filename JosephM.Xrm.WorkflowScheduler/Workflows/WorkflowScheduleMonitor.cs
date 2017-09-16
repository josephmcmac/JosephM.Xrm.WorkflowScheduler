using JosephM.Xrm.WorkflowScheduler.Emails;
using Microsoft.Xrm.Sdk.Workflow;
using Schema;
using System;
using System.Activities;
using System.Linq;

namespace JosephM.Xrm.WorkflowScheduler.Workflows
{
    public class WorkflowScheduleMonitor : XrmWorkflowActivityRegistration
    {
        [Output("Is Behind Schedule")]
        public OutArgument<bool> IsBehindSchedule { get; set; }

        protected override XrmWorkflowActivityInstanceBase CreateInstance()
        {
            return new WorkflowScheduleMonitorInstance();
        }
    }

    public class WorkflowScheduleMonitorInstance
        : WorkflowSchedulerWorkflowActivityInstance<WorkflowScheduleMonitor>
    {
        protected override void Execute()
        {
            var behindSchedule = IsBehindSchedule();
            ActivityThisType.IsBehindSchedule.Set(ExecutionContext, behindSchedule);
            if(behindSchedule && !WorkflowSchedulerService.GetRecurringInstances(TargetId).Any())
            {
                try
                {
                    WorkflowSchedulerService.StartNewContinuousWorkflowFor(TargetId);
                }
                catch(Exception ex)
                {
                    Trace(string.Format("Error starting monitor: {0}", ex.XrmDisplayString()));
                }
            }
            if(behindSchedule && Target.GetBoolean(Fields.jmcg_workflowtask_.jmcg_sendnotificationforschedulefailures))
            {
                var recipientQueue = Target.GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto);
                if (!recipientQueue.HasValue)
                    throw new NullReferenceException(string.Format("Error required field {0} is empty on the target {1}", XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto, TargetType), XrmService.GetEntityLabel(TargetType)));

                var crmUrl = GetCrmURL();
                //var entityType = GetViewFetchAsQuery().EntityName;
                //var viewHyperlink = string.Format("<a href={0}>{0}</a>", baseUrl);
                //var pStyle = "style='font-family: Arial,sans-serif;font-size: 12pt;padding:6.15pt 6.15pt 6.15pt 6.15pt'";
                //var content =
                //    string.Format(@"<p {0}>This is an automated notification there are '{1}' to be actioned</p>
                //                <p {0}>Please review and process the records</p>
                //                <p {0}>{2}</p>", pStyle, View.GetStringField(Fields.savedquery_.name), viewHyperlink);
                //exlude primary key and fields in linked entities in list because label
                var email = new HtmlEmailGenerator(XrmService, crmUrl);
                email.AppendParagraph(string.Format("This is an automated notification that the {0} named '{1}' fell behind schedule. The system has attempted to restart the continuous workflow but any failed instance should be reviewed"
                    , XrmService.GetEntityLabel(TargetType)
                    , Target.GetStringField(Fields.jmcg_workflowtask_.jmcg_name)));
                if (!string.IsNullOrWhiteSpace(crmUrl))
                {
                    email.AppendParagraph(email.CreateHyperlink(email.CreateUrl(Target), "Open " + XrmService.GetEntityLabel(TargetType)));
                }

                var subject = XrmService.GetEntityLabel(TargetType) + " Schedule Failure: " + Target.GetStringField(Fields.jmcg_workflowtask_.jmcg_name);
                SendNotificationEmail(Entities.queue, recipientQueue.Value, subject, email.GetContent());

                Target.SetField(Fields.jmcg_workflowtask_.jmcg_minimumschedulefailuredatetime, DateTime.UtcNow);
                XrmService.Update(Target, new[] { Fields.jmcg_workflowtask_.jmcg_minimumschedulefailuredatetime });
            }
        }

        public bool IsBehindSchedule()
        {
            var nextExecutionTime = Target.GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime);
            var threshold = Target.GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_minimumschedulefailuredatetime);
            var on = Target.GetBoolean(Fields.jmcg_workflowtask_.jmcg_on);

            //one hour leeway
            return on
                 && nextExecutionTime.HasValue
                && (!threshold.HasValue || threshold.Value < nextExecutionTime.Value)
                && DateTime.UtcNow.AddHours(-1) > nextExecutionTime.Value;
        }
    }
}
