using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Schema;
using JosephM.Xrm.WorkflowScheduler.Emails;
using Microsoft.Xrm.Sdk;

namespace JosephM.Xrm.WorkflowScheduler.Workflows
{
    public class TargetWorkflowTaskMonitor : XrmWorkflowActivityRegistration
    {
        [Output("Has New Failures")]
        public OutArgument<bool> HasNewFailures { get; set; }

        protected override XrmWorkflowActivityInstanceBase CreateInstance()
        {
            return new TargetWorkflowTaskMonitorInstance();
        }
    }

    public class TargetWorkflowTaskMonitorInstance
        : WorkflowSchedulerWorkflowActivityInstance<TargetWorkflowTaskMonitor>
    {
        protected override void Execute()
        {
            var failedJobs = GetfailedJobs();
            var hasNewFailure = failedJobs.Any();
            ActivityThisType.HasNewFailures.Set(ExecutionContext, hasNewFailure);
            if (hasNewFailure && Target.GetBoolean(Fields.jmcg_workflowtask_.jmcg_sendnotificationfortargetfailures))
            {
                var recipientQueue = Target.GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto);
                if (!recipientQueue.HasValue)
                    throw new NullReferenceException(string.Format("Error required field {0} is empty on the target {1}", XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto, TargetType), XrmService.GetEntityLabel(TargetType)));

                var crmUrl = GetCrmURL();
                var email = new HtmlEmailGenerator(XrmService, crmUrl);
                email.AppendParagraph(string.Format("This is an automated notification that workflows triggered by the {0} named '{1}' have failed. You will need to review the failures to fix any issues"
                    , XrmService.GetEntityLabel(TargetType)
                    , Target.GetStringField(Fields.jmcg_workflowtask_.jmcg_name)));
                if (!string.IsNullOrWhiteSpace(crmUrl))
                {
                    email.AppendParagraph(email.CreateHyperlink(email.CreateUrl(Target), "Open " + XrmService.GetEntityLabel(TargetType)));
                }
                email.AppendTable(failedJobs, LocalisationService, GetFieldsToDisplayInNotificationEmail());

                var subject = XrmService.GetEntityLabel(TargetType) + " Target Failures: " + Target.GetStringField(Fields.jmcg_workflowtask_.jmcg_name);
                SendNotificationEmail(Entities.queue, recipientQueue.Value, subject, email.GetContent());

                Target.SetField(Fields.jmcg_workflowtask_.jmcg_minimumtargetfailuredatetime, DateTime.UtcNow);
                XrmService.Update(Target, new[] { Fields.jmcg_workflowtask_.jmcg_minimumtargetfailuredatetime });
            }
        }



        public bool HasNewFailure()
        {
            var failedJobs = GetfailedJobs();
            return failedJobs.Any();
        }

        private IEnumerable<Entity> GetfailedJobs()
        {
            var lastFailure = Target.GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_minimumtargetfailuredatetime)
                ?? new DateTime(1910, 01, 01);
            var targetWorkflow = Target.GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_targetworkflow)
                                 ?? Guid.Empty;

            var lastFailedJobQuery = XrmService.BuildQuery(Entities.asyncoperation,
                GetFieldsToDisplayInNotificationEmail(),
                new[]
                {
                    new ConditionExpression(Fields.asyncoperation_.statuscode, ConditionOperator.Equal,
                        OptionSets.SystemJob.StatusReason.Failed),
                    new ConditionExpression(Fields.asyncoperation_.modifiedon, ConditionOperator.GreaterThan,
                        lastFailure)
                }, null);
            var workflowLink = lastFailedJobQuery.AddLink(Entities.workflow, Fields.asyncoperation_.workflowactivationid,
                Fields.workflow_.workflowid);
            workflowLink.LinkCriteria.AddCondition(new ConditionExpression(Fields.workflow_.parentworkflowid, ConditionOperator.Equal, targetWorkflow));
            return XrmService.RetrieveFirstX(lastFailedJobQuery, HtmlEmailGenerator.MaximumNumberOfEntitiesToList + 1); //add 1 to determine if exceeded the limit
        }

        private IEnumerable<string> GetFieldsToDisplayInNotificationEmail()
        {
            return new[] {
                    Fields.asyncoperation_.regardingobjectid,
                    Fields.asyncoperation_.friendlymessage
                };
        }
    }
}
