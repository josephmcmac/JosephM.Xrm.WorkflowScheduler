using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Schema;

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
            var hasNewFailure = HasNewFailure();
            ActivityThisType.HasNewFailures.Set(ExecutionContext, hasNewFailure);
        }

        public bool HasNewFailure()
        {
            var target = XrmService.Retrieve(TargetType, TargetId);
            var lastFailure = target.GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_minimumtargetfailuredatetime)
                ?? new DateTime(1910, 01, 01);
            var targetWorkflow = target.GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_targetworkflow)
                                 ?? Guid.Empty;

            var lastFailedJobQuery = XrmService.BuildQuery(Entities.asyncoperation,
                new string[0],
                new[]
                {
                    new ConditionExpression(Fields.asyncoperation_.statuscode, ConditionOperator.Equal,
                        OptionSets.SystemJob.StatusReason.Failed),
                    new ConditionExpression(Fields.asyncoperation_.modifiedon, ConditionOperator.GreaterThan,
                        lastFailure)
                }
                , null);
            var workflowLink = lastFailedJobQuery.AddLink(Entities.workflow, Fields.asyncoperation_.workflowactivationid,
                Fields.workflow_.workflowid);
            workflowLink.LinkCriteria.AddCondition(new ConditionExpression(Fields.workflow_.parentworkflowid, ConditionOperator.Equal, targetWorkflow));
            var failedJob = XrmService.RetrieveFirst(lastFailedJobQuery);
            return failedJob != null;
        }
    }
}
