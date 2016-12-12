using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Schema;

namespace JosephM.Xrm.WorkflowScheduler.Services
{
    public class WorkflowSchedulerService
    {
        private XrmService XrmService { get; set; }

        public WorkflowSchedulerService(XrmService xrmService)
        {
            XrmService = xrmService;
        }

        public Guid GetContinuousWorkflowProcessId()
        {
            return new Guid("E6997F39-7726-4718-94DE-D40B1C893F9B");
        }

        public void StartNewContinuousWorkflowFor(Guid targetId)
        {
            XrmService.StartWorkflow(GetContinuousWorkflowProcessId(), targetId);
        }

        public Guid GetMonitorWorkflowProcessId()
        {
            return new Guid("5F69D869-0A3C-4DBB-871B-4BED39E20A89");
        }

        public void StartNewMonitorWorkflowFor(Guid targetId)
        {
            XrmService.StartWorkflow(GetMonitorWorkflowProcessId(), targetId);
        }

        public void StopMonitorWorkflowFor(Guid targetId)
        {
            var instances = GetMonitorInstances(targetId);
            StopRunningWorkflowInstances(instances);
        }

        public void StopContinuousWorkflowFor(Guid targetId)
        {
            var instances = GetRecurringInstances(targetId);
            StopRunningWorkflowInstances(instances);
        }

        private void StopRunningWorkflowInstances(IEnumerable<Entity> instances)
        {
            foreach (var instance in instances)
            {
                var thisStatus = instance.GetOptionSetValue(Fields.asyncoperation_.statuscode);
                var runningStatus = new[]
                {
                    OptionSets.SystemJob.StatusReason.Waiting, OptionSets.SystemJob.StatusReason.WaitingForResources,
                    OptionSets.SystemJob.StatusReason.Pausing
                };
                if (runningStatus.Contains(thisStatus))
                {
                    instance.SetOptionSetField(Fields.asyncoperation_.statecode, XrmPicklists.AsynchOperationState.Completed);
                    instance.SetOptionSetField(Fields.asyncoperation_.statuscode, OptionSets.SystemJob.StatusReason.Canceled);
                    XrmService.Update(instance, new[] { Fields.asyncoperation_.statecode, Fields.asyncoperation_.statuscode });
                }
            }
        }

        public IEnumerable<Entity> GetRecurringInstancesFailed(Guid workflowTaskId)
        {
            var lastFailedJobQuery = XrmService.BuildQuery(Entities.asyncoperation,
                new string[0],
                new[]
                {
                                new ConditionExpression(Fields.asyncoperation_.statuscode, ConditionOperator.Equal,
                                    OptionSets.SystemJob.StatusReason.Failed),
                                new ConditionExpression(Fields.asyncoperation_.regardingobjectid, ConditionOperator.Equal,
                                    workflowTaskId)
                }
                , null);
            var workflowLink = lastFailedJobQuery.AddLink(Entities.workflow, Fields.asyncoperation_.workflowactivationid,
                Fields.workflow_.workflowid);
            workflowLink.LinkCriteria.AddCondition(new ConditionExpression(Fields.workflow_.parentworkflowid, ConditionOperator.Equal, GetContinuousWorkflowProcessId()));
            return XrmService.RetrieveAll(lastFailedJobQuery);
        }

        public IEnumerable<Entity> GetMonitorInstances(Guid targetId)
        {
            var workflowId = GetMonitorWorkflowProcessId();
            return GetRunningInstances(targetId, workflowId);
        }

        public IEnumerable<Entity> GetRecurringInstances(Guid targetId)
        {
            var workflowId = GetContinuousWorkflowProcessId();
            return GetRunningInstances(targetId, workflowId);
        }

        private IEnumerable<Entity> GetRunningInstances(Guid targetId, Guid workflowId)
        {
            var conditions = new[]
            {
                new ConditionExpression(Fields.asyncoperation_.statuscode, ConditionOperator.In,
                    new object[]
                    {
                        OptionSets.SystemJob.StatusReason.Waiting, OptionSets.SystemJob.StatusReason.WaitingForResources,
                        OptionSets.SystemJob.StatusReason.InProgress, OptionSets.SystemJob.StatusReason.Pausing
                    }),
                new ConditionExpression(Fields.asyncoperation_.regardingobjectid, ConditionOperator.Equal, targetId)
            };
            var query = XrmService.BuildQuery(Entities.asyncoperation, null, conditions, null);
            var workflowLink = query.AddLink(Entities.workflow, Fields.asyncoperation_.workflowactivationid,
                Fields.workflow_.workflowid);
            workflowLink.LinkCriteria.AddCondition(new ConditionExpression(Fields.workflow_.parentworkflowid,
                ConditionOperator.Equal, workflowId));
            var instances = XrmService.RetrieveAll(query);
            return instances;
        }

        public Entity GetWorkflow(Guid guid)
        {
            return XrmService.Retrieve(Entities.workflow, guid);
        }

        public int MinimumExecutionMinutes { get { return 10; } }
    }
}
