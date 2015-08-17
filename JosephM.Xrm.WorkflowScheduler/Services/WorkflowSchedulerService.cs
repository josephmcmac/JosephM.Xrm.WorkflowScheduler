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

        public void StopContinuousWorkflowFor(Guid targetId)
        {
            var instances = GetRecurringInstances(targetId);
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

        public IEnumerable<Entity> GetRecurringInstances(Guid targetId)
        {
            var conditions = new[]
            {
                new ConditionExpression(Fields.asyncoperation_.statuscode, ConditionOperator.In, new object[] { OptionSets.SystemJob.StatusReason.Waiting,OptionSets.SystemJob.StatusReason.WaitingForResources,OptionSets.SystemJob.StatusReason.InProgress, OptionSets.SystemJob.StatusReason.Pausing }),
                new ConditionExpression(Fields.asyncoperation_.regardingobjectid, ConditionOperator.Equal, targetId)
            };
            var query = XrmService.BuildQuery(Entities.asyncoperation, null, conditions, null);
            var workflowLink = query.AddLink(Entities.workflow, Fields.asyncoperation_.workflowactivationid, Fields.workflow_.workflowid);
            workflowLink.LinkCriteria.AddCondition(new ConditionExpression(Fields.workflow_.parentworkflowid, ConditionOperator.Equal, GetContinuousWorkflowProcessId()));
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
