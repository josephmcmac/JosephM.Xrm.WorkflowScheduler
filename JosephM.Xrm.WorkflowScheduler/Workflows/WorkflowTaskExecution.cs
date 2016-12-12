using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JosephM.Xrm.WorkflowScheduler.Core;
using Microsoft.Xrm.Sdk.Workflow;
using Schema;
using System.Threading;

namespace JosephM.Xrm.WorkflowScheduler.Workflows
{
    public class WorkflowTaskExecution : XrmWorkflowActivityRegistration
    {
        [Output("Next Execution Time")]
        public OutArgument<DateTime> NextExecutionTime { get; set; }

        protected override XrmWorkflowActivityInstanceBase CreateInstance()
        {
            return new WorkflowTaskExecutionInstance();
        }
    }

    public class WorkflowTaskExecutionInstance
        : WorkflowSchedulerWorkflowActivityInstance<WorkflowTaskExecution>
    {
        protected override void Execute()
        {
            var nextExecutionTime = DoIt(IsSandboxIsolated);
            ActivityThisType.NextExecutionTime.Set(ExecutionContext, nextExecutionTime);
        }

        public DateTime DoIt(bool isSandboxIsolated)
        {
            var startedAt = DateTime.UtcNow;
            var target = XrmService.Retrieve(TargetType, TargetId);
            var thisExecutionTime = target.GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime)
                                    ?? DateTime.UtcNow;
            var waitSeconds = target.GetInt(Fields.jmcg_workflowtask_.jmcg_waitsecondspertargetworkflowcreation);
            if (!target.GetBoolean(Fields.jmcg_workflowtask_.jmcg_on))
                return thisExecutionTime;

            var targetWorkflow = target.GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_targetworkflow);
            if (!targetWorkflow.HasValue)
                throw new NullReferenceException(string.Format("Error required field {0} is empty on the target {1}", XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_targetworkflow, TargetType), XrmService.GetEntityLabel(TargetType)));
            var type = target.GetOptionSetValue(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype);
            switch (type)
            {
                case OptionSets.WorkflowTask.WorkflowExecutionType.TargetThisWorkflowTask:
                    {
                        XrmService.StartWorkflow(targetWorkflow.Value, TargetId);
                        break;
                    }
                case OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerFetchResult:
                    {
                        var fetchQuery = target.GetStringField(Fields.jmcg_workflowtask_.jmcg_fetchquery);
                        if (fetchQuery.IsNullOrWhiteSpace())
                            throw new NullReferenceException(
                                string.Format("The target {0} is set as {1} of {2} but the required field {3} is empty"
                                    , XrmService.GetEntityLabel(TargetType),
                                    XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype,
                                        TargetType)
                                    ,
                                    XrmService.GetOptionLabel(
                                        OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerFetchResult,
                                        Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype, TargetType)
                                    , XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_fetchquery, TargetType)));
                        var results = XrmService.RetrieveAllFetch(fetchQuery);
                        var numberToDo = results.Count();
                        var numberDone = 0;
                        foreach (var result in results)
                        {
                            if (isSandboxIsolated && DateTime.UtcNow - startedAt > new TimeSpan(0, 0, MaxSandboxIsolationExecutionSeconds - 10))
                                break;
                            XrmService.StartWorkflow(targetWorkflow.Value, result.Id);
                            numberDone++;
                            if (numberDone >= numberToDo)
                                break;
                            if (waitSeconds > 0)
                                Thread.Sleep(waitSeconds * 1000);
                        }
                        break;
                    }
            }
            var periodUnit = target.GetOptionSetValue(Fields.jmcg_workflowtask_.jmcg_periodperrununit);
            var periodAmount = target.GetInt(Fields.jmcg_workflowtask_.jmcg_periodperrunamount);
            var nextExecutionTime = CalculateNextExecutionTime(thisExecutionTime, periodUnit, periodAmount, XrmService);
            if (nextExecutionTime <= DateTime.UtcNow)
                nextExecutionTime = CalculateNextExecutionTime(DateTime.UtcNow, periodUnit, periodAmount, XrmService);
            return nextExecutionTime;
        }

        public static DateTime CalculateNextExecutionTime(DateTime thisExecutionTime, int periodUnit, int periodAmount, XrmService xrmService)
        {
            switch (periodUnit)
            {
                case OptionSets.WorkflowTask.PeriodPerRunUnit.Minutes:
                    {
                        return thisExecutionTime.AddMinutes(periodAmount);
                    }
                case OptionSets.WorkflowTask.PeriodPerRunUnit.Hours:
                    {
                        return thisExecutionTime.AddHours(periodAmount);
                    }
                case OptionSets.WorkflowTask.PeriodPerRunUnit.Days:
                    {
                        return thisExecutionTime.AddDays(periodAmount);
                    }
                case OptionSets.WorkflowTask.PeriodPerRunUnit.Months:
                    {
                        return thisExecutionTime.AddMonths(periodAmount);
                    }
            }
            throw new Exception(string.Format("Error there is no logic implemented for the {0} with option value of {1}", xrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_periodperrununit, Entities.jmcg_workflowtask), periodAmount));
        }
    }
}
