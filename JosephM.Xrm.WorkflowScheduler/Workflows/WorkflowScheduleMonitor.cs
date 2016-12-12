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
        }

        public bool IsBehindSchedule()
        {
            var target = XrmService.Retrieve(TargetType, TargetId);
            var nextExecutionTime = target.GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime);
            var threshold = target.GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_minimumschedulefailuredatetime);
            var on = target.GetBoolean(Fields.jmcg_workflowtask_.jmcg_on);


            //one hour leeway
            return on
                 && nextExecutionTime.HasValue
                && (!threshold.HasValue || threshold.Value < nextExecutionTime.Value)
                && DateTime.UtcNow.AddHours(-1) > nextExecutionTime.Value;
        }
    }
}
