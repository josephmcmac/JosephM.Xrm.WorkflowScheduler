using Microsoft.Xrm.Sdk.Workflow;
using Schema;
using System;
using System.Activities;

namespace JosephM.Xrm.WorkflowScheduler.Workflows
{
    public class RescheduleMonitor : XrmWorkflowActivityRegistration
    {
        [Output("Next Monitor Time")]
        public OutArgument<DateTime> NextMonitorTime { get; set; }

        protected override XrmWorkflowActivityInstanceBase CreateInstance()
        {
            return new RescheduleMonitorInstance();
        }
    }

    public class RescheduleMonitorInstance
        : WorkflowSchedulerWorkflowActivityInstance<RescheduleMonitor>
    {
        protected override void Execute()
        {
            DateTime nextMonitorTime = DoTheWork();
            ActivityThisType.NextMonitorTime.Set(ExecutionContext, nextMonitorTime);
        }

        private DateTime DoTheWork()
        {
            WorkflowSchedulerService.CheckOtherMonitor(Target, 1);
            //calculate and return the next execution time
            var otherMonitorTime = Target.GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_nextmonitortime2);
            if (!otherMonitorTime.HasValue || otherMonitorTime < DateTime.UtcNow)
                otherMonitorTime = DateTime.UtcNow.AddHours(WorkflowSchedulerService.GetMonitorPeriod());
            return otherMonitorTime.Value.AddHours(WorkflowSchedulerService.GetMonitorPeriod());
        }
    }
}
