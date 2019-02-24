using Microsoft.Xrm.Sdk.Workflow;
using Schema;
using System;
using System.Activities;

namespace JosephM.Xrm.WorkflowScheduler.Workflows
{
    public class RescheduleMonitor2 : XrmWorkflowActivityRegistration
    {
        [Output("Next Monitor Time")]
        public OutArgument<DateTime> NextMonitorTime { get; set; }

        protected override XrmWorkflowActivityInstanceBase CreateInstance()
        {
            return new RescheduleMonitor2Instance();
        }
    }

    public class RescheduleMonitor2Instance
        : WorkflowSchedulerWorkflowActivityInstance<RescheduleMonitor2>
    {
        protected override void Execute()
        {
            DateTime nextMonitorTime = DoTheWork();
            ActivityThisType.NextMonitorTime.Set(ExecutionContext, nextMonitorTime);
        }

        private DateTime DoTheWork()
        {
            WorkflowSchedulerService.CheckOtherMonitor(Target, 2);
            //calculate and return the next execution time
            var otherMonitorTime = Target.GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_nextmonitortime);
            if (!otherMonitorTime.HasValue || otherMonitorTime < DateTime.UtcNow)
                otherMonitorTime = DateTime.UtcNow.AddHours(WorkflowSchedulerService.GetMonitorPeriod());
            return otherMonitorTime.Value.AddHours(WorkflowSchedulerService.GetMonitorPeriod());
        }
    }
}
