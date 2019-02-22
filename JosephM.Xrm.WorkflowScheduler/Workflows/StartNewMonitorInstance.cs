namespace JosephM.Xrm.WorkflowScheduler.Workflows
{
    public class StartNewMonitorInstance : XrmWorkflowActivityRegistration
    {
        protected override XrmWorkflowActivityInstanceBase CreateInstance()
        {
            return new StartNewMonitorInstanceInstance();
        }
    }

    public class StartNewMonitorInstanceInstance
        : WorkflowSchedulerWorkflowActivityInstance<StartNewMonitorInstance>
    {
        protected override void Execute()
        {
            DoIt();
        }

        private void DoIt()
        {
            WorkflowSchedulerService.StartNewMonitorWorkflowFor(TargetId, 1);
        }
    }
}
