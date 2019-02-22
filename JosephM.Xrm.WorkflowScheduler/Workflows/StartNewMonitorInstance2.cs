namespace JosephM.Xrm.WorkflowScheduler.Workflows
{
    public class StartNewMonitorInstance2 : XrmWorkflowActivityRegistration
    {
        protected override XrmWorkflowActivityInstanceBase CreateInstance()
        {
            return new StartNewMonitorInstance2Instance();
        }
    }

    public class StartNewMonitorInstance2Instance
        : WorkflowSchedulerWorkflowActivityInstance<StartNewMonitorInstance2>
    {
        protected override void Execute()
        {
            DoIt();
        }

        private void DoIt()
        {
            WorkflowSchedulerService.StartNewMonitorWorkflowFor(TargetId, 2);
        }
    }
}
