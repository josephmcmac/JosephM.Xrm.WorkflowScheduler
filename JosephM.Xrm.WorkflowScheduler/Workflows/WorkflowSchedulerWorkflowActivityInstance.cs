using JosephM.Xrm.WorkflowScheduler.Services;

namespace JosephM.Xrm.WorkflowScheduler.Workflows
{
    public abstract class WorkflowSchedulerWorkflowActivityInstance<T> : XrmWorkflowActivityInstance<T>
        where T : XrmWorkflowActivityRegistration
    {
        private WorkflowSchedulerService _workflowSchedulerService;
        public WorkflowSchedulerService WorkflowSchedulerService
        {
            get
            {
                if (_workflowSchedulerService == null)
                    _workflowSchedulerService = new WorkflowSchedulerService(XrmService);
                return _workflowSchedulerService;
            }
        }
    }
}