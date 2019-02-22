using JosephM.Xrm.WorkflowScheduler.Services;
using Schema;

namespace JosephM.Xrm.WorkflowScheduler.Plugins
{
    public class WorkflowSchedulerPlugin : XrmEntityPlugin
    {
        private WorkflowSchedulerService _workflowSchedulerService;
        public WorkflowSchedulerService WorkflowSchedulerService
        {
            get
            {
                if (_workflowSchedulerService == null)
                    _workflowSchedulerService = new WorkflowSchedulerService(XrmService, Controller);
                return _workflowSchedulerService;
            }
        }
    }
}