using System.Collections.Generic;
using JosephM.Xrm.WorkflowScheduler.Services;

namespace JosephM.Xrm.WorkflowScheduler.Test
{
    public abstract class JosephMXrmTest : XrmTest
    {
        protected override IEnumerable<string> EntitiesToDelete
        {
            get { return base.EntitiesToDelete; }
        }

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