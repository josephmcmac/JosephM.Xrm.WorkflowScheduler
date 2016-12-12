using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Workflow;
using Schema;

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
            WorkflowSchedulerService.StartNewMonitorWorkflowFor(TargetId);
        }
    }
}
