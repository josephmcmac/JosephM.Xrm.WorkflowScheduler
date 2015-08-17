using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JosephM.Xrm.WorkflowScheduler.Core;
using Microsoft.Xrm.Sdk.Workflow;
using Schema;

namespace JosephM.Xrm.WorkflowScheduler.Workflows
{
    public class StartNewContinuousInstance : XrmWorkflowActivityRegistration
    {
        protected override XrmWorkflowActivityInstanceBase CreateInstance()
        {
            return new StartNewContinuousInstanceInstance();
        }
    }

    public class StartNewContinuousInstanceInstance
        : WorkflowSchedulerWorkflowActivityInstance<StartNewContinuousInstance>
    {
        protected override void Execute()
        {
            DoIt();
        }

        private void DoIt()
        {
            WorkflowSchedulerService.StartNewContinuousWorkflowFor(TargetId);
        }
    }
}
