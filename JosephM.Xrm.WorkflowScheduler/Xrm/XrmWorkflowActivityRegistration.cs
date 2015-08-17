using System.Activities;

namespace JosephM.Xrm.WorkflowScheduler
{
    public abstract class XrmWorkflowActivityRegistration : CodeActivity
    {
        protected override sealed void Execute(CodeActivityContext executionContext)
        {
            var instance = CreateInstance();
            instance.ExecuteBase(executionContext, this);
        }

        protected abstract XrmWorkflowActivityInstanceBase CreateInstance();
    }
}