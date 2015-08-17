using JosephM.Xrm.WorkflowScheduler.Plugins;
using Schema;

namespace JosephM.Xrm.WorkflowScheduler
{
    public class JosephMPluginRegistration : XrmPluginRegistration
    {
        public override XrmPlugin CreateEntityPlugin(string entityType, bool isRelationship)
        {
            switch (entityType)
            {
                case Entities.jmcg_workflowtask: return new WorkflowTaskPlugin();
            }
            return null;
        }
    }
}