using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JosephM.Xrm.WorkflowScheduler.Core;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Schema;

namespace JosephM.Xrm.WorkflowScheduler.Plugins
{
    public class WorkflowTaskPlugin : WorkflowSchedulerPlugin
    {
        private int MinimumExecutionMinutes
        {
            get { return WorkflowSchedulerService.MinimumExecutionMinutes; }
        }

        public override void GoExtention()
        {
            VerifyPeriod();
            ValidateTarget();
            SpawnOrTurnOffRecurrance();
        }

        private void SpawnOrTurnOffRecurrance()
        {
            if (IsMessage(PluginMessage.Create, PluginMessage.Update) && IsStage(PluginStage.PostEvent) &&
                IsMode(PluginMode.Synchronous))
            {
                if (BooleanChangingToTrue(Fields.jmcg_workflowtask_.jmcg_on))
                {
                    if (!GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime).HasValue)
                        SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, DateTime.UtcNow);
                    var workflows = WorkflowSchedulerService.GetRecurringInstances(TargetId);
                    if (!workflows.Any())
                        WorkflowSchedulerService.StartNewContinuousWorkflowFor(TargetId);
                }
                if (BooleanChangingToFalse(Fields.jmcg_workflowtask_.jmcg_on))
                {
                    WorkflowSchedulerService.StopContinuousWorkflowFor(TargetId);
                }
            }
        }

        private void ValidateTarget()
        {
            if (IsMessage(PluginMessage.Create, PluginMessage.Update) && IsStage(PluginStage.PreOperationEvent))
            {
                if (FieldChanging(Fields.jmcg_workflowtask_.jmcg_targetworkflow,
                    Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype
                    , Fields.jmcg_workflowtask_.jmcg_fetchquery))
                {
                    var requiredTargetedType = Entities.jmcg_workflowtask;
                    //if a fetch target validate the fetch and get its target type
                    if (GetOptionSet(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype) == OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerFetchResult)
                    {
                        var fetch = GetStringField(Fields.jmcg_workflowtask_.jmcg_fetchquery);
                        if (fetch.IsNullOrWhiteSpace())
                            throw new InvalidPluginExecutionException(string.Format("{0} is required when {1} is {2}", GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_fetchquery),
                                GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype), GetOptionLabel(OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerFetchResult, Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype)));
                        try
                        {
                            requiredTargetedType = XrmService.ConvertFetchToQueryExpression(fetch).EntityName.ToLower();
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidPluginExecutionException(string.Format("There was an error validating {0} it could not be converted to a {1}", GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_fetchquery), typeof(QueryExpression).Name), ex);
                        }
                    }
                    var targetWorkflowId = GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_targetworkflow);
                    if (!targetWorkflowId.HasValue)
                        throw new InvalidPluginExecutionException(string.Format("{0} is required", GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_targetworkflow)));
                    var workflow = WorkflowSchedulerService.GetWorkflow(targetWorkflowId.Value);
                    if (workflow.GetStringField(Fields.workflow_.primaryentity) != requiredTargetedType)
                        throw new InvalidPluginExecutionException(
                            string.Format(
                                "Error the {0} targets the entity type of {1} but was expected to target the type {2}"
                                , GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_targetworkflow),
                                workflow.GetStringField(Fields.workflow_.primaryentity), requiredTargetedType));
                }
            }
        }

        private void VerifyPeriod()
        {
            if (IsMessage(PluginMessage.Create, PluginMessage.Update) && IsStage(PluginStage.PreOperationEvent))
            {
                var amount = GetIntField(Fields.jmcg_workflowtask_.jmcg_periodperrunamount);
                var unit = GetOptionSet(Fields.jmcg_workflowtask_.jmcg_periodperrununit);
                if (amount <= 0)
                    throw new InvalidPluginExecutionException(
                        string.Format("{0} is required to be greater than or equal to zero",
                            GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_periodperrununit)));
                if (unit == OptionSets.WorkflowTask.PeriodPerRunUnit.Minutes && amount < MinimumExecutionMinutes)
                    throw new InvalidPluginExecutionException(string.Format("{0} is required to be at least {1} {2}",
                        GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_periodperrununit), MinimumExecutionMinutes,
                        GetOptionLabel(OptionSets.WorkflowTask.PeriodPerRunUnit.Minutes,
                            Fields.jmcg_workflowtask_.jmcg_periodperrununit)));
            }
        }
    }
}
