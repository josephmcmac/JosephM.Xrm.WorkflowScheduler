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
            TurnOffIfDeactivated();
            VerifyPeriod();
            ValidateTarget();
            SpawnOrTurnOffRecurrance();
            ValidateNotifications();
            ResetThresholdsWhenMonitorTurnedOn();
            SpawnMonitorInstance();
        }

        private void TurnOffIfDeactivated()
        {
            if (IsMessage(PluginMessage.Create, PluginMessage.Update) && IsStage(PluginStage.PreOperationEvent) &&
                IsMode(PluginMode.Synchronous))
            {
                if (OptionSetChangedTo(Fields.jmcg_workflowtask_.statecode, XrmPicklists.State.Inactive))
                    SetField(Fields.jmcg_workflowtask_.jmcg_on, false);
            }
        }

        private void SpawnMonitorInstance()
        {
            if (IsMessage(PluginMessage.Create, PluginMessage.Update) && IsStage(PluginStage.PostEvent) &&
                IsMode(PluginMode.Synchronous))
            {
                if (BooleanChangingToTrue(Fields.jmcg_workflowtask_.jmcg_on))
                {
                    WorkflowSchedulerService.StartNewMonitorWorkflowFor(TargetId);
                }
                if (BooleanChangingToFalse(Fields.jmcg_workflowtask_.jmcg_on))
                {
                    WorkflowSchedulerService.StopMonitorWorkflowFor(TargetId);
                }
            }
        }

        private void ResetThresholdsWhenMonitorTurnedOn()
        {
            if (IsMessage(PluginMessage.Create, PluginMessage.Update) && IsStage(PluginStage.PreOperationEvent) &&
                IsMode(PluginMode.Synchronous))
            {
                //set these when turned on so only subsequent failures picked up
                if (BooleanChangingToTrue(Fields.jmcg_workflowtask_.jmcg_sendnotificationforschedulefailures)
                    && !TargetEntity.Contains(Fields.jmcg_workflowtask_.jmcg_minimumschedulefailuredatetime))
                    SetField(Fields.jmcg_workflowtask_.jmcg_minimumschedulefailuredatetime, DateTime.UtcNow);
                if (BooleanChangingToTrue(Fields.jmcg_workflowtask_.jmcg_sendnotificationfortargetfailures)
                    && !TargetEntity.Contains(Fields.jmcg_workflowtask_.jmcg_minimumtargetfailuredatetime))
                    SetField(Fields.jmcg_workflowtask_.jmcg_minimumtargetfailuredatetime, DateTime.UtcNow);
            }
        }


        private void ValidateNotifications()
        {
            if (IsMessage(PluginMessage.Create, PluginMessage.Update) && IsStage(PluginStage.PreOperationEvent) &&
                IsMode(PluginMode.Synchronous))
            {
                //the email queues are required if either of the notificatins are on
                if (FieldChanging(Fields.jmcg_workflowtask_.jmcg_sendnotificationfortargetfailures
                    , Fields.jmcg_workflowtask_.jmcg_sendnotificationforschedulefailures
                    , Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom
                    , Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto))
                {
                    var eitherNotificationOn = GetBoolean(
                        Fields.jmcg_workflowtask_.jmcg_sendnotificationfortargetfailures)
                                               ||
                                               GetBoolean(
                                                   Fields.jmcg_workflowtask_.jmcg_sendnotificationforschedulefailures);
                    if (eitherNotificationOn)
                    {
                        if (GetField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom) == null)
                            throw new NullReferenceException(string.Format("{0} Is Required",
                                GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom)));
                        if (GetField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto) == null)
                            throw new NullReferenceException(string.Format("{0} Is Required",
                                GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto)));
                    }
                }
            }
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
                    , Fields.jmcg_workflowtask_.jmcg_fetchquery
                    , Fields.jmcg_workflowtask_.jmcg_targetviewid))
                {
                    var requiredTargetedType = Entities.jmcg_workflowtask;
                    //if a fetch target validate the fetch and get its target type
                    var type = GetOptionSet(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype);
                    switch (type)
                    {
                        case OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerFetchResult:
                        case OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerViewResult:
                            {
                                var fetch = GetStringField(Fields.jmcg_workflowtask_.jmcg_fetchquery);
                                if (type == OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerViewResult)
                                {
                                    var savedQueryId = GetStringField(Fields.jmcg_workflowtask_.jmcg_targetviewid);
                                    if (string.IsNullOrWhiteSpace(savedQueryId))
                                        throw new InvalidPluginExecutionException(string.Format("{0} is required when {1} is {2}", GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_targetviewid),
                                                GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype), GetOptionLabel(OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerViewResult, Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype)));
                                    var savedQuery = XrmService.Retrieve(Entities.savedquery, new Guid(savedQueryId));
                                    fetch = savedQuery.GetStringField(Fields.savedquery_.fetchxml);
                                    if (!XrmEntity.FieldsEqual(GetField(Fields.jmcg_workflowtask_.jmcg_targetviewselectedname), savedQuery.GetStringField(Fields.savedquery_.name)))
                                        SetField(Fields.jmcg_workflowtask_.jmcg_targetviewselectedname, savedQuery.GetStringField(Fields.savedquery_.name));
                                }
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
                                break;
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
