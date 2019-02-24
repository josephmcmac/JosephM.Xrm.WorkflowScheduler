using JosephM.Xrm.WorkflowScheduler.Core;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Schema;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JosephM.Xrm.WorkflowScheduler.Services
{
    public class WorkflowSchedulerService
    {
        private XrmService XrmService { get; set; }
        public LogController Controller { get; }

        public WorkflowSchedulerService(XrmService xrmService, LogController controller)
        {
            XrmService = xrmService;
            Controller = controller ?? new LogController();
        }

        public Guid GetContinuousWorkflowProcessId()
        {
            return new Guid("E6997F39-7726-4718-94DE-D40B1C893F9B");
        }

        public void StartNewContinuousWorkflowFor(Guid targetId)
        {
            XrmService.StartWorkflow(GetContinuousWorkflowProcessId(), targetId);
        }

        public Guid GetMonitorWorkflow1ProcessId()
        {
            return new Guid("5F69D869-0A3C-4DBB-871B-4BED39E20A89");
        }

        public Guid GetMonitorWorkflow2ProcessId()
        {
            return new Guid("4CA9C24B-BE47-4060-88F5-BBE57CA4DEEA");
        }

        public int GetMonitorPeriod()
        {
            return 1;
        }

        public void CheckOtherMonitor(Entity workflowTask, int monitorNUmber)
        {
            //while we dont want this to fail
            //if it does we would rather continue this monitor
            //than fail this one due to an error processing the error
            //so lets just ignore/supress any errors
            try
            {
                //okay for this check if the other workflow time is earlier than this one
                var otherMonitorTimeField = monitorNUmber < 2
                    ? Fields.jmcg_workflowtask_.jmcg_nextmonitortime2
                    : Fields.jmcg_workflowtask_.jmcg_nextmonitortime;
                var otherMonitorTime = workflowTask.GetDateTimeField(otherMonitorTimeField);
                if (!otherMonitorTime.HasValue
                    || otherMonitorTime.Value < DateTime.UtcNow)
                {
                    //in that case lets concel if it exists as seems to have failed
                    var otherMonitorId = monitorNUmber < 2
                        ? GetMonitorWorkflow2ProcessId()
                        : GetMonitorWorkflow1ProcessId();
                    var instancesToCancel = GetRunningInstances(workflowTask.Id, otherMonitorId);
                    StopRunningWorkflowInstances(instancesToCancel);
                    workflowTask.SetField(otherMonitorTimeField, DateTime.UtcNow.AddHours(1));
                    XrmService.Update(workflowTask, new[] { otherMonitorTimeField });
                    //and spawn it again
                    StartNewMonitorWorkflowFor(workflowTask.Id, monitorNUmber < 2 ? 2 : 1);
                }
            }
            catch(Exception ex)
            {
                //swallow the error as we dont want to fail this monitor
                Controller.LogLiteral(string.Format("Error verifying other monitor: {0}", ex.XrmDisplayString()));
            }
        }

        public void StartNewMonitorWorkflowFor(Guid targetId, int monitorNumber)
        {
            var monitorWorkflowId = monitorNumber < 2
                ? GetMonitorWorkflow1ProcessId()
                : GetMonitorWorkflow2ProcessId();
            XrmService.StartWorkflow(monitorWorkflowId, targetId);
        }

        public void StopMonitorWorkflowsFor(Guid targetId)
        {
            var instances = GetMonitorInstances(targetId);
            StopRunningWorkflowInstances(instances);
            instances = GetMonitor2Instances(targetId);
            StopRunningWorkflowInstances(instances);
        }

        public void StopContinuousWorkflowFor(Guid targetId)
        {
            var instances = GetRecurringInstances(targetId);
            StopRunningWorkflowInstances(instances);
        }

        private void StopRunningWorkflowInstances(IEnumerable<Entity> instances)
        {
            foreach (var instance in instances)
            {
                var thisStatus = instance.GetOptionSetValue(Fields.asyncoperation_.statuscode);
                var runningStatus = new[]
                {
                    OptionSets.SystemJob.StatusReason.Waiting, OptionSets.SystemJob.StatusReason.WaitingForResources,
                    OptionSets.SystemJob.StatusReason.Pausing
                };
                if (runningStatus.Contains(thisStatus))
                {
                    instance.SetOptionSetField(Fields.asyncoperation_.statecode, XrmPicklists.AsynchOperationState.Completed);
                    instance.SetOptionSetField(Fields.asyncoperation_.statuscode, OptionSets.SystemJob.StatusReason.Canceled);
                    XrmService.Update(instance, new[] { Fields.asyncoperation_.statecode, Fields.asyncoperation_.statuscode });
                }
            }
        }

        public IEnumerable<Entity> GetRecurringInstancesFailed(Guid workflowTaskId)
        {
            var lastFailedJobQuery = XrmService.BuildQuery(Entities.asyncoperation,
                new string[0],
                new[]
                {
                      new ConditionExpression(Fields.asyncoperation_.statuscode, ConditionOperator.Equal,
                                    OptionSets.SystemJob.StatusReason.Failed),
                      new ConditionExpression(Fields.asyncoperation_.regardingobjectid, ConditionOperator.Equal,
                                    workflowTaskId)
                }
                , null);
            var workflowLink = lastFailedJobQuery.AddLink(Entities.workflow, Fields.asyncoperation_.workflowactivationid,
                Fields.workflow_.workflowid);
            workflowLink.LinkCriteria.AddCondition(new ConditionExpression(Fields.workflow_.parentworkflowid, ConditionOperator.Equal, GetContinuousWorkflowProcessId()));
            return XrmService.RetrieveAll(lastFailedJobQuery);
        }

        public IEnumerable<Entity> GetMonitorInstances(Guid targetId)
        {
            return GetRunningInstances(targetId, GetMonitorWorkflow1ProcessId());
        }

        public IEnumerable<Entity> GetMonitor2Instances(Guid targetId)
        {
            return GetRunningInstances(targetId, GetMonitorWorkflow2ProcessId());
        }

        public IEnumerable<Entity> GetRecurringInstances(Guid targetId)
        {
            var workflowId = GetContinuousWorkflowProcessId();
            return GetRunningInstances(targetId, workflowId);
        }

        private IEnumerable<Entity> GetRunningInstances(Guid targetId, Guid workflowId)
        {
            var conditions = new[]
            {
                new ConditionExpression(Fields.asyncoperation_.statuscode, ConditionOperator.In,
                    new object[]
                    {
                        OptionSets.SystemJob.StatusReason.Waiting, OptionSets.SystemJob.StatusReason.WaitingForResources,
                        OptionSets.SystemJob.StatusReason.InProgress, OptionSets.SystemJob.StatusReason.Pausing
                    }),
                new ConditionExpression(Fields.asyncoperation_.regardingobjectid, ConditionOperator.Equal, targetId)
            };
            var query = XrmService.BuildQuery(Entities.asyncoperation, null, conditions, null);
            var workflowLink = query.AddLink(Entities.workflow, Fields.asyncoperation_.workflowactivationid,
                Fields.workflow_.workflowid);
            workflowLink.LinkCriteria.AddCondition(new ConditionExpression(Fields.workflow_.parentworkflowid,
                ConditionOperator.Equal, workflowId));
            var instances = XrmService.RetrieveAll(query);
            return instances;
        }

        public Entity GetWorkflow(Guid guid)
        {
            return XrmService.Retrieve(Entities.workflow, guid);
        }

        public int MinimumExecutionMinutes { get { return 10; } }

        public string GetUserAppId(Guid userId, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName))
                return null;
            var query = XrmService.BuildQuery(Entities.team, new[] { fieldName }, new[] {
                        new ConditionExpression(fieldName, ConditionOperator.NotNull)
            }, null);
            var userMemberLink = query.AddLink(Relationships.team_.teammembership_association.EntityName, Fields.team_.teamid, Fields.team_.teamid);
            userMemberLink.LinkCriteria.AddCondition(new ConditionExpression(Fields.systemuser_.systemuserid, ConditionOperator.Equal, userId));
            var results = XrmService.RetrieveAll(query);
            return results.Count() == 1 ? results.First().GetStringField(fieldName) : null;
        }

        public string GetQueueAppId(Guid queueId, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName))
                return null;
            var query = XrmService.BuildQuery(Entities.team, new[] { fieldName }, new[] {
                        new ConditionExpression(fieldName, ConditionOperator.NotNull),
                        new ConditionExpression(Fields.team_.queueid, ConditionOperator.Equal, queueId)
            }, null);
            var results = XrmService.RetrieveAll(query);
            return results.Count() == 1 ? results.First().GetStringField(fieldName) : null;
        }

        public TimeSpan? GetStartTimeSpan(Func<string, object> getField)
        {
            TimeSpan? timeSpan = null;
            if (XrmEntity.GetBoolean(getField(Fields.jmcg_workflowtask_.jmcg_onlyrunbetweenhours)))
            {
                var hour = XrmEntity.GetOptionSetValue(getField(Fields.jmcg_workflowtask_.jmcg_starthour));
                var minute = XrmEntity.GetOptionSetValue(getField(Fields.jmcg_workflowtask_.jmcg_startminute));
                var ampm = XrmEntity.GetOptionSetValue(getField(Fields.jmcg_workflowtask_.jmcg_startampm));
                if (hour != -1 && minute != -1 && ampm != -1)
                {
                    if (hour == 12)
                        hour = 0;
                    if (ampm == OptionSets.WorkflowTask.StartAMPM.PM)
                        hour = hour + 12;
                    timeSpan = new TimeSpan(hour, minute, 0);
                }
            }
            return timeSpan;
        }

        public TimeSpan? GetEndTimeSpan(Func<string, object> getField)
        {
            TimeSpan? timeSpan = null;
            if (XrmEntity.GetBoolean(getField(Fields.jmcg_workflowtask_.jmcg_onlyrunbetweenhours)))
            {
                var hour = XrmEntity.GetOptionSetValue(getField(Fields.jmcg_workflowtask_.jmcg_endhour));
                var minute = XrmEntity.GetOptionSetValue(getField(Fields.jmcg_workflowtask_.jmcg_endminute));
                var ampm = XrmEntity.GetOptionSetValue(getField(Fields.jmcg_workflowtask_.jmcg_endampm));
                if (hour != -1 && minute != -1 && ampm != -1)
                {
                    if (hour == 12)
                        hour = 0;
                    if (ampm == OptionSets.WorkflowTask.StartAMPM.PM)
                        hour = hour + 12;
                    timeSpan = new TimeSpan(hour, minute, 0);
                }
            }
            return timeSpan;
        }
    }
}
