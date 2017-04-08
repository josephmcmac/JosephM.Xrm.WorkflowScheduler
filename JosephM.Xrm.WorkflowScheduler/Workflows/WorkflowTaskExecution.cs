using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JosephM.Xrm.WorkflowScheduler.Core;
using Microsoft.Xrm.Sdk.Workflow;
using Schema;
using System.Threading;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using JosephM.Xrm.WorkflowScheduler.Services;

namespace JosephM.Xrm.WorkflowScheduler.Workflows
{
    public class WorkflowTaskExecution : XrmWorkflowActivityRegistration
    {
        [Output("Next Execution Time")]
        public OutArgument<DateTime> NextExecutionTime { get; set; }

        protected override XrmWorkflowActivityInstanceBase CreateInstance()
        {
            return new WorkflowTaskExecutionInstance();
        }
    }

    public class WorkflowTaskExecutionInstance
        : WorkflowSchedulerWorkflowActivityInstance<WorkflowTaskExecution>
    {
        protected override void Execute()
        {
            var nextExecutionTime = DoIt(IsSandboxIsolated);
            ActivityThisType.NextExecutionTime.Set(ExecutionContext, nextExecutionTime);
        }

        public DateTime DoIt(bool isSandboxIsolated)
        {
            var startedAt = DateTime.UtcNow;
            var target = XrmService.Retrieve(TargetType, TargetId);
            var thisExecutionTime = target.GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime)
                                    ?? DateTime.UtcNow;
            var waitSeconds = target.GetInt(Fields.jmcg_workflowtask_.jmcg_waitsecondspertargetworkflowcreation);
            if (!target.GetBoolean(Fields.jmcg_workflowtask_.jmcg_on))
                return thisExecutionTime;

            var targetWorkflow = target.GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_targetworkflow);
            if (!targetWorkflow.HasValue)
                throw new NullReferenceException(string.Format("Error required field {0} is empty on the target {1}", XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_targetworkflow, TargetType), XrmService.GetEntityLabel(TargetType)));
            var type = target.GetOptionSetValue(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype);
            switch (type)
            {
                case OptionSets.WorkflowTask.WorkflowExecutionType.TargetThisWorkflowTask:
                    {
                        XrmService.StartWorkflow(targetWorkflow.Value, TargetId);
                        break;
                    }
                case OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerFetchResult:
                case OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerViewResult:
                    {
                        var fetchQuery = target.GetStringField(Fields.jmcg_workflowtask_.jmcg_fetchquery);
                        if (type == OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerViewResult)
                        {
                            var savedQueryId = target.GetStringField(Fields.jmcg_workflowtask_.jmcg_targetviewid);
                            var savedQuery = XrmService.Retrieve(Entities.savedquery, new Guid(savedQueryId));
                            fetchQuery = savedQuery.GetStringField(Fields.savedquery_.fetchxml);
                        }
                        if (fetchQuery.IsNullOrWhiteSpace())
                            throw new NullReferenceException(
                                string.Format("The target {0} is set as {1} of {2} but the required field {3} is empty"
                                    , XrmService.GetEntityLabel(TargetType),
                                    XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype,
                                        TargetType)
                                    ,
                                    XrmService.GetOptionLabel(
                                        OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerFetchResult,
                                        Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype, TargetType)
                                    , XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_fetchquery, TargetType)));
                        var results = XrmService.RetrieveAllFetch(fetchQuery);
                        var numberToDo = results.Count();
                        var numberDone = 0;
                        foreach (var result in results)
                        {
                            if (isSandboxIsolated && DateTime.UtcNow - startedAt > new TimeSpan(0, 0, MaxSandboxIsolationExecutionSeconds - 10))
                                break;
                            XrmService.StartWorkflow(targetWorkflow.Value, result.Id);
                            numberDone++;
                            if (numberDone >= numberToDo)
                                break;
                            if (waitSeconds > 0)
                                Thread.Sleep(waitSeconds * 1000);
                        }
                        break;
                    }
            }
            var periodUnit = target.GetOptionSetValue(Fields.jmcg_workflowtask_.jmcg_periodperrununit);
            var periodAmount = target.GetInt(Fields.jmcg_workflowtask_.jmcg_periodperrunamount);
            var skipWeekendsAndClosures = target.GetBoolean(Fields.jmcg_workflowtask_.jmcg_skipweekendsandbusinessclosures);
            var nextExecutionTime = CalculateNextExecutionTime(thisExecutionTime, periodUnit, periodAmount, skipWeekendsAndClosures);
            return nextExecutionTime;
        }

        public DateTime CalculateNextExecutionTime(DateTime thisExecutionTime, int periodUnit, int periodAmount, bool skipWeekendsAndClosures)
        {
            var executionTime = thisExecutionTime;
            switch (periodUnit)
            {
                case OptionSets.WorkflowTask.PeriodPerRunUnit.Minutes:
                    {
                        executionTime = thisExecutionTime.AddMinutes(periodAmount);
                        break;
                    }
                case OptionSets.WorkflowTask.PeriodPerRunUnit.Hours:
                    {
                        executionTime = thisExecutionTime.AddHours(periodAmount);
                        break;
                    }
                case OptionSets.WorkflowTask.PeriodPerRunUnit.Days:
                    {
                        executionTime = thisExecutionTime.AddDays(periodAmount);
                        break;
                    }
                case OptionSets.WorkflowTask.PeriodPerRunUnit.Months:
                    {
                        executionTime = thisExecutionTime.AddMonths(periodAmount);
                        break;
                    }
                default:
                    {
                        throw new Exception(string.Format("Error there is no logic implemented for the {0} with option value of {1}", XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_periodperrununit, Entities.jmcg_workflowtask), periodAmount));
                    }
            }
            if (executionTime <= DateTime.UtcNow)
                executionTime = CalculateNextExecutionTime(DateTime.UtcNow, periodUnit, periodAmount, skipWeekendsAndClosures);

            if (!skipWeekendsAndClosures)
                return executionTime;

            //if skip weekends and closures and the calculated date is one
            //then skip to the next date
            var executionTimeUserLocal = ConvertToUserLocal(executionTime);

            //if weekend
            if (executionTimeUserLocal.DayOfWeek == DayOfWeek.Saturday
                || executionTimeUserLocal.DayOfWeek == DayOfWeek.Sunday)
                return CalculateNextExecutionTime(executionTime, periodUnit, periodAmount, skipWeekendsAndClosures);
            //if business closure
            if(IsBusinessClosure(executionTime))
                return CalculateNextExecutionTime(executionTime, periodUnit, periodAmount, skipWeekendsAndClosures);

            return executionTime;
        }

        private bool IsBusinessClosure(DateTime executionTime)
        {
            return GetClosureTimes(executionTime)
                .Any(ti => ti.Key <= LocalisationService.ConvertToTargetTime(executionTime)
                && ti.Value >= LocalisationService.ConvertToTargetTime(executionTime));
        }

        private IEnumerable<KeyValuePair<DateTime, DateTime>> _closureTimes;
        public IEnumerable<KeyValuePair<DateTime,DateTime>> GetClosureTimes(DateTime requiredStartTime)
        {
            if (_closureTimes == null)
            {
                var start = requiredStartTime.AddDays(-1);
                var end = requiredStartTime.AddYears(1);

                var query = XrmService.BuildQuery(Entities.organization, new string[0], null, null);
                var join1 = query.AddLink(Entities.calendar, Fields.organization_.businessclosurecalendarid, Fields.calendar_.calendarid);
                var join2 = join1.AddLink(Entities.calendarrule, Fields.calendar_.calendarid, Fields.calendarrule_.calendarid);
                join2.EntityAlias = "CR";
                join2.Columns = XrmService.CreateColumnSet(new string[] { Fields.calendarrule_.starttime, Fields.calendarrule_.effectiveintervalend });
                join2.LinkCriteria.AddCondition(new ConditionExpression(Fields.calendarrule_.starttime, ConditionOperator.GreaterEqual, start));
                join2.LinkCriteria.AddCondition(new ConditionExpression(Fields.calendarrule_.starttime, ConditionOperator.LessEqual, end));

                var calendarRules = XrmService.RetrieveAll(query);

                //todo consider cast to datetime
                //todo not actually utc really
                _closureTimes = calendarRules
                    .Select(c => new KeyValuePair<DateTime, DateTime>(
                        LocalisationService.ChangeUtcToLocal((DateTime)c.GetFieldValue("CR." + Fields.calendarrule_.starttime)),
                        LocalisationService.ChangeUtcToLocal((DateTime)c.GetFieldValue("CR." + Fields.calendarrule_.effectiveintervalend))))
                    .ToArray();
            }
            return _closureTimes;
        }

        private DateTime ConvertToUserLocal(DateTime executionTime)
        {
            return LocalisationService.ConvertToTargetTime(executionTime);
        }

        private int? _userTimeZoneCode;
        private int UserTimeZoneCode
        {
            get
            {
                if (!_userTimeZoneCode.HasValue)
                {
                    var userSettings = XrmService.GetFirst(Entities.usersettings, Fields.usersettings_.systemuserid, CurrentUserId, new[] { Fields.usersettings_.timezonecode });
                    if (userSettings == null)
                        throw new NullReferenceException(string.Format("Error getting {0} for user ", XrmService.GetEntityLabel(Entities.usersettings)));
                    if (userSettings.GetField(Fields.usersettings_.timezonecode) == null)
                        throw new NullReferenceException(string.Format("Error {0} is empty in the {1} record", XrmService.GetFieldLabel(Fields.usersettings_.timezonecode, Entities.usersettings), XrmService.GetEntityLabel(Entities.usersettings)));


                    _userTimeZoneCode = userSettings.GetInt(Fields.usersettings_.timezonecode);
                }
                return _userTimeZoneCode.Value;
            }
        }

        private Entity _timeZone;
        private Entity TimeZone
        {
            get
            {
                if (_timeZone == null)
                {
                    _timeZone = XrmService.GetFirst(Entities.timezonedefinition, Fields.timezonedefinition_.timezonecode, UserTimeZoneCode, new[] { Fields.timezonedefinition_.standardname });
                }
                return _timeZone;
            }
        }

        private LocalisationService _localisationService;
        private LocalisationService LocalisationService
        {
            get
            {
                if (_localisationService == null)
                {
                    _localisationService = new LocalisationService(new LocalisationSettings(TimeZone.GetStringField(Fields.timezonedefinition_.standardname)));
                }
                return _localisationService;
            }
        }
    }
}
