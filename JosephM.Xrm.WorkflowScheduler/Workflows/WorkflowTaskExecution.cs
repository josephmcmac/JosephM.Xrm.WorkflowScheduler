using JosephM.Xrm.WorkflowScheduler.Core;
using JosephM.Xrm.WorkflowScheduler.Emails;
using JosephM.Xrm.WorkflowScheduler.Extentions;
using JosephM.Xrm.WorkflowScheduler.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Schema;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;

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
            var thisExecutionTime = Target.GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime)
                                    ?? DateTime.UtcNow;
            var waitSeconds = Target.GetInt(Fields.jmcg_workflowtask_.jmcg_waitsecondspertargetworkflowcreation);
            if (!Target.GetBoolean(Fields.jmcg_workflowtask_.jmcg_on))
                return thisExecutionTime;

            var type = Target.GetOptionSetValue(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype);
            var targetWorkflow = Target.GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_targetworkflow);
            if (type != OptionSets.WorkflowTask.WorkflowExecutionType.ViewNotification
                && !targetWorkflow.HasValue)
                throw new NullReferenceException(string.Format("Error required field {0} is empty on the target {1}", XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_targetworkflow, TargetType), XrmService.GetEntityLabel(TargetType)));

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
                        var fetchQuery = Target.GetStringField(Fields.jmcg_workflowtask_.jmcg_fetchquery);
                        if (type == OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerViewResult)
                        {
                            fetchQuery = View.GetStringField(Fields.savedquery_.fetchxml);
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
                            XrmService.StartWorkflow(targetWorkflow.Value, result.Id);
                            numberDone++;
                            if (numberDone >= numberToDo)
                                break;
                            if (isSandboxIsolated && ((DateTime.UtcNow - startedAt) > new TimeSpan(0, 0, MaxSandboxIsolationExecutionSeconds - (waitSeconds + 5))))
                                break;
                            if (waitSeconds > 0)
                                Thread.Sleep(waitSeconds * 1000);
                        }
                        break;
                    }
                case OptionSets.WorkflowTask.WorkflowExecutionType.ViewNotification:
                    {
                        SendViewNotifications(Target);
                        break;
                    }
            }
            var periodUnit = Target.GetOptionSetValue(Fields.jmcg_workflowtask_.jmcg_periodperrununit);
            var periodAmount = Target.GetInt(Fields.jmcg_workflowtask_.jmcg_periodperrunamount);
            var skipWeekendsAndClosures = Target.GetBoolean(Fields.jmcg_workflowtask_.jmcg_skipweekendsandbusinessclosures);
            var nextExecutionTime = CalculateNextExecutionTime(thisExecutionTime, periodUnit, periodAmount, skipWeekendsAndClosures);
            return nextExecutionTime;
        }

        private Entity _view;
        private Entity View
        {
            get
            {
                if (_view == null)
                {
                    var savedQueryId = Target.GetStringField(Fields.jmcg_workflowtask_.jmcg_targetviewid);
                    _view = XrmService.Retrieve(Entities.savedquery, new Guid(savedQueryId));
                }
                return _view;
            }
        }

        public QueryExpression GetViewFetchAsQuery()
        {
            var fetchXml = View.GetStringField(Fields.savedquery_.fetchxml);
            var query = XrmService.ConvertFetchToQueryExpression(fetchXml);
            return query;
        }

        private void AppendAliasTypeMaps(LinkEntity linkEntity, Dictionary<string, string> aliasTypeMaps)
        {
            if (linkEntity != null)
            {
                if(!string.IsNullOrWhiteSpace(linkEntity.EntityAlias) && !aliasTypeMaps.ContainsKey(linkEntity.EntityAlias))
                {
                    aliasTypeMaps.Add(linkEntity.EntityAlias, linkEntity.LinkToEntityName);
                }
                if (linkEntity.LinkEntities != null)
                {
                    foreach(var link in linkEntity.LinkEntities)
                    {
                        AppendAliasTypeMaps(link, aliasTypeMaps);
                    }
                }
            }
        }

        private void SendViewNotifications(Entity target)
        {
            var query = GetViewFetchAsQuery();
            var aliasTypeMaps = new Dictionary<string, string>();
            if(query.LinkEntities != null)
            {
                foreach (var link in query.LinkEntities)
                {
                    AppendAliasTypeMaps(link, aliasTypeMaps);
                }
            }

            var sendOption = target.GetOptionSetValue(Fields.jmcg_workflowtask_.jmcg_viewnotificationoption);
            switch(sendOption)
            {
                case OptionSets.WorkflowTask.ViewNotificationOption.EmailOwningUsers:
                    {
                        //add owner to query
                        query.ColumnSet.AddColumn("ownerid");
                        //remove any current user woner filters
                        if (query.Criteria != null && query.Criteria.Conditions != null)
                        {
                            foreach (var condition in query.Criteria.Conditions)
                            {
                                if (condition.Operator == ConditionOperator.EqualUserId)
                                    condition.Operator = ConditionOperator.NotNull;
                            }
                        }
                        //ensure owner has email
                        var ownerLink = query.AddLink(Entities.systemuser, "ownerid", Fields.systemuser_.systemuserid);
                        ownerLink.LinkCriteria.AddCondition(new ConditionExpression(Fields.systemuser_.internalemailaddress, ConditionOperator.NotNull));
                        var recordsForReminder = XrmService.RetrieveAll(query); //get all as results for multiple owners
                        var usersToNotifyIds = new List<Guid>();
                        foreach(var item in recordsForReminder)
                        {
                            if(item.GetLookupType("ownerid") == Entities.systemuser)
                            {
                                var userId = item.GetLookupGuid("ownerid");
                                if (userId.HasValue)
                                    usersToNotifyIds.Add(userId.Value);
                            }
                        }
                        usersToNotifyIds = usersToNotifyIds.Distinct().ToList();
                        var userTimeZones = IndexUserTimeZones(usersToNotifyIds);

                        foreach (var userId in usersToNotifyIds)
                        {
                            var localisationService = userTimeZones.ContainsKey(userId)
                                ? new LocalisationService(new LocalisationSettings(userTimeZones[userId]))
                                : LocalisationService;

                            var thisDudesRecords = recordsForReminder
                                .Where(r => r.GetLookupGuid("ownerid") == userId)
                                .ToArray();
                            SendViewNotificationEmailWithTable(Entities.systemuser, userId, thisDudesRecords, localisationService, aliasTypeMaps);
                        }
                        break;
                    }
                case OptionSets.WorkflowTask.ViewNotificationOption.EmailQueue:
                    {
                        var recordsForReminder = XrmService.RetrieveFirstX(query, HtmlEmailGenerator.MaximumNumberOfEntitiesToList + 1); //add 1 to determine if exceeded the limit
                        var recipientQueue = Target.GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_viewnotificationqueue);
                        if(!recipientQueue.HasValue)
                            throw new NullReferenceException(string.Format("Error required field {0} is empty on the target {1}", XrmService.GetFieldLabel(Fields.jmcg_workflowtask_.jmcg_viewnotificationqueue, TargetType), XrmService.GetEntityLabel(TargetType)));
                        SendViewNotificationEmailWithTable(Entities.queue, recipientQueue.Value, recordsForReminder, LocalisationService, aliasTypeMaps);
                        break;
                    }
            }
        }

        private IDictionary<Guid, string> IndexUserTimeZones(List<Guid> usersToNotifyIds)
        {
            var linkEntity = new LinkEntity(Entities.systemuser, Entities.usersettings, Fields.systemuser_.systemuserid, Fields.usersettings_.systemuserid, JoinOperator.Inner);
            var tzLink = linkEntity.AddLink(Entities.timezonedefinition, Fields.usersettings_.timezonecode, Fields.timezonedefinition_.timezonecode);
            tzLink.EntityAlias = "TZ";
            tzLink.Columns = XrmService.CreateColumnSet(new[] { Fields.timezonedefinition_.standardname });

            var users = XrmService.RetrieveAllOrClauses(Entities.systemuser
                , usersToNotifyIds.Select(uid => new ConditionExpression(Fields.systemuser_.systemuserid, ConditionOperator.Equal, uid))
                , new string[0]
                , linkEntity);

            return users
                .ToDictionary(u => u.Id, u => (string)u.GetFieldValue("TZ." + Fields.timezonedefinition_.standardname));
        }

        private void RemoveAllFields(LinkEntity link)
        {
            link.Columns = new ColumnSet(false);
            if (link.LinkEntities != null)
            {
                foreach (var childLink in link.LinkEntities)
                {
                    RemoveAllFields(link);
                }
            }
        }

        public void SendViewNotificationEmailWithTable(string recipientType, Guid recipientId, IEnumerable<Entity> recordsToList, LocalisationService localisationService, IDictionary<string, string> aliasTypeMaps)
        {
            if (recordsToList.Any())
            {
                var crmUrl = GetCrmURL();
                var isToOwner = Target.GetOptionSetValue(Fields.jmcg_workflowtask_.jmcg_viewnotificationoption) == OptionSets.WorkflowTask.ViewNotificationOption.EmailOwningUsers;
                //var entityType = GetViewFetchAsQuery().EntityName;
                //var viewHyperlink = string.Format("<a href={0}>{0}</a>", baseUrl);
                //var pStyle = "style='font-family: Arial,sans-serif;font-size: 12pt;padding:6.15pt 6.15pt 6.15pt 6.15pt'";
                //var content =
                //    string.Format(@"<p {0}>This is an automated notification there are '{1}' to be actioned</p>
                //                <p {0}>Please review and process the records</p>
                //                <p {0}>{2}</p>", pStyle, View.GetStringField(Fields.savedquery_.name), viewHyperlink);
                //exlude primary key and fields in linked entities in list because label
                var fieldsForTable = GetViewLayoutcellFieldNames()
                    .Except(new[] { XrmService.GetPrimaryKeyField(recordsToList.First().LogicalName) })
                    .ToList();
                if(isToOwner && fieldsForTable.Contains("ownerid"))
                {
                    fieldsForTable.Remove("ownerid");
                }

                var appId = recipientType == Entities.systemuser
                    ? WorkflowSchedulerService.GetUserAppId(recipientId, Target.GetStringField(Fields.jmcg_workflowtask_.jmcg_fieldforteamappid))
                    : WorkflowSchedulerService.GetQueueAppId(recipientId, Target.GetStringField(Fields.jmcg_workflowtask_.jmcg_fieldforteamappid));
                var email = new HtmlEmailGenerator(XrmService, crmUrl, appId);
                email.AppendParagraph(string.Format("This is an automated notification {0} {1}"
                    , isToOwner ? "that you own" : "there are"
                    , View.GetStringField(Fields.savedquery_.name)));
                email.AppendTable(recordsToList, localisationService, fields: fieldsForTable, aliasTypeMaps: aliasTypeMaps);
                var viewName = View.GetStringField(Fields.savedquery_.name);
                var subject = viewName + " Notification";
                SendNotificationEmail(recipientType, recipientId, subject, email.GetContent());
            }
        }

        public IEnumerable<string> GetViewLayoutcellFieldNames()
        {
            var layoutXml = View.GetStringField(Fields.savedquery_.layoutxml);
            var xml = "<xml>" + layoutXml + "</xml>";
            var attributeNames = new List<string>();
            var xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(xml);
            if (xmlDocument.ChildNodes != null
                && xmlDocument.ChildNodes.Count > 0
                && xmlDocument.ChildNodes[0].ChildNodes != null
                && xmlDocument.ChildNodes[0].ChildNodes.Count > 0
                && xmlDocument.ChildNodes[0].ChildNodes[0].ChildNodes != null
                && xmlDocument.ChildNodes[0].ChildNodes[0].ChildNodes.Count > 0)
            {
                var attributeNodes = xmlDocument
                    .ChildNodes[0] //xml
                    .ChildNodes[0] //grid
                    .ChildNodes[0] //row
                    .ChildNodes; //cells
                foreach (XmlNode child in attributeNodes)
                {
                    attributeNames.Add(child.Attributes["name"].Value);
                }
            }
            return attributeNames;
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
    }
}
