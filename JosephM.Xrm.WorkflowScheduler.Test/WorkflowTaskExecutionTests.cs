using JosephM.Xrm.WorkflowScheduler.Workflows;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Schema;
using System;
using System.Collections.Generic;
using System.Threading;

namespace JosephM.Xrm.WorkflowScheduler.Test
{
    [TestClass]
    public class WorkflowTaskExecutionTests : JosephMXrmTest
    {
        /// <summary>
        /// Verifies continuous workflow not started, started or killed for change of On value
        /// </summary>
        [TestMethod]
        public void WorkflowTaskExecutionGetAppIdForEmailLinksTests()
        {
            var team = GetTestTeam();
            Assert.IsNotNull(team.GetStringField(Fields.team_.jmcg_appid));
            var testTeamsQueue = team.GetLookupGuid(Fields.team_.queueid);
            Assert.IsTrue(testTeamsQueue.HasValue);

            var workflow = InitialiseValidWorkflowTask();

            var testGuid = Guid.NewGuid();
            workflow.SetField(Fields.jmcg_workflowtask_.jmcg_fieldforteamappid, testGuid.ToString());
            workflow = CreateAndRetrieve(workflow);

            var workflowInstance = CreateWorkflowInstance<WorkflowTaskExecutionInstance>(workflow);
            Assert.AreEqual(testGuid.ToString(), workflowInstance.GetAppIdForTarget(team.LogicalName, team.Id));

            var testSettings = GetTestSettings();
            Assert.IsNotNull(testSettings.GetStringField(Fields.jmcg_wstestsettings_.jmcg_jmcg_appid));

            workflow.SetField(Fields.jmcg_workflowtask_.jmcg_fieldforteamappid, Entities.jmcg_wstestsettings + "." + Fields.jmcg_wstestsettings_.jmcg_jmcg_appid);
            workflow = UpdateFieldsAndRetreive(workflow, Fields.jmcg_workflowtask_.jmcg_fieldforteamappid);

            workflowInstance = CreateWorkflowInstance<WorkflowTaskExecutionInstance>(workflow);
            Assert.AreEqual(testSettings.GetStringField(Fields.jmcg_wstestsettings_.jmcg_jmcg_appid), workflowInstance.GetAppIdForTarget(team.LogicalName, team.Id));



            workflow.SetField(Fields.jmcg_workflowtask_.jmcg_fieldforteamappid, Fields.team_.jmcg_appid);
            workflow = UpdateFieldsAndRetreive(workflow, Fields.jmcg_workflowtask_.jmcg_fieldforteamappid);

            workflowInstance = CreateWorkflowInstance<WorkflowTaskExecutionInstance>(workflow);
            Assert.AreEqual(team.GetStringField(Fields.team_.jmcg_appid), workflowInstance.GetAppIdForTarget(Entities.queue, testTeamsQueue.Value));
        }

        /// <summary>
        /// Verifies continuous workflow not started, started or killed for change of On value
        /// </summary>
        [TestMethod]
        public void WorkflowTaskExecutionCalculateNextExecutionTimeTests()
        {
            Entity calendar = DeleteAllBusinessClosures();

            //first verify a tuesday holiday is skipped after a monday execution

            //move to a future monday
            var today = DateTime.Today.AddDays(2);
            var executionTime = new DateTime(today.Year, today.Month, today.Day, 0, 0, 0, DateTimeKind.Utc); ;
            while (executionTime.DayOfWeek != DayOfWeek.Monday)
                executionTime = executionTime.AddDays(1);

            //create a closure for the tuesday
            calendar = CreateBusinessClosure(calendar, executionTime.Date.AddDays(1));

            //verify skipped if skip and not if not
            var workflowInstance = CreateWorkflowInstance<WorkflowTaskExecutionInstance>();
            workflowInstance.CurrentUserId = CurrentUserId;
            Assert.AreEqual(executionTime.AddDays(2), workflowInstance.CalculateNextExecutionTime(executionTime, OptionSets.WorkflowTask.PeriodPerRunUnit.Days, 1, true, false, null, null));
            Assert.AreEqual(executionTime.AddDays(1), workflowInstance.CalculateNextExecutionTime(executionTime, OptionSets.WorkflowTask.PeriodPerRunUnit.Days, 1, false, false, null, null));

            //now lets verify that Saturday and Sunday are skipped

            //move to a future friday
            while (executionTime.DayOfWeek != DayOfWeek.Friday)
                executionTime = executionTime.AddDays(1);
            //verify skipped if skip and not if not
            Assert.AreEqual(executionTime.AddDays(3), workflowInstance.CalculateNextExecutionTime(executionTime, OptionSets.WorkflowTask.PeriodPerRunUnit.Days, 1, true, false, null, null));
            Assert.AreEqual(executionTime.AddDays(1), workflowInstance.CalculateNextExecutionTime(executionTime, OptionSets.WorkflowTask.PeriodPerRunUnit.Days, 1, false, false, null, null));

            //now lets verify monday skipped as well if a holiday
            //create the monday holiday
            calendar = CreateBusinessClosure(calendar, executionTime.AddDays(3));
            workflowInstance = CreateWorkflowInstance<WorkflowTaskExecutionInstance>();
            workflowInstance.CurrentUserId = CurrentUserId;
            //verify skipped if skip and not if not
            Assert.AreEqual(executionTime.AddDays(4), workflowInstance.CalculateNextExecutionTime(executionTime, OptionSets.WorkflowTask.PeriodPerRunUnit.Days, 1, true, false, null, null));
            Assert.AreEqual(executionTime.AddDays(1), workflowInstance.CalculateNextExecutionTime(executionTime, OptionSets.WorkflowTask.PeriodPerRunUnit.Days, 1, false, false, null, null));

            //okay lets check the limited business hours
            calendar = DeleteAllBusinessClosures();
            var timeOfDay = DateTime.Now.TimeOfDay;
            var tenMinutesAgo = timeOfDay.Add(new TimeSpan(0, 0, -10, 0, 0));
            var oneHourTenMinutesAgo = timeOfDay.Add(new TimeSpan(0, -1, -10, 0, 0));
            var expected = DateTime.Today.AddDays(1).Add(oneHourTenMinutesAgo).ToUniversalTime();
            Assert.AreEqual(expected, workflowInstance.CalculateNextExecutionTime(DateTime.UtcNow, OptionSets.WorkflowTask.PeriodPerRunUnit.Hours, 1, false, true, oneHourTenMinutesAgo, tenMinutesAgo));
        }

        /// <summary>
        /// Verifies continuous workflow not started, started or killed for change of On value
        /// </summary>
        [TestMethod]
        public void WorkflowTaskExecutionCalculateNextExecutionTimeRemovesSecondsTests()
        {
            var utcNow = DateTime.UtcNow;
            while(utcNow.Second == 0)
            {
                Thread.Sleep(1000);
                utcNow = DateTime.UtcNow;
            }

            var workflowInstance = CreateWorkflowInstance<WorkflowTaskExecutionInstance>();
            workflowInstance.CurrentUserId = CurrentUserId;

            var next = workflowInstance.CalculateNextExecutionTime(utcNow, OptionSets.WorkflowTask.PeriodPerRunUnit.Days, 1, true, false, null, null);

            var nextSecondsPart = next.Second;
            var nextMilliSecondsPart = next.Millisecond;
            Assert.AreEqual(0, nextSecondsPart);
            Assert.AreEqual(0, nextMilliSecondsPart);
        }

        private Entity CreateBusinessClosure(Entity calendar, DateTime holidayDate)
        {
            // Create a new calendar rule and assign the inner calendar id to it
            Entity calendarRule = new Entity(Entities.calendarrule);
            calendarRule.SetField(Fields.calendarrule_.duration, 1440); // 24hrs in minutes
            //It specifies the extent of the Calendar rule,generally an Integer value.
            calendarRule.SetField(Fields.calendarrule_.extentcode, 1);
            //Pattern of the rule recurrence. As we have given FREQ=DAILY it will create a calendar rule on daily basis. We can even create on Weekly basis by specifying FREQ=WEEKLY.
            //INTERVAL = 1; – This means how many days interval between the next same schedule.For e.g if the date was 6th April and interval was 2, it would create a schedule for 8th april, 10th april and so on…
            //COUNT = 1; This means how many recurring records should be created, if in the above example the count was given as 2, it would create schedule for 6th and 8th and then stop.If the count was 3, it would go on until 10th and then stop.
            calendarRule.SetField(Fields.calendarrule_.pattern, "FREQ=DAILY;INTERVAL=1;COUNT=1");
            //Rank is an Integer value which specifies the Rank value of the Calendar rule
            calendarRule.SetField(Fields.calendarrule_.rank, 0);

            // Timezone code to be set which the calendar rule will follow
            calendarRule.SetField(Fields.calendarrule_.timezonecode, -1);

            //Specifying the Calendar Id
            calendarRule.SetField(Fields.calendarrule_.issimple, false);
            calendarRule.SetLookupField(Fields.calendarrule_.calendarid, calendar.Id, Entities.calendar);

            //Start time for the created Calendar rule
            calendarRule.SetField(Fields.calendarrule_.starttime, holidayDate);
            calendarRule.SetField(Fields.calendarrule_.effectiveintervalend, holidayDate.AddDays(1));
            //Now we will add this rule to the earlier retrieved calendar rules
            calendarRule.SetField(Fields.calendarrule_.timecode, 2);
            calendarRule.SetField(Fields.calendarrule_.subcode, 5);
            calendarRule.SetField(Fields.calendarrule_.name, "TEST Closure");
            calendar.SetField("calendarrules", new EntityCollection(new List<Entity>(new[] { calendarRule })));
            calendar = UpdateFieldsAndRetreive(calendar, "calendarrules");
            return calendar;
        }

        private Entity DeleteAllBusinessClosures()
        {
            var organisation = XrmService.GetFirst(Entities.organization, new[] { Fields.organization_.businessclosurecalendarid });
            var businessClosureCalendarId = organisation.GetGuidField(Fields.organization_.businessclosurecalendarid);
            if (businessClosureCalendarId == Guid.Empty)
                throw new NullReferenceException(string.Format("Error {0} is empty in the {1} record", XrmService.GetFieldLabel(Fields.organization_.businessclosurecalendarid, Entities.organization), XrmService.GetEntityLabel(Entities.organization)));
            var calendar = XrmService.Retrieve(Entities.calendar, businessClosureCalendarId);
            var rules = calendar.GetEntitiesField("calendarrules");
            calendar.SetField("calendarrules", new EntityCollection());
            calendar = UpdateFieldsAndRetreive(calendar, "calendarrules");
            return calendar;
        }
    }
}