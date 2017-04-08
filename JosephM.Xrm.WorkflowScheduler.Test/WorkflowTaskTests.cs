using System;
using System.Collections.Generic;
using System.Linq;
using JosephM.Core.Extentions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Schema;
using JosephM.Xrm.WorkflowScheduler.Workflows;
using System.Threading;

namespace JosephM.Xrm.WorkflowScheduler.Test
{
    [TestClass]
    public class WorkflowTaskTests : JosephMXrmTest
    {
        /// <summary>
        /// Verifies continuous workflow not started, started or killed for change of On value
        /// </summary>
        [TestMethod]
        public void WorkflowTaskTurnOffIfDeactivatedTest()
        {
            var workflowTask = InitialiseValidWorkflowTask();
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_on, true);
            workflowTask = CreateAndRetrieve(workflowTask);

            Assert.IsTrue(workflowTask.GetBoolean(Fields.jmcg_workflowtask_.jmcg_on));
            Assert.IsTrue(WorkflowSchedulerService.GetRecurringInstances(workflowTask.Id).Any());

            XrmService.Deactivate(workflowTask);
            workflowTask = Refresh(workflowTask);
            Assert.IsFalse(workflowTask.GetBoolean(Fields.jmcg_workflowtask_.jmcg_on));
            Assert.IsFalse(WorkflowSchedulerService.GetRecurringInstances(workflowTask.Id).Any());
        }

        /// <summary>
        /// Verifies continuous workflow not started, started or killed for change of On value
        /// </summary>
        [TestMethod]
        public void WorkflowTaskMonitorTurnOnTest()
        {
            var workflowTask = InitialiseValidWorkflowTask();
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_targetworkflow, GetTargetAccountWorkflow());
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_sendnotificationfortargetfailures, true);
            try
            {
                workflowTask = CreateAndRetrieve(workflowTask);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_sendnotificationfortargetfailures, false);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_sendnotificationforschedulefailures, true);
            try
            {
                workflowTask = CreateAndRetrieve(workflowTask);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }

            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom, TestQueue);
            try
            {
                workflowTask = CreateAndRetrieve(workflowTask);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_sendnotificationfortargetfailures, true);
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto, TestQueue);
            workflowTask = CreateAndRetrieve(workflowTask);
            Assert.IsNotNull(workflowTask.GetField(Fields.jmcg_workflowtask_.jmcg_minimumtargetfailuredatetime));
            Assert.IsNotNull(workflowTask.GetField(Fields.jmcg_workflowtask_.jmcg_minimumschedulefailuredatetime));

            //okay need to set conditions for the target having not executed
            WaitTillTrue(() => WorkflowSchedulerService.GetMonitorInstances(workflowTask.Id).Any(), 60);

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_on, false);
            workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_on);
            Assert.IsFalse(WorkflowSchedulerService.GetMonitorInstances(workflowTask.Id).Any());

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_on, true);
            workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_on);
            Assert.IsTrue(WorkflowSchedulerService.GetMonitorInstances(workflowTask.Id).Any());
        }

        /// <summary>
        /// Verifies a notification generated for schedule workflow failures
        /// </summary>
        [TestMethod]
        public void WorkflowTaskMonitorScheduleTest()
        {
            DeleteAll(Entities.account);

            var workflowName = "Test Account Target Schedule Failure";

            DeleteWorkflowTasks(workflowName);

            var account = CreateAccount();

            var scheduleFailWorkflow = GetWorkflow(workflowName);

            var initialThreshold = DateTime.UtcNow.AddDays(-3);

            var workflowTask = InitialiseValidWorkflowTask();
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_targetworkflow, scheduleFailWorkflow);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, DateTime.UtcNow.AddDays(1));
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_name, workflowName);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_sendnotificationfortargetfailures, true);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_sendnotificationforschedulefailures, true);
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom, TestQueue);
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto, TestQueue);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_minimumschedulefailuredatetime, initialThreshold);
            workflowTask = CreateAndRetrieve(workflowTask);

            //wait until the monitor completed its first check - will respawn another the check it
            Thread.Sleep(10000);

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, DateTime.UtcNow.AddDays(-2));
            workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_nextexecutiontime);

            WaitTillTrue(() => WorkflowSchedulerService.GetRecurringInstancesFailed(workflowTask.Id).Any(), 60);

            WorkflowSchedulerService.StartNewMonitorWorkflowFor(workflowTask.Id);
            WaitTillTrue(() => GetRegardingEmails(workflowTask).Count() == 1, 60);
            WaitTillTrue(() => initialThreshold < Refresh(workflowTask).GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_minimumschedulefailuredatetime), 60);

            workflowTask = Refresh(workflowTask);
            Thread.Sleep(1000);

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, DateTime.UtcNow.AddDays(-1));
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_minimumschedulefailuredatetime, DateTime.UtcNow.AddDays(-2));
            workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, Fields.jmcg_workflowtask_.jmcg_minimumschedulefailuredatetime);

            WorkflowSchedulerService.StartNewContinuousWorkflowFor(workflowTask.Id);
            WaitTillTrue(() => WorkflowSchedulerService.GetRecurringInstancesFailed(workflowTask.Id).Count() > 1, 60);

            WorkflowSchedulerService.StartNewMonitorWorkflowFor(workflowTask.Id);
            WaitTillTrue(() => GetRegardingEmails(workflowTask).Count() == 2, 60);

            WorkflowSchedulerService.StartNewMonitorWorkflowFor(workflowTask.Id);
            Thread.Sleep(60000);
            Assert.AreEqual(2, GetRegardingEmails(workflowTask).Count());

            XrmService.Delete(workflowTask);
            XrmService.Delete(account);
        }
        /// <summary>
        /// Verifies a notification generated for target workflow failures
        /// </summary>
        [TestMethod]
        public void WorkflowTaskMonitorTargetsTest()
        {
            var initialMinumumThreshold = DateTime.UtcNow.AddDays(-1);

            DeleteAll(Entities.account, true);
            var workflowName = "Test Account Target Failure";

            DeleteWorkflowTasks(workflowName);

            var scheduleFailWorkflow = GetWorkflow(workflowName);
            DeleteSystemJobs(scheduleFailWorkflow);

            var account = CreateAccount();

            var workflowTask = InitialiseValidWorkflowTask();
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_targetworkflow, scheduleFailWorkflow);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, DateTime.UtcNow.AddDays(1));
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_name, workflowName);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_sendnotificationfortargetfailures, true);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_sendnotificationforschedulefailures, true);
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom, TestQueue);
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto, TestQueue);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_minimumtargetfailuredatetime, initialMinumumThreshold);
            workflowTask = CreateAndRetrieve(workflowTask);

            //wait until the monitor completed its first check - will respawn another the check it
            Thread.Sleep(10000);

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, DateTime.UtcNow.AddDays(-2));
            workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_nextexecutiontime);

            WaitTillTrue(() => Refresh(workflowTask).GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime) > DateTime.UtcNow.AddDays(-1), 60);

            var workflow = CreateWorkflowInstance<TargetWorkflowTaskMonitorInstance>(workflowTask);
            WaitTillTrue(() => workflow.HasNewFailure(), 60);

            WorkflowSchedulerService.StartNewMonitorWorkflowFor(workflowTask.Id);

            //okay need to set conditions for the target having not executed
            WaitTillTrue(() => GetRegardingEmails(workflowTask).Any(), 60);

            WaitTillTrue(() => Refresh(workflowTask).GetDateTimeField(Fields.jmcg_workflowtask_.jmcg_minimumtargetfailuredatetime) > initialMinumumThreshold, 60);
            workflowTask = Refresh(workflowTask);

            Assert.IsFalse(workflow.HasNewFailure());

            Thread.Sleep(1000);

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, DateTime.UtcNow.AddDays(-2));
            workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_nextexecutiontime);


            WaitTillTrue(() => workflow.HasNewFailure(), 60);

            XrmService.Delete(workflowTask);
            XrmService.Delete(account);
        }

        /// <summary>
        /// Verifies tasks executions wait between spawns and stop before maximum defined (isolation) timespan
        /// </summary>
        [TestMethod]
        public void WorkflowTaskVerifyWaitAndIsolationThresholdTests()
        {
            //get target workflow creates a note on the workflow task
            var targetName = TestAccountTargetWorkflowName;
            var targetWorkflow = GetTargetAccountWorkflow();
            //delete previous instances
            DeleteWorkflowTasks(targetName);
            //fetch all accounts
            var fetchXml = GetFetchAllAccounts();

            DeleteAll(Entities.account);
            //create some target accounts
            var toCreate = 5;
            var createdAccounts = new List<Entity>();
            for (var i = 0; i < toCreate; i++)
                createdAccounts.Add(CreateAccount());
            //create workflow task
            var waitTime = 3;
            var expectedMinimumTime = (toCreate * waitTime) - waitTime;

            var workflow = CreateWorkflowTask(targetName, targetWorkflow, fetchXml, waitPeriod: waitTime, nextExecution: DateTime.UtcNow.AddDays(2));
            //verify spawned and attached notes to the accounts
            var activity = CreateWorkflowInstance<WorkflowTaskExecutionInstance>(workflow);

            var now = DateTime.UtcNow;
            activity.DoIt(true);
            var finished = DateTime.UtcNow;
            Assert.IsTrue(finished - now >= new TimeSpan(0, 0, expectedMinimumTime));

            var fakeSandboxThreshold = 11;
            Assert.IsTrue(expectedMinimumTime > fakeSandboxThreshold);
            activity.MaxSandboxIsolationExecutionSeconds = fakeSandboxThreshold;
            now = DateTime.UtcNow;
            activity.DoIt(true);
            finished = DateTime.UtcNow;
            Assert.IsTrue(finished - now < new TimeSpan(0, 0, fakeSandboxThreshold));
        }

        /// <summary>
        /// Verifies continuous workflow not started, started or killed for change of On value
        /// </summary>
        [TestMethod]
        public void WorkflowTaskOnOffTest()
        {
            var workflowTask = InitialiseValidWorkflowTask();
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_on, false);
            workflowTask = CreateAndRetrieve(workflowTask);
            Assert.IsFalse(WorkflowSchedulerService.GetRecurringInstances(workflowTask.Id).Any());
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_on, true);
            workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_on);
            Assert.IsTrue(WorkflowSchedulerService.GetRecurringInstances(workflowTask.Id).Any());
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_on, false);
            workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_on);
            WaitTillTrue(() => !WorkflowSchedulerService.GetRecurringInstances(workflowTask.Id).Any(), 60);
        }

        /// <summary>
        /// Verifies error thrown when below minimum period threshold
        /// </summary>
        [TestMethod]
        public void WorkflowTaskVerifyTarget()
        {
            var workflowTask = InitialiseValidWorkflowTask();
            //invaldi or empty fetch throws error
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype, OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerFetchResult);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_fetchquery, null);
            VerifyCreateOrUpdateError(workflowTask);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_fetchquery, "12345");
            VerifyCreateOrUpdateError(workflowTask);

            //not targeting correct type invalid
            var targetWorkflow = GetTargetAccountWorkflow();
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype, OptionSets.WorkflowTask.WorkflowExecutionType.TargetThisWorkflowTask);
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_targetworkflow, targetWorkflow);
            VerifyCreateOrUpdateError(workflowTask);
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype, OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerFetchResult);
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_targetworkflow, targetWorkflow);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_fetchquery, GetFetchAll(Entities.workflow));
            VerifyCreateOrUpdateError(workflowTask);

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_fetchquery, GetFetchAll(Entities.account));
            workflowTask = CreateAndRetrieve(workflowTask);
            //verify triggered update as well
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_fetchquery, GetFetchAll(Entities.workflow));
            VerifyCreateOrUpdateError(workflowTask);
        }

        /// <summary>
        /// Verifies error thrown when below minimum period threshold
        /// </summary>
        [TestMethod]
        public void WorkflowTaskVerifyPeriod()
        {
            var workflow = InitialiseValidWorkflowTask();
            //0 interval is invalid
            workflow.SetField(Fields.jmcg_workflowtask_.jmcg_periodperrunamount, 0);
            VerifyCreateOrUpdateError(workflow);
            //1 minute is invalid
            workflow.SetField(Fields.jmcg_workflowtask_.jmcg_periodperrunamount, 1);
            workflow.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_periodperrununit, OptionSets.WorkflowTask.PeriodPerRunUnit.Minutes);
            VerifyCreateOrUpdateError(workflow);
            //create it valid
            workflow.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_periodperrununit, OptionSets.WorkflowTask.PeriodPerRunUnit.Days);
            workflow = CreateAndRetrieve(workflow);
            //verify also triggered update
            //1 minute is invalid
            workflow.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_periodperrununit, OptionSets.WorkflowTask.PeriodPerRunUnit.Minutes);
            VerifyCreateOrUpdateError(workflow);
            //valid
            workflow.SetField(Fields.jmcg_workflowtask_.jmcg_periodperrunamount, WorkflowSchedulerService.MinimumExecutionMinutes);
            workflow = UpdateFieldsAndRetreive(workflow);
        }

        private void VerifyCreateOrUpdateError(Entity workflow)
        {
            try
            {
                if (workflow.Id == Guid.Empty)
                    XrmService.Create(workflow);
                else
                    XrmService.Update(workflow);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
        }

        /// <summary>
        /// Verifies that a workflow task with workflow targeting itself correctly spawns and executes on creation
        /// </summary>
        [TestMethod]
        public void WorkflowTaskTargetTest()
        {
            //get target workflow creates a note on the workflow task
            var targetName = "Test Workflow Task Target";
            var targetWorkflowTaskWorkflow = GetWorkflow(targetName);
            //delete previous instance
            DeleteWorkflowTasks(targetName);
            //create new workflow task
            var entity = CreateWorkflowTask(targetName, targetWorkflowTaskWorkflow);
            //verify spawned and created task
            WaitTillTrue(() => GetRegardingNotes(entity).Any(), 60);
        }

        /// <summary>
        /// Verifies that a workflow task with workflow targeting fetch query correctly spawns and executes on creation
        /// </summary>
        [TestMethod]
        public void WorkflowTaskFetchTargetTest()
        {
            //get target workflow creates a note on the workflow task
            var targetName = TestAccountTargetWorkflowName;
            var targetWorkflow = GetTargetAccountWorkflow();
            //delete previous instances
            DeleteWorkflowTasks(targetName);
            //fetch all accounts
            var fetchXml = GetFetchAllAccounts();

            //delete all accounts
            var accounts = XrmService.RetrieveAllEntityType(Entities.account);
            foreach (var account in accounts)
                XrmService.Delete(account);
            //create some target accounts
            var toCreate = 3;
            var createdAccounts = new List<Entity>();
            for (var i = 0; i < toCreate; i++)
                createdAccounts.Add(CreateAccount());
            //create workflow task
            CreateWorkflowTask(targetName, targetWorkflow, fetchXml);
            //verify spawned and attached notes to the accounts
            WaitTillTrue(() => createdAccounts.All(e => GetRegardingNotes(e).Any()), 60);
        }

        [TestMethod]
        public void WorkflowTaskViewTargetTest()
        {
            //delete all accounts
            var accounts = XrmService.RetrieveAllEntityType(Entities.account);
            foreach (var account in accounts)
                XrmService.Delete(account);

            var testAccount = GetTESTACCOUNTVIEWAccount();

            var testViewName = "Test Custom Account System View";
            var testView = XrmService.GetFirst(Entities.savedquery, Fields.savedquery_.name, testViewName);
            var workflowActivity = InitialiseValidWorkflowTask();
            Assert.AreEqual(GetTargetAccountWorkflow().Id, workflowActivity.GetLookupGuid(Fields.jmcg_workflowtask_.jmcg_targetworkflow));
            workflowActivity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype, OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerViewResult);
            workflowActivity.SetField(Fields.jmcg_workflowtask_.jmcg_targetviewid, testView.Id.ToString());
            workflowActivity = CreateAndRetrieve(workflowActivity);
            Assert.AreEqual(testViewName, workflowActivity.GetStringField(Fields.jmcg_workflowtask_.jmcg_targetviewselectedname));
            //var instance = CreateWorkflowInstance<WorkflowTaskExecutionInstance>(workflowActivity);
            //instance.DoIt(true);

            WaitTillTrue(() => GetRegardingNotes(testAccount).Any(), 60);
        }

        private Entity GetTESTACCOUNTVIEWAccount()
        {
            var account = XrmService.GetFirst(Entities.account, Fields.account_.name, "TESTACCOUNTVIEW");
            if (account == null)
            {
                account = CreateTestRecord(Entities.account, new Dictionary<string, object>()
                {
                    { Fields.account_.name, "TESTACCOUNTVIEW" }
                });
            }
            return account;
        }

        private Entity GetTargetAccountWorkflow()
        {
            return GetWorkflow(TestAccountTargetWorkflowName);
        }

        private static string GetFetchAllAccounts()
        {
            var type = Entities.account;
            return GetFetchAll(type);
        }

        private static string GetFetchAll(string type)
        {
            var fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='" + type + @"'>
                          </entity>
                        </fetch>";
            return fetchXml;
        }

        private Entity CreateWorkflowTask(string name, Entity targetWorkflow, string fetchXml, int? waitPeriod = null, DateTime? nextExecution = null)
        {
            var entity = InitialiseWorkflowTask(name, targetWorkflow, fetchXml);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_waitsecondspertargetworkflowcreation, waitPeriod);
            if (nextExecution.HasValue)
                entity.SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, nextExecution);
            entity = CreateAndRetrieve(entity);
            return entity;
        }

        private Entity CreateWorkflowTask(string name, Entity targetWorkflow)
        {
            return CreateWorkflowTask(name, targetWorkflow, null);
        }

        private Entity CreateWorkflowTask(string name, Entity targetWorkflow, string fetchXml)
        {
            var entity = InitialiseWorkflowTask(name, targetWorkflow, fetchXml);
            entity = CreateAndRetrieve(entity);
            return entity;
        }

        private void DeleteWorkflowTasks(string targetName)
        {
            var entities = XrmService.RetrieveAllAndClauses(Entities.jmcg_workflowtask, new[]
            {
                new ConditionExpression(Fields.jmcg_workflowtask_.jmcg_name, ConditionOperator.Equal, targetName)
            });
            foreach (var item in entities)
                XrmService.Delete(item);
        }

        private void DeleteSystemJobs(Entity targetWorkflow)
        {
            var lastFailedJobQuery = XrmService.BuildQuery(Entities.asyncoperation,
                new string[0],
                null
                , null);
            var workflowLink = lastFailedJobQuery.AddLink(Entities.workflow, Fields.asyncoperation_.workflowactivationid,
                Fields.workflow_.workflowid);
            workflowLink.LinkCriteria.AddCondition(new ConditionExpression(Fields.workflow_.parentworkflowid, ConditionOperator.Equal, targetWorkflow.Id));
            var failedJobs = XrmService.RetrieveAll(lastFailedJobQuery);
            foreach (var item in failedJobs)
                XrmService.Delete(item);
        }
    }
}