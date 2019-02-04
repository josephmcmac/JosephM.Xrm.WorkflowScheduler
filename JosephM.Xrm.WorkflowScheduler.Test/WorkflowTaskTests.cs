using JosephM.Xrm.WorkflowScheduler.Plugins;
using JosephM.Xrm.WorkflowScheduler.Workflows;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Schema;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace JosephM.Xrm.WorkflowScheduler.Test
{
    [TestClass]
    public class WorkflowTaskTests : JosephMXrmTest
    {
        /// <summary>
        ///validates creation of a monitor only workflow task
        ///and that it picks up and sends a notification for
        ///a failure of the taret workflow
        /// </summary>
        [TestMethod]
        public void WorkflowTaskMonitorOnlyTest()
        {
            var workflowName = "Test Account Target Failure";
            var taskName = "Test Monitor Only";

            DeleteWorkflowTasks(taskName);

            var workflowWillFail = GetWorkflow(workflowName);
            //create the workflow task montior only
            var workflowTask = new Entity(Entities.jmcg_workflowtask);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_name, taskName);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_notesfortargetfailureemail, "These are some fake instructions\nto fix the problem");
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_targetworkflow, workflowWillFail);
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype,
                OptionSets.WorkflowTask.WorkflowExecutionType.MonitorOnly);
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom, TestQueue);
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto, TestQueue);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_on, true);
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_crmbaseurl, "jmcg_wstestsettings.jmcg_crminstanceurl");
            workflowTask = CreateAndRetrieve(workflowTask);

            //wait a second and verify only the montior workflow spawned
            //i.e this doesn't have a continuous workflow as it doesn't
            //run a process on a schedule it just monitors
            Thread.Sleep(5000);
            Assert.IsFalse(WorkflowSchedulerService.GetRecurringInstances(workflowTask.Id).Any());
            Assert.IsTrue(WorkflowSchedulerService.GetMonitorInstances(workflowTask.Id).Any());
            //stop the monitor we will start it again in a second
            WorkflowSchedulerService.StopMonitorWorkflowFor(workflowTask.Id);

            //run the workflow which will fail
            var account = CreateAccount();
            XrmService.StartWorkflow(workflowWillFail.Id, account.Id);
            //wait a second and spawn the monitor
            Thread.Sleep(5000);
            WorkflowSchedulerService.StartNewMonitorWorkflowFor(workflowTask.Id);
            //verify the notification email created
            WaitTillTrue(() => GetRegardingEmails(workflowTask).Any(), 60);
            
            DeleteMyToday();
        }

        /// <summary>
        /// Vaslidatesselected queues have email addresses populated
        /// </summary>
        [TestMethod]
        public void WorkflowTaskValidateQueueEmailsPopulatedTest()
        {
            var workflowTask = InitialiseValidWorkflowTask();

            workflowTask.SetLookupField(WorkflowTaskPlugin.QueueFields.First(), TestQueueNoEmailAddress);
            try
            {
                XrmService.Create(workflowTask);
                Assert.Fail();
            }
            catch(Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflowTask.SetLookupField(WorkflowTaskPlugin.QueueFields.First(), TestQueue);
            workflowTask = XrmService.CreateAndRetreive(workflowTask);

            foreach (var queueField in WorkflowTaskPlugin.QueueFields)
            {
                workflowTask.SetLookupField(queueField, TestQueueNoEmailAddress);
                try
                {
                    UpdateFieldsAndRetreive(workflowTask, queueField);
                    Assert.Fail();
                }
                catch (Exception ex)
                {
                    Assert.IsFalse(ex is AssertFailedException);
                }
                workflowTask.SetLookupField(queueField, TestQueue);
                workflowTask = UpdateFieldsAndRetreive(workflowTask, queueField);
            }
        }

        /// <summary>
        /// Vaslidates erros thrown when the jmcg_crmbaseurl field does not have a valid value
        /// </summary>
        [TestMethod]
        public void WorkflowTaskValidateUrlTest()
        {
            var workflowTask = InitialiseValidWorkflowTask();
            var validUrlSettingsFieldValue = workflowTask.GetStringField(Fields.jmcg_workflowtask_.jmcg_crmbaseurl);
            Assert.IsNotNull(validUrlSettingsFieldValue);
            var split = validUrlSettingsFieldValue.Split('.');
            Assert.AreEqual(2, split.Count());
            Assert.IsTrue(XrmService.IsString(split.ElementAt(1), split.First()));

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_crmbaseurl, "BlahBlah");
            try
            {
                workflowTask = CreateAndRetrieve(workflowTask);
                Assert.Fail();
            }
            catch(Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_crmbaseurl, "https://url");
            workflowTask = CreateAndRetrieve(workflowTask);

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_crmbaseurl, "NotAnEntity.NotAField");
            try
            {
                workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_crmbaseurl);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_crmbaseurl, Entities.account + ".NotAField");
            try
            {
                workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_crmbaseurl);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_crmbaseurl, Entities.account + "." + Fields.account_.createdon);
            try
            {
                workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_crmbaseurl);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_crmbaseurl, validUrlSettingsFieldValue);
            workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_crmbaseurl);
        }

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
            //queues required when sending
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom, null);
            workflowTask.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto, null);
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
            workflow = CreateWorkflowInstance<TargetWorkflowTaskMonitorInstance>(workflowTask);

            Assert.IsFalse(workflow.HasNewFailure());

            Thread.Sleep(1000);

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, DateTime.UtcNow.AddDays(-2));
            workflowTask = UpdateFieldsAndRetreive(workflowTask, Fields.jmcg_workflowtask_.jmcg_nextexecutiontime);
            workflow = CreateWorkflowInstance<TargetWorkflowTaskMonitorInstance>(workflowTask);

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

        /// <summary>
        /// Verifies error thrown when below minimum period threshold
        /// </summary>
        [TestMethod]
        public void WorkflowTaskVerifyActivePeriod()
        {
            var workflow = InitialiseValidWorkflowTask();
            workflow.SetField(Fields.jmcg_workflowtask_.jmcg_onlyrunbetweenhours, true);
            try
            {
                workflow = CreateAndRetrieve(workflow);
                Assert.Fail();
            }
            catch(Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflow.SetField(Fields.jmcg_workflowtask_.jmcg_onlyrunbetweenhours, false);
            workflow = CreateAndRetrieve(workflow);

            workflow.SetField(Fields.jmcg_workflowtask_.jmcg_onlyrunbetweenhours, true);
            try
            {
                workflow = UpdateFieldsAndRetreive(workflow, Fields.jmcg_workflowtask_.jmcg_onlyrunbetweenhours);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflow.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_starthour, 5);
            workflow.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_startminute, 5);
            workflow.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_startampm, OptionSets.WorkflowTask.StartAMPM.PM);
            workflow.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_endhour, 4);
            workflow.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_endminute, 4);
            workflow.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_endampm, OptionSets.WorkflowTask.EndAMPM.PM);
            try
            {
                workflow = UpdateFieldsAndRetreive(workflow, Fields.jmcg_workflowtask_.jmcg_onlyrunbetweenhours
                    , Fields.jmcg_workflowtask_.jmcg_starthour
                    , Fields.jmcg_workflowtask_.jmcg_startminute
                    , Fields.jmcg_workflowtask_.jmcg_startampm
                    , Fields.jmcg_workflowtask_.jmcg_endhour
                    , Fields.jmcg_workflowtask_.jmcg_endminute
                    , Fields.jmcg_workflowtask_.jmcg_endampm
                    );
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflow.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_startampm, OptionSets.WorkflowTask.StartAMPM.AM);
            workflow = UpdateFieldsAndRetreive(workflow, Fields.jmcg_workflowtask_.jmcg_onlyrunbetweenhours
                    , Fields.jmcg_workflowtask_.jmcg_starthour
                    , Fields.jmcg_workflowtask_.jmcg_startminute
                    , Fields.jmcg_workflowtask_.jmcg_startampm
                    , Fields.jmcg_workflowtask_.jmcg_endhour
                    , Fields.jmcg_workflowtask_.jmcg_endminute
                    , Fields.jmcg_workflowtask_.jmcg_endampm
                    );
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

        [TestMethod]
        public void WorkflowTaskViewNotificationsTest()
        {
            DeleteAll(Entities.jmcg_workflowtask);

            //delete all accounts
            var accounts = XrmService.RetrieveAllEntityType(Entities.account);
            foreach (var account in accounts)
                XrmService.Delete(account);
            
            var account1 = CreateAccount();
            var account2 = CreateAccount();
            var account3 = CreateAccount();
            XrmService.Assign(account3, OtherUserId);
            var account4 = CreateAccount();
            XrmService.Assign(account4, GetTestTeam().Id, Entities.team);

            //field set when notification sent
            Assert.IsNull(account1.GetField(Fields.account_.lastusedincampaign));
            Assert.IsNull(account2.GetField(Fields.account_.lastusedincampaign));
            Assert.IsNull(account3.GetField(Fields.account_.lastusedincampaign));
            Assert.IsNull(account4.GetField(Fields.account_.lastusedincampaign));
            Assert.IsFalse(account1.GetBoolean(Fields.account_.donotfax));
            Assert.IsFalse(account2.GetBoolean(Fields.account_.donotfax));
            Assert.IsFalse(account3.GetBoolean(Fields.account_.donotfax));
            Assert.IsFalse(account4.GetBoolean(Fields.account_.donotfax));

            var testViewName = "Active Accounts";
            var testView = XrmService.GetFirst(Entities.savedquery, Fields.savedquery_.name, testViewName);
            var workflowActivity = InitialiseValidWorkflowTask();
            workflowActivity.SetLookupField(Fields.jmcg_workflowtask_.jmcg_targetworkflow, null);
            workflowActivity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype, OptionSets.WorkflowTask.WorkflowExecutionType.ViewNotification);
            //email sender required
            workflowActivity.SetField(Fields.jmcg_workflowtask_.jmcg_viewnotificationentitytype, null);
            try
            {
                workflowActivity = CreateAndRetrieve(workflowActivity);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflowActivity.SetField(Fields.jmcg_workflowtask_.jmcg_viewnotificationentitytype, Entities.account);
            //email sender required
            workflowActivity.SetField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom, null);
            try
            {
                workflowActivity = CreateAndRetrieve(workflowActivity);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflowActivity.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom, TestQueue);
            //email notification option required
            workflowActivity.SetField(Fields.jmcg_workflowtask_.jmcg_viewnotificationoption, null);
            try
            {
                workflowActivity = CreateAndRetrieve(workflowActivity);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflowActivity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_viewnotificationoption, OptionSets.WorkflowTask.ViewNotificationOption.EmailOwningUsers);

            //view required
            workflowActivity.SetField(Fields.jmcg_workflowtask_.jmcg_targetviewid, null);
            try
            {
                workflowActivity = CreateAndRetrieve(workflowActivity);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflowActivity.SetField(Fields.jmcg_workflowtask_.jmcg_targetviewid, testView.Id.ToString());
            workflowActivity.SetField(Fields.jmcg_workflowtask_.jmcg_setfieldwhennotificationsent, Fields.account_.lastusedincampaign);
            workflowActivity = CreateAndRetrieve(workflowActivity);
            Assert.AreEqual(testViewName, workflowActivity.GetStringField(Fields.jmcg_workflowtask_.jmcg_targetviewselectedname));
            //var instance = CreateWorkflowInstance<WorkflowTaskExecutionInstance>(workflowActivity);
            //instance.DoIt(true);

            WaitTillTrue(() => GetRegardingEmails(workflowActivity).Count() == 2, 60);
            var emails = GetRegardingEmails(workflowActivity);
            Assert.AreEqual(1, emails.Count(e => e.GetActivityPartyReferences(Fields.email_.to).First().Id == CurrentUserId));
            Assert.AreEqual(1, emails.Count(e => e.GetActivityPartyReferences(Fields.email_.to).First().Id == OtherUserId));

            //check field set when notification sent
            WaitTillTrue(() => Refresh(account3).GetField(Fields.account_.lastusedincampaign) != null, 60);
            account1 = Refresh(account1);
            account2 = Refresh(account2);
            account3 = Refresh(account3);
            account4 = Refresh(account4);
            Assert.IsNotNull(account1.GetField(Fields.account_.lastusedincampaign));
            Assert.IsNotNull(account2.GetField(Fields.account_.lastusedincampaign));
            Assert.IsNotNull(account3.GetField(Fields.account_.lastusedincampaign));
            //this one wasnt sent since linked to team with no email
            Assert.IsNull(account4.GetField(Fields.account_.lastusedincampaign));

            workflowActivity = InitialiseValidWorkflowTask();
            workflowActivity.SetLookupField(Fields.jmcg_workflowtask_.jmcg_targetworkflow, null);
            workflowActivity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype, OptionSets.WorkflowTask.WorkflowExecutionType.ViewNotification);
            workflowActivity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_viewnotificationoption, OptionSets.WorkflowTask.ViewNotificationOption.EmailQueue);
            workflowActivity.SetField(Fields.jmcg_workflowtask_.jmcg_targetviewid, testView.Id.ToString());
            workflowActivity.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom, TestQueue);
            workflowActivity.SetField(Fields.jmcg_workflowtask_.jmcg_viewnotificationentitytype, Entities.account);
            //recipient queue required
            workflowActivity.SetLookupField(Fields.jmcg_workflowtask_.jmcg_viewnotificationqueue, null);
            try
            {
                workflowActivity = CreateAndRetrieve(workflowActivity);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.IsFalse(ex is AssertFailedException);
            }
            workflowActivity.SetLookupField(Fields.jmcg_workflowtask_.jmcg_viewnotificationqueue, TestQueue);
            workflowActivity.SetField(Fields.jmcg_workflowtask_.jmcg_setfieldwhennotificationsent, Fields.account_.donotfax);
            workflowActivity = CreateAndRetrieve(workflowActivity);
            Assert.AreEqual(testViewName, workflowActivity.GetStringField(Fields.jmcg_workflowtask_.jmcg_targetviewselectedname));
            //var instance = CreateWorkflowInstance<WorkflowTaskExecutionInstance>(workflowActivity);
            //instance.DoIt(true);

            WaitTillTrue(() => GetRegardingEmails(workflowActivity).Count() == 1, 60);
            var secondEmail = GetRegardingEmails(workflowActivity).First();
            Assert.AreEqual(TestQueue.Id, secondEmail.GetActivityPartyReferences(Fields.email_.to).First().Id);

            //check field set when notification sent
            WaitTillTrue(() => Refresh(account4).GetBoolean(Fields.account_.donotfax), 60);
            account1 = Refresh(account1);
            account2 = Refresh(account2);
            account3 = Refresh(account3);
            account4 = Refresh(account4);
            Assert.IsTrue(account1.GetBoolean(Fields.account_.donotfax));
            Assert.IsTrue(account2.GetBoolean(Fields.account_.donotfax));
            Assert.IsTrue(account3.GetBoolean(Fields.account_.donotfax));
            Assert.IsTrue(account4.GetBoolean(Fields.account_.donotfax));
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