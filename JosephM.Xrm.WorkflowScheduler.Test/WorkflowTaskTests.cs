using System;
using System.Collections.Generic;
using System.Linq;
using JosephM.Core.Extentions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Schema;

namespace JosephM.Xrm.WorkflowScheduler.Test
{
    [TestClass]
    public class WorkflowTaskTests : JosephMXrmTest
    {
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

        private Entity InitialiseValidWorkflowTask()
        {
            var targetName = TestAccountTargetWorkflowName;
            var targetWorkflowTaskWorkflow = GetWorkflow(targetName);
            var fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='account'>
                          </entity>
                        </fetch>";

            return InitialiseWorkflowTask(targetName, targetWorkflowTaskWorkflow, fetchXml);
        }

        private static string TestAccountTargetWorkflowName
        {
            get
            {
                var targetName = "Test Account Target";
                return targetName;
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
        public void WorkflowTaskFetchTargetTestTest()
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

        private static Entity InitialiseWorkflowTask(string name, Entity targetWorkflow, string fetchXml)
        {
            var entity = new Entity(Entities.jmcg_workflowtask);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_name, name);
            entity.SetLookupField(Fields.jmcg_workflowtask_.jmcg_targetworkflow, targetWorkflow);
            entity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype,
                OptionSets.WorkflowTask.WorkflowExecutionType.TargetThisWorkflowTask);
            entity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_periodperrununit,
                OptionSets.WorkflowTask.PeriodPerRunUnit.Days);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_periodperrunamount, 1);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, DateTime.UtcNow);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_on, true);
            if (!fetchXml.IsNullOrWhiteSpace())
            {
                entity.SetField(Fields.jmcg_workflowtask_.jmcg_fetchquery, fetchXml);
                entity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype,
                    OptionSets.WorkflowTask.WorkflowExecutionType.TargetPerFetchResult);
            }
            else
            {
                entity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype,
                    OptionSets.WorkflowTask.WorkflowExecutionType.TargetThisWorkflowTask);
            }
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

        private Entity GetWorkflow(string name)
        {
            var query = XrmService.BuildQuery(Entities.workflow, null, new[]
            {
                new ConditionExpression(Fields.workflow_.type, ConditionOperator.Equal, OptionSets.Process.Type.Definition),
                new ConditionExpression(Fields.workflow_.name, ConditionOperator.Equal, name) 
            }, null);
            var workflow = XrmService.RetrieveFirst(query);
            if (workflow == null)
                throw new NullReferenceException("Couldn't find workflow " + name);
            return workflow;
        }
    }
}