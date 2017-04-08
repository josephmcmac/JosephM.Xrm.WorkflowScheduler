using System.Collections.Generic;
using JosephM.Xrm.WorkflowScheduler.Services;
using Microsoft.Xrm.Sdk;
using Schema;
using JosephM.Core.Extentions;
using System;
using Microsoft.Xrm.Sdk.Query;

namespace JosephM.Xrm.WorkflowScheduler.Test
{
    public abstract class JosephMXrmTest : XrmTest
    {
        protected override IEnumerable<string> EntitiesToDelete
        {
            get { return base.EntitiesToDelete; }
        }

        private WorkflowSchedulerService _workflowSchedulerService;
        public WorkflowSchedulerService WorkflowSchedulerService
        {
            get
            {
                if (_workflowSchedulerService == null)
                    _workflowSchedulerService = new WorkflowSchedulerService(XrmService);
                return _workflowSchedulerService;
            }
        }

        public static string TestAccountTargetWorkflowName
        {
            get
            {
                var targetName = "Test Account Target";
                return targetName;
            }
        }

        public Entity InitialiseValidWorkflowTask()
        {
            var targetName = TestAccountTargetWorkflowName;
            var targetWorkflowTaskWorkflow = GetWorkflow(targetName);
            var fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='account'>
                          </entity>
                        </fetch>";

            return InitialiseWorkflowTask(targetName, targetWorkflowTaskWorkflow, fetchXml);
        }

        public static Entity InitialiseWorkflowTask(string name, Entity targetWorkflow, string fetchXml)
        {
            var entity = new Entity(Entities.jmcg_workflowtask);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_name, name);
            entity.SetLookupField(Fields.jmcg_workflowtask_.jmcg_targetworkflow, targetWorkflow);
            entity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype,
                OptionSets.WorkflowTask.WorkflowExecutionType.TargetThisWorkflowTask);
            entity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_periodperrununit,
                OptionSets.WorkflowTask.PeriodPerRunUnit.Days);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_periodperrunamount, 1);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, DateTime.UtcNow.AddMinutes(-10));
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

        public Entity GetWorkflow(string name)
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