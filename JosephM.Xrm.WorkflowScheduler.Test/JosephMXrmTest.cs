using System.Collections.Generic;
using JosephM.Xrm.WorkflowScheduler.Services;
using Microsoft.Xrm.Sdk;
using Schema;
using JosephM.Core.Extentions;
using System;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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

        public Entity InitialiseWorkflowTask(string name, Entity targetWorkflow, string fetchXml)
        {
            Assert.IsTrue(GetTestSettings() != null);

            var entity = new Entity(Entities.jmcg_workflowtask);
            entity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_periodperrununit,
                OptionSets.WorkflowTask.PeriodPerRunUnit.Days);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_name, name);
            entity.SetLookupField(Fields.jmcg_workflowtask_.jmcg_targetworkflow, targetWorkflow);
            entity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_workflowexecutiontype,
                OptionSets.WorkflowTask.WorkflowExecutionType.TargetThisWorkflowTask);
            entity.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_periodperrununit,
                OptionSets.WorkflowTask.PeriodPerRunUnit.Days);
            entity.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsfrom, TestQueue);
            entity.SetLookupField(Fields.jmcg_workflowtask_.jmcg_sendfailurenotificationsto, TestQueue);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_periodperrunamount, 1);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_nextexecutiontime, DateTime.UtcNow.AddMinutes(-10));
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_on, true);
            entity.SetField(Fields.jmcg_workflowtask_.jmcg_crmbaseurl, "jmcg_wstestsettings.jmcg_crminstanceurl");
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


        public Entity GetTestSettings()
        {
            var settings = XrmService.GetFirst("jmcg_wstestsettings");
            if (settings == null)
            {
                settings = new Entity("jmcg_wstestsettings");
                var connection = new XrmConnection(XrmConfiguration);
                settings.SetField("jmcg_name", "Test Workflow Settings");
                settings.SetField("jmcg_crminstanceurl", connection.GetWebUrl());
                settings = CreateAndRetrieve(settings);
            }
            return settings;
        }

        public Guid? _otherUserId;
        public Guid OtherUserId
        {
            get
            {
                if (!_otherUserId.HasValue)
                {
                    var userQuery = XrmService.BuildQuery(Entities.systemuser, null, new[]
                    {
                new ConditionExpression(Fields.systemuser_.systemuserid, ConditionOperator.NotEqual, CurrentUserId),
                new ConditionExpression(Fields.systemuser_.isintegrationuser, ConditionOperator.Equal, false),
                new ConditionExpression(Fields.systemuser_.isdisabled, ConditionOperator.Equal, false)
                }, null);
                    var otherUser = XrmService.RetrieveFirst(userQuery);
                    Assert.IsNotNull(otherUser);
                    _otherUserId = otherUser.Id;
                }
                return _otherUserId.Value;
            }
        }

        public Entity GetTestTeam()
        {
            var team = XrmService.GetFirst(Entities.team, Fields.team_.name, "TESTTEAM");
            if (team == null)
            {
                team = new Entity(Entities.team);
                team.SetField(Fields.team_.name, "TESTTEAM");
                team.SetLookupField(Fields.team_.businessunitid, XrmService.GetFirst(Entities.businessunit));
                team = CreateAndRetrieve(team);
            }
            return team;
        }
    }
}