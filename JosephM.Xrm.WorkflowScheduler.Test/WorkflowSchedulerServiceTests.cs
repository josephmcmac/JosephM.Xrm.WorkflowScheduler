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
    public class WorkflowSchedulerServiceTests : JosephMXrmTest
    {
        /// <summary>
        ///validates creation of a monitor only workflow task
        ///and that it picks up and sends a notification for
        ///a failure of the taret workflow
        /// </summary>
        [TestMethod]
        public void WorkflowSchedulerServiceGetTimeSpanTests()
        {
            var workflowTask = InitialiseValidWorkflowTask();

            Assert.IsNull(WorkflowSchedulerService.GetStartTimeSpan(workflowTask.GetField));
            Assert.IsNull(WorkflowSchedulerService.GetEndTimeSpan(workflowTask.GetField));

            workflowTask.SetField(Fields.jmcg_workflowtask_.jmcg_onlyrunbetweenhours, true);

            Assert.IsNull(WorkflowSchedulerService.GetStartTimeSpan(workflowTask.GetField));
            Assert.IsNull(WorkflowSchedulerService.GetEndTimeSpan(workflowTask.GetField));

            //STARTS
            //1:01 AM
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_starthour, 1);
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_startminute, 1);
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_startampm, OptionSets.WorkflowTask.StartAMPM.AM);
            Assert.AreEqual(new TimeSpan(1, 1, 0), WorkflowSchedulerService.GetStartTimeSpan(workflowTask.GetField));

            //1:01 PM
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_startampm, OptionSets.WorkflowTask.StartAMPM.PM);
            Assert.AreEqual(new TimeSpan(13, 1, 0), WorkflowSchedulerService.GetStartTimeSpan(workflowTask.GetField));

            //12AM midnight
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_starthour, 12);
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_startminute, 0);
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_startampm, OptionSets.WorkflowTask.StartAMPM.AM);
            Assert.AreEqual(new TimeSpan(0, 0, 0), WorkflowSchedulerService.GetStartTimeSpan(workflowTask.GetField));

            //12PM noon
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_startampm, OptionSets.WorkflowTask.StartAMPM.PM);
            Assert.AreEqual(new TimeSpan(12, 0, 0), WorkflowSchedulerService.GetStartTimeSpan(workflowTask.GetField));

            //ENDS
            //1:01 AM
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_endhour, 1);
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_endminute, 1);
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_endampm, OptionSets.WorkflowTask.EndAMPM.AM);
            Assert.AreEqual(new TimeSpan(1, 1, 0), WorkflowSchedulerService.GetEndTimeSpan(workflowTask.GetField));

            //1:01 PM
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_endampm, OptionSets.WorkflowTask.EndAMPM.PM);
            Assert.AreEqual(new TimeSpan(13, 1, 0), WorkflowSchedulerService.GetEndTimeSpan(workflowTask.GetField));

            //12AM midnight
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_endhour, 12);
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_endminute, 0);
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_endampm, OptionSets.WorkflowTask.EndAMPM.AM);
            Assert.AreEqual(new TimeSpan(0, 0, 0), WorkflowSchedulerService.GetEndTimeSpan(workflowTask.GetField));

            //12PM noon
            workflowTask.SetOptionSetField(Fields.jmcg_workflowtask_.jmcg_endampm, OptionSets.WorkflowTask.EndAMPM.PM);
            Assert.AreEqual(new TimeSpan(12, 0, 0), WorkflowSchedulerService.GetEndTimeSpan(workflowTask.GetField));

        }
    }
}