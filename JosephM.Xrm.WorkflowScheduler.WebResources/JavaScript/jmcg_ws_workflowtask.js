wsWorkflowTasks = new Object();
wsWorkflowTasks.Options = new Object();
wsWorkflowTasks.Options.WorkflowExecutionType = new Object();
wsWorkflowTasks.Options.WorkflowExecutionType.TargetThisWorkflowTask = 1;
wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerFetchResult = 2;

wsWorkflowTasks.RunOnLoad = function() {
    wsPageUtility.CommonForm(wsWorkflowTasks.RunOnChange, wsWorkflowTasks.RunOnSave);
    wsWorkflowTasks.RefreshVisibility();
}

wsWorkflowTasks.RunOnChange = function (fieldName) {
    switch (fieldName) {
        case "jmcg_workflowexecutiontype":
            wsWorkflowTasks.RefreshVisibility();
            break;
    }
}

wsWorkflowTasks.RunOnSave = function () {
}

wsWorkflowTasks.RefreshVisibility = function () {
    var isFetchTarget = wsPageUtility.GetFieldValue("jmcg_workflowexecutiontype") == wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerFetchResult;
    wsPageUtility.SetFieldVisibility("jmcg_fetchquery", isFetchTarget);
    wsPageUtility.SetFieldVisibility("jmcg_waitsecondspertargetworkflowcreation", isFetchTarget);
}