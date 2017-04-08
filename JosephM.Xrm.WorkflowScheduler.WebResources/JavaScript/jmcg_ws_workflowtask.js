/// <reference path="jmcg_ws_pageutility.js" />
/// <reference path="jmcg_ws_webserviceutility.js" />

wsWorkflowTasks = new Object();
wsWorkflowTasks.Options = new Object();
wsWorkflowTasks.Options.WorkflowExecutionType = new Object();
wsWorkflowTasks.Options.WorkflowExecutionType.TargetThisWorkflowTask = 1;
wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerFetchResult = 2;
wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerViewResult = 3;

wsWorkflowTasks.RunOnLoad = function() {
    wsPageUtility.CommonForm(wsWorkflowTasks.RunOnChange, wsWorkflowTasks.RunOnSave);
    wsWorkflowTasks.RefreshVisibility();
    wsWorkflowTasks.PopulateViewSelectionList();
}

wsWorkflowTasks.RunOnChange = function (fieldName) {
    switch (fieldName) {
        case "jmcg_workflowexecutiontype":
            wsWorkflowTasks.RefreshVisibility();
            wsWorkflowTasks.PopulateViewSelectionList();
            break;
        case "jmcg_targetworkflow":
            wsWorkflowTasks.PopulateViewSelectionList();
            wsWorkflowTasks.RefreshVisibility();
            break;
        case "jmcg_targetviewselectionfield":
            wsWorkflowTasks.SetViewSelection();
            break;
    }
}

wsWorkflowTasks.RunOnSave = function () {
}

wsWorkflowTasks.RefreshVisibility = function () {
    var isFetchTarget = wsPageUtility.GetFieldValue("jmcg_workflowexecutiontype") == wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerFetchResult;
    var isViewTarget = wsPageUtility.GetFieldValue("jmcg_workflowexecutiontype") == wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerViewResult;
    wsPageUtility.SetFieldVisibility("jmcg_fetchquery", isFetchTarget);
    wsPageUtility.SetFieldVisibility("jmcg_waitsecondspertargetworkflowcreation", isFetchTarget || isViewTarget);
    var isTargetWorkflowSelected = wsPageUtility.GetFieldValue("jmcg_targetworkflow") != null;
    wsPageUtility.SetFieldVisibility("jmcg_targetviewselectionfield", isViewTarget && isTargetWorkflowSelected);
    wsPageUtility.SetFieldVisibility("jmcg_targetviewselectedname", isViewTarget && isTargetWorkflowSelected);
    wsPageUtility.SetFieldVisibility("jmcg_targetviewid", isViewTarget && isTargetWorkflowSelected);
}

wsWorkflowTasks.SetViewSelection = function () {
    var selectedoption = Xrm.Page.getAttribute("jmcg_targetviewselectionfield").getSelectedOption();
    if (selectedoption != null && parseInt(selectedoption.value) != 0) {
        var selectedViewId = null;
        var selectedViewName = null;
        var value = selectedoption.value;
        var selectedView = wsWorkflowTasks.ViewSelections[parseInt(value) - 1];
        selectedViewId = selectedView["savedqueryid"];
        selectedViewName = selectedView["name"];
        wsPageUtility.SetFieldValue("jmcg_targetviewid", selectedViewId);
        wsPageUtility.SetFieldValue("jmcg_targetviewselectedname", selectedViewName);
        wsPageUtility.SetFieldValue("jmcg_targetviewselectionfield", 0);
    }
}

wsWorkflowTasks.ClearSelectedView = function () {
    wsPageUtility.SetFieldValue("jmcg_targetviewselectionfield", null);
}

wsWorkflowTasks.ViewSelections = [];
wsWorkflowTasks.PopulateViewSelectionList = function () {
    var isViewTarget = wsPageUtility.GetFieldValue("jmcg_workflowexecutiontype") == wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerViewResult;
    if (isViewTarget) {
        var processTargetWorkflow = function (workflow) {
            var processViewResults = function (views) {
                wsWorkflowTasks.ViewSelections = views;
                var viewoptions = new Array();
                viewoptions.push(new wsPageUtility.PicklistOption(0, "Select to change the selected view"));
                var currentViewInOptions = false;
                for (var i = 1; i <= wsWorkflowTasks.ViewSelections.length; i++)
                {
                    viewoptions.push(new wsPageUtility.PicklistOption(i, wsWorkflowTasks.ViewSelections[i - 1]["name"]));
                    if (wsWorkflowTasks.ViewSelections[i - 1]["savedqueryid"] == wsPageUtility.GetFieldValue("jmcg_targetviewid"))
                        currentViewInOptions = true;
                }
                wsPageUtility.SetPicklistOptions("jmcg_targetviewselectionfield", viewoptions);
                wsPageUtility.SetFieldValue("jmcg_targetviewselectionfield", 0);
                if(!currentViewInOptions)
                {
                    wsPageUtility.SetFieldValue("jmcg_targetviewid", null);
                    wsPageUtility.SetFieldValue("jmcg_targetviewselectedname", null);
                }
            };

            var targetType = workflow["primaryentity"];
            var conditions = [new wsServiceUtility.FilterCondition("returnedtypecode", wsServiceUtility.FilterOperator.Equal, targetType)];
            var orders = [new wsServiceUtility.OrderCondition("name", false)];
            wsServiceUtility.RetrieveMultipleAsync("savedquery", ["savedqueryid", "name"], conditions, orders, processViewResults);
        };

        var targetWorkflowId = wsPageUtility.GetLookupId("jmcg_targetworkflow");
        if (targetWorkflowId != null && targetWorkflowId != "") {
            wsServiceUtility.RetrieveAsync("workflow", targetWorkflowId, ["primaryentity"], processTargetWorkflow);
        }
    }
    else {
        wsWorkflowTasks.ViewSelections = [];
        wsPageUtility.SetFieldValue("jmcg_targetviewid", null);
        wsPageUtility.SetFieldValue("jmcg_targetviewselectedname", null);
    }
}