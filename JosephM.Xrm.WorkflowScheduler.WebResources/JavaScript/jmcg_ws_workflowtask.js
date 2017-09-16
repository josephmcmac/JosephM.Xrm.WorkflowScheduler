/// <reference path="jmcg_ws_pageutility.js" />
/// <reference path="jmcg_ws_webserviceutility.js" />

wsWorkflowTasks = new Object();
wsWorkflowTasks.Options = new Object();
wsWorkflowTasks.Options.WorkflowExecutionType = new Object();
wsWorkflowTasks.Options.WorkflowExecutionType.TargetThisWorkflowTask = 1;
wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerFetchResult = 2;
wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerViewResult = 3;
wsWorkflowTasks.Options.WorkflowExecutionType.ViewNotification = 4;
wsWorkflowTasks.Options.ViewNotificationOption = new Object();
wsWorkflowTasks.Options.ViewNotificationOption.EmailQueue = 0;
wsWorkflowTasks.Options.ViewNotificationOption.EmailOwningUsers = 1;

wsWorkflowTasks.RunOnLoad = function() {
    wsPageUtility.CommonForm(wsWorkflowTasks.RunOnChange, wsWorkflowTasks.RunOnSave);
    wsWorkflowTasks.RefreshVisibility();
    wsWorkflowTasks.PopulateViewSelectionList();
    wsWorkflowTasks.PopulateUrlField();
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
        case "jmcg_viewnotificationentitytype":
            wsWorkflowTasks.PopulateViewSelectionList();
            wsWorkflowTasks.RefreshVisibility();
            break; 
        case "jmcg_viewnotificationoption":
            wsWorkflowTasks.RefreshVisibility();
            break;
    }
}

wsWorkflowTasks.RunOnSave = function () {
}

wsWorkflowTasks.RefreshVisibility = function () {
    var typeSelected = wsPageUtility.GetFieldValue("jmcg_workflowexecutiontype") != null;
    var isFetchTarget = wsPageUtility.GetFieldValue("jmcg_workflowexecutiontype") == wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerFetchResult;
    var isViewTarget = wsPageUtility.GetFieldValue("jmcg_workflowexecutiontype") == wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerViewResult;
    var isViewNotification = wsPageUtility.GetFieldValue("jmcg_workflowexecutiontype") == wsWorkflowTasks.Options.WorkflowExecutionType.ViewNotification;
    var isTargetWorkflowSelected = wsPageUtility.GetFieldValue("jmcg_targetworkflow") != null;
    var isViewNotificationEntityTypeEntered = wsPageUtility.GetFieldValue("jmcg_viewnotificationentitytype") != null;
    var isViewNotificationToQueue = wsPageUtility.GetFieldValue("jmcg_viewnotificationoption") == wsWorkflowTasks.Options.ViewNotificationOption.EmailQueue;

    var displayViewSelectForViewTarget = isViewTarget && isTargetWorkflowSelected;
    var displayViewSelectForViewNotifications = isViewNotification && isViewNotificationEntityTypeEntered;

    wsPageUtility.SetSectionVisibility("secViewNotification", isViewNotification);
    wsPageUtility.SetSectionVisibility("secFetchTarget", isFetchTarget);
    wsPageUtility.SetSectionVisibility("secViewTarget", isViewTarget);
    wsPageUtility.SetSectionVisibility("secViewSelection", displayViewSelectForViewTarget || displayViewSelectForViewNotifications);

    wsPageUtility.SetFieldVisibility("jmcg_sendnotificationfortargetfailures", typeSelected && !isViewNotification);
    wsPageUtility.SetFieldVisibility("jmcg_targetworkflow", typeSelected && !isViewNotification);
    wsPageUtility.SetFieldVisibility("jmcg_viewnotificationqueue", isViewNotificationToQueue);
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
    //todo could do something while loading/changing the view selections

    var processViewResults = function (views) {
        wsWorkflowTasks.ViewSelections = views;
        var viewoptions = new Array();
        viewoptions.push(new wsPageUtility.PicklistOption(0, "Select to change the selected view"));
        var currentViewInOptions = false;
        for (var i = 1; i <= wsWorkflowTasks.ViewSelections.length; i++) {
            viewoptions.push(new wsPageUtility.PicklistOption(i, wsWorkflowTasks.ViewSelections[i - 1]["name"]));
            if (wsWorkflowTasks.ViewSelections[i - 1]["savedqueryid"] == wsPageUtility.GetFieldValue("jmcg_targetviewid"))
                currentViewInOptions = true;
        }
        wsPageUtility.SetPicklistOptions("jmcg_targetviewselectionfield", viewoptions);
        wsPageUtility.SetFieldValue("jmcg_targetviewselectionfield", 0);
        if (!currentViewInOptions) {
            wsPageUtility.SetFieldValue("jmcg_targetviewid", null);
            wsPageUtility.SetFieldValue("jmcg_targetviewselectedname", null);
        }
    };

    var loadViewsForEntityType = function (entityType) {
        var conditions = [new wsServiceUtility.FilterCondition("returnedtypecode", wsServiceUtility.FilterOperator.Equal, entityType)];
        var orders = [new wsServiceUtility.OrderCondition("name", false)];
        wsServiceUtility.RetrieveMultipleAsync("savedquery", ["savedqueryid", "name"], conditions, orders, processViewResults);
    };

    var processTargetWorkflow = function (workflow) {
        var targetType = workflow["primaryentity"];
        loadViewsForEntityType(targetType);
    };

    var isViewTarget = wsPageUtility.GetFieldValue("jmcg_workflowexecutiontype") == wsWorkflowTasks.Options.WorkflowExecutionType.TargetPerViewResult;
    var isViewNotification = wsPageUtility.GetFieldValue("jmcg_workflowexecutiontype") == wsWorkflowTasks.Options.WorkflowExecutionType.ViewNotification;
    var isTargetWorkflowSelected = wsPageUtility.GetFieldValue("jmcg_targetworkflow") != null;
    var isViewNotificationEntityTypeEntered = wsPageUtility.GetFieldValue("jmcg_viewnotificationentitytype") != null;

    var displayViewSelectForViewTarget = isViewTarget && isTargetWorkflowSelected;
    var displayViewSelectForViewNotifications = isViewNotification && isViewNotificationEntityTypeEntered;

    if (displayViewSelectForViewTarget || displayViewSelectForViewNotifications)
    {
        if (isViewTarget) {
            var targetWorkflowId = wsPageUtility.GetLookupId("jmcg_targetworkflow");
            wsServiceUtility.RetrieveAsync("workflow", targetWorkflowId, ["primaryentity"], processTargetWorkflow);
        }
        else {
            loadViewsForEntityType(wsPageUtility.GetFieldValue("jmcg_viewnotificationentitytype"));
        }
    }
    else {
        wsWorkflowTasks.ViewSelections = [];
        wsPageUtility.SetFieldValue("jmcg_targetviewid", null);
        wsPageUtility.SetFieldValue("jmcg_targetviewselectedname", null);
    }
}

wsWorkflowTasks.PopulateUrlField = function () {
    var processResults = function (results) {
        if (results.length > 0) {
            wsPageUtility.SetFieldValue("jmcg_crmbaseurl", results[0]["jmcg_crmbaseurl"]);
        }
    }
    if (wsPageUtility.GetFormType() == wsPageUtility.FormMode.Create) {
        var condition = new wsServiceUtility.FilterCondition("jmcg_crmbaseurl", wsServiceUtility.FilterOperator.NotNull);
        wsServiceUtility.RetrieveMultipleAsync("jmcg_workflowtask", "jmcg_crmbaseurl", [condition], null, processResults);
    }
}