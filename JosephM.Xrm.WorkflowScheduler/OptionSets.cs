namespace JosephM.Xrm.WorkflowScheduler
{
    public static class OptionSets
    {
        public static class Process
        {
            public static class Deletestage
            {
                public const int Preoperation = 20;
                public const int Postoperation = 40;
            }
            public static class RunAsUser
            {
                public const int Owner = 0;
                public const int CallingUser = 1;
            }
            public static class Mode
            {
                public const int Background = 0;
                public const int Realtime = 1;
            }
            public static class UpdateStage
            {
                public const int Preoperation = 20;
                public const int Postoperation = 40;
            }
            public static class Type
            {
                public const int Definition = 1;
                public const int Activation = 2;
                public const int Template = 3;
            }
            public static class Scope
            {
                public const int User = 1;
                public const int BusinessUnit = 2;
                public const int ParentChildBusinessUnits = 3;
                public const int Organization = 4;
            }
            public static class StatusReason
            {
                public const int Draft = 1;
                public const int Activated = 2;
            }
            public static class ComponentState
            {
                public const int Published = 0;
                public const int Unpublished = 1;
                public const int Deleted = 2;
                public const int DeletedUnpublished = 3;
            }
            public static class Category
            {
                public const int Workflow = 0;
                public const int Dialog = 1;
                public const int PBL = 2;
                public const int Action = 3;
                public const int BusinessProcessFlow = 4;
            }
            public static class CreateStage
            {
                public const int Preoperation = 20;
                public const int Postoperation = 40;
            }
        }
        public static class SystemJob
        {
            public static class SystemJobType
            {
                public const int SystemEvent = 1;
                public const int BulkEmail = 2;
                public const int ImportFileParse = 3;
                public const int TransformParseData = 4;
                public const int Import = 5;
                public const int ActivityPropagation = 6;
                public const int DuplicateDetectionRulePublish = 7;
                public const int BulkDuplicateDetection = 8;
                public const int SQMDataCollection = 9;
                public const int Workflow = 10;
                public const int QuickCampaign = 11;
                public const int MatchcodeUpdate = 12;
                public const int BulkDelete = 13;
                public const int DeletionService = 14;
                public const int IndexManagement = 15;
                public const int CollectOrganizationStatistics = 16;
                public const int ImportSubprocess = 17;
                public const int CalculateOrganizationStorageSize = 18;
                public const int CollectOrganizationDatabaseStatistics = 19;
                public const int CollectionOrganizationSizeStatistics = 20;
                public const int DatabaseTuning = 21;
                public const int CalculateOrganizationMaximumStorageSize = 22;
                public const int BulkDeleteSubprocess = 23;
                public const int UpdateStatisticIntervals = 24;
                public const int OrganizationFullTextCatalogIndex = 25;
                public const int Databaselogbackup = 26;
                public const int UpdateContractStates = 27;
                public const int DBCCSHRINKDATABASEmaintenancejob = 28;
                public const int DBCCSHRINKFILEmaintenancejob = 29;
                public const int Reindexallindicesmaintenancejob = 30;
                public const int StorageLimitNotification = 31;
                public const int Cleanupinactiveworkflowassemblies = 32;
                public const int RecurringSeriesExpansion = 35;
                public const int ImportSampleData = 38;
                public const int GoalRollUp = 40;
                public const int AuditPartitionCreation = 41;
                public const int CheckForLanguagePackUpdates = 42;
                public const int ProvisionLanguagePack = 43;
                public const int UpdateOrganizationDatabase = 44;
                public const int UpdateSolution = 45;
                public const int RegenerateEntityRowCountSnapshotData = 46;
                public const int RegenerateReadShareSnapshotData = 47;
                public const int OutgoingActivity = 50;
                public const int IncomingEmailProcessing = 51;
                public const int MailboxTestAccess = 52;
                public const int EncryptionHealthCheck = 53;
                public const int ExecuteAsyncRequest = 54;
                public const int PosttoYammer = 49;
            }
            public static class StatusReason
            {
                public const int WaitingForResources = 0;
                public const int Waiting = 10;
                public const int InProgress = 20;
                public const int Pausing = 21;
                public const int Canceling = 22;
                public const int Succeeded = 30;
                public const int Failed = 31;
                public const int Canceled = 32;
            }
        }
        public static class WorkflowTask
        {
            public static class StatusReason
            {
                public const int Active = 1;
                public const int Inactive = 2;
            }
            public static class PeriodPerRunUnit
            {
                public const int Minutes = 1;
                public const int Hours = 2;
                public const int Days = 3;
                public const int Months = 4;
            }
            public static class WorkflowExecutionType
            {
                public const int TargetThisWorkflowTask = 1;
                public const int TargetPerFetchResult = 2;
            }
        }
    }
}
