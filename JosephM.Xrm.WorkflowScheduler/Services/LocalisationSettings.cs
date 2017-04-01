namespace JosephM.Xrm.WorkflowScheduler.Services
{
    public class LocalisationSettings
    {
        public LocalisationSettings(string targetTimeZoneId)
        {
            TargetTimeZoneId = targetTimeZoneId;
        }

        public string TargetTimeZoneId
        {
            get; private set;
        }
    }
}