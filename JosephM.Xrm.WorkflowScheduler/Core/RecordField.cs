using System.Runtime.Serialization;

namespace JosephM.Xrm.WorkflowScheduler.Core
{
    [DataContract]
    public class RecordField : PicklistOption
    {
        public RecordField()
        {
        }

        public RecordField(string key, string value)
            : base(key, value)
        {
        }
    }
}