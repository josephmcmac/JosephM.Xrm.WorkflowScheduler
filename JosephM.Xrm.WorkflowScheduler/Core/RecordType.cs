using System.Runtime.Serialization;

namespace JosephM.Xrm.WorkflowScheduler.Core
{
    [DataContract]
    public class RecordType : PicklistOption
    {
        public RecordType()
        {
        }

        public RecordType(string key, string value)
            : base(key, value)
        {
        }
    }
}