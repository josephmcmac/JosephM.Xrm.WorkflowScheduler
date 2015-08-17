#region

using System;
using System.Collections.Generic;

#endregion

namespace JosephM.Xrm.WorkflowScheduler.Core
{
    public class ServiceResponseBase<TResponseItem>
        where TResponseItem : ServiceResponseItem
    {
        private readonly List<TResponseItem> _errors = new List<TResponseItem>();

        public ServiceResponseBase()
        {
            Success = true;
        }

        public void SetFatalError(Exception ex)
        {
            Success = false;
            Exception = ex;
        }

        public bool Success { get; private set; }
        public Exception Exception { get; set; }

        public void AddResponseItem(TResponseItem responseItem)
        {
            _errors.Add(responseItem);
        }

        public void AddResponseItems(IEnumerable<TResponseItem> responseItems)
        {
            _errors.AddRange(responseItems);
        }

        public IEnumerable<TResponseItem> ResponseItems
        {
            get { return _errors; }
        }
    }
}