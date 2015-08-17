using System;

namespace JosephM.Xrm.WorkflowScheduler.Core
{
    /// <summary>
    ///     Base Class For Logic To Define A Property As Having A Valid State
    /// </summary>
    [AttributeUsage(
        AttributeTargets.Property,
        AllowMultiple = true)]
    public abstract class PropertyValidator : Attribute
    {
        public abstract bool IsValid(object value, object instance);

        public abstract string GetErrorMessage(string propertyName, object instance);
    }
}