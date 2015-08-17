using System;
using System.Linq;

namespace JosephM.Xrm.WorkflowScheduler.Core
{
    /// <summary>
    ///     Defines The Property With The Attribute In Context If The Property With A Given Name Has One Of The Given Values
    /// </summary>
    [AttributeUsage(
        AttributeTargets.Property,
        AllowMultiple = false)]
    public class PropertyInContextByPropertyValues : PropertyInContext
    {
        public string PropertyDependency { get; set; }
        public object[] ValidValues { get; set; }

        public PropertyInContextByPropertyValues(string propertyDependency, object[] validValues)
        {
            PropertyDependency = propertyDependency;
            ValidValues = validValues;
        }

        public override bool IsInContext(object instance)
        {
            var propertyValue = instance.GetPropertyValue(PropertyDependency);
            return
                propertyValue != null && ValidValues.Contains(instance.GetPropertyValue(PropertyDependency));
        }
    }
}