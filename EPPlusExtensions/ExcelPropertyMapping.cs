using System;
using System.Collections.Generic;
using System.Reflection;

namespace EPPlusExtensions
{
    public class ExcelPropertyMapping
    {
        public static Dictionary<Type, string> DefaultFormatters { get; } = new Dictionary<Type, string>
        {
            {typeof(DateTime), "dd/mm/yyyy"}
        };
        
        public ExcelPropertyMapping(PropertyInfo runtimeProperty, Func<object, object> transformValue,
                                    string header, int order = -1)
        {
            RuntimeProperty = runtimeProperty;
            TransformValue = transformValue;
            Header = header;
            Order = order;
            Format = DefaultFormatters.GetValueOrDefault(Nullable.GetUnderlyingType(runtimeProperty.PropertyType) ?? runtimeProperty.PropertyType);
        }

        public PropertyInfo RuntimeProperty { get; }

        public Func<object, object> TransformValue { get; set; }

        public string Header { get; set; }

        public string Format { get; set; }

        public int Order { get; set; }

        internal object GetValue(object item)
        {
            var propertyValue = RuntimeProperty?.GetValue(item);
            var transformedValue = TransformValue?.Invoke(propertyValue) ?? propertyValue;
            return transformedValue;
        }
    }
}