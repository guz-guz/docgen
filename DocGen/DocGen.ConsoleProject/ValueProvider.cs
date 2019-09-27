using System;
using System.Collections.Generic;
using System.Reflection;
using DocumentFormat.OpenXml.Bibliography;

namespace Test
{
    class ValueProvider
    {
        private readonly Type _modelType;
        private readonly string _prefix;

        public ValueProvider(Type modelType, string prefix = null)
        {
            _modelType = modelType;
            _prefix = prefix;
        }
        
        public object Model { get; set; }

        public string Prefix => _prefix;
        
        public object GetValue(string fieldPath)
        {
            var fieldName = _prefix != null ? fieldPath.Replace(_prefix, string.Empty) : fieldPath;
            var propertyInfo = _modelType.GetProperty(fieldName, BindingFlags.Instance | BindingFlags.Public)
                               ?? throw new KeyNotFoundException($"There is no property {fieldPath}");
            
            return propertyInfo.GetValue(Model);
        }
    }
}