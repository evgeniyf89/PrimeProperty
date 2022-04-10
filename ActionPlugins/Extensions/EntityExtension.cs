using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.Extensions
{
    public static class EntityExtension
    {
        public static T GetValue<T>(this Entity entity, string key)
        {
            if (!entity.Contains(key))
                return default;
            var value = entity[key];
            return value is AliasedValue alias
                ? (T)(alias.Value)
                : (T)value;
        }
    }
}
