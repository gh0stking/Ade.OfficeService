using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public static class FilterFactory
    {
        private static readonly Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));
        public static IFilter CreateInstance(Type attrType)
        {
            if (Table[attrType] != null)
            {
                return (IFilter)Table[attrType];
            }

            IFilter filter = null;

            Type filterType = Assembly.GetAssembly(attrType).GetTypes().ToList()?.
                 Where(t => typeof(IFilter).IsAssignableFrom(t))?.
                 FirstOrDefault(t => t.IsDefined(typeof(FilterBindAttribute))
                 && t.GetCustomAttribute<FilterBindAttribute>()?.FilterAttributeType == attrType);

            if (filterType != null)
            {
                filter = (IFilter)Activator.CreateInstance(filterType);
                Table[attrType] = filter;
            }

            return filter;
        }
    }
}
