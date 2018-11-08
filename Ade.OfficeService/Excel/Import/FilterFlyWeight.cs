using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class FilterFlyWeight
    {
        private FilterFlyWeight()
        {
        }
        private static readonly Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));
        public static IFilter CreateInstance(BaseFilterAttribute attr)
        {
            if (Table[attr] != null)
            {
                return (IFilter)Table[attr];
            }

            IFilter filter = null;

            Type filterType = Assembly.GetAssembly(typeof(IFilter)).GetTypes().ToList()?.
                 Where(t => typeof(IFilter).IsAssignableFrom(t))?.
                 FirstOrDefault(t => t.IsDefined(typeof(FilterBindOptionAttribute))
                 && t.GetCustomAttribute<FilterBindOptionAttribute>()?.FilterAttributeType == attr.GetType());

            if (filterType != null)
            {
                filter = (IFilter)Activator.CreateInstance(filterType);
                Table[attr] = filter;
            }

            return filter;
        }

        public List<PropertyFilterAttrs> PropertyFilterAttrs { get; set; } = new List<PropertyFilterAttrs>();
        
    }
}
