using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    [AttributeUsage( AttributeTargets.Class,AllowMultiple =false,Inherited =false)]
    public class FilterBindOptionAttribute : Attribute
    {
        public FilterBindOptionAttribute(Type filterAttributeType)
        {
            if (!filterAttributeType.IsSubclassOf(typeof(BaseFilterAttribute)))
            {
                throw new ArgumentOutOfRangeException(filterAttributeType.ToString()+ "is not subclass of BaseFilterAttribute");
            }

            this.FilterAttributeType = filterAttributeType;
        }
        public Type FilterAttributeType { get; set; }
    }
}
