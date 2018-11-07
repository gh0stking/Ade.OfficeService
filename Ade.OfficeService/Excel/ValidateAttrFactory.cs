using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public static class ValidateAttrFactory
    {
        private static readonly Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));
        public static BaseValidateAttribute Create(Type attrType,Type importDTOType, string excelColName)
        {
            var key = new ValidateAttributeKey { ImportDTOType = importDTOType, ExcelColName = excelColName };
            if (Table[key] != null)
            {
                return (BaseValidateAttribute)Table[key];
            }

            var props = importDTOType.GetProperties();
            PropertyInfo propMatch = null;
            foreach (var prop in props)
            {
                if (prop.IsDefined(typeof(ColNameAttribute)) 
                    && prop.GetCustomAttribute<ColNameAttribute>().ColName == excelColName)
                {
                    propMatch = prop;
                    break;
                }
            }
            
            if (propMatch != null)
            {
                var validateAttr = (BaseValidateAttribute)propMatch.GetCustomAttribute(attrType);
                Table[key] = validateAttr;
                return validateAttr;
            }

            return null;
        }
    }
}
