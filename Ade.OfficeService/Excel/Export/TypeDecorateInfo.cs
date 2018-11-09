using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class TypeDecorateInfo
    {
        public List<IDecorator> GlobalDecorators { get; set; }
       public List<PropertyDecorateInfo> PropertyExportInfos { get; set; }
    }
}
