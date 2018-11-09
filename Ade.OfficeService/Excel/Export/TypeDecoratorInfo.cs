using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class TypeDecoratorInfo
    {
        public List<BaseDecorateAttribute> TypeDecoratorAttrs { get; set; }
        public List<PropertyDecoratorInfo> PropertyDecoratorInfos { get; set; }
    }
}
