using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class PropertyDecoratorInfo
    {
        public int ColIndex { get; set; }
        public List<BaseDecorateAttribute> DecoratorAttrs { get; set; }
    }
}
