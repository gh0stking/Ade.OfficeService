using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class PropertyFilterInfo
    {
        public int ColIndex { get; set; }
        public List<BaseFilterAttribute> FilterAttrs { get; set; }
    }
}
