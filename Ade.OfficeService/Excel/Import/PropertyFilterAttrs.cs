using System;
using System.Collections.Generic;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class PropertyFilterAttrs
    {
        public int ColIndex { get; set; }
        public string PropertyName { get; set; }
        public List<BaseFilterAttribute> FilterAttrs { get; set; }
    }
}
