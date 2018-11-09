using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Ade.OfficeService.Excel
{
    public class PropertyDecorateInfo
    {
        public string ColName { get; set; }
        public List<IDecorator> Decorators { get; set; }
    }
}
